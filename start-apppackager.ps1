<#
.SYNOPSIS
    MahApps.Metro WPF front-end for application packager scripts (metadata-driven, no network on launch).

.DESCRIPTION
    MahApps.Metro 2.4.10 WPF front-end for application packager scripts.
    Modern sidebar layout with dark/light theme toggle.

    On launch, the tool performs LOCAL-ONLY operations:
      - Enumerates packager scripts in the PackagersRoot folder
      - Parses metadata tags from each script header
      - Populates the grid with Vendor/Application and placeholders

    No network operations are performed on launch.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER ProviderMachineName
    ConfigMgr SMS Provider machine name from the AdminUI connect script.

.PARAMETER PackagersRoot
    Local folder containing packager scripts (e.g., .\Packagers).

.EXAMPLE
    .\start-apppackager.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8.2
      - MahApps.Metro 2.4.10 DLLs in .\Lib\
      - 7-Zip (required by Tableau packagers)
      - Local administrator (required by some packagers)

    ScriptName : start-apppackager.ps1
    Purpose    : MahApps WPF front-end for packager scripts
    Owner      : CM Engineering
    Version    : 1.4.0.11
    Updated    : 2026-08-17
#>

param(
    [string]$SiteCode = "MCM",
    [string]$ProviderMachineName = "",
    [string]$PackagersRoot = (Join-Path $PSScriptRoot "Packagers"),

    # Headless batch mode: skips the WPF shell, runs a currency check across
    # the packagers in -Apps, takes action per -OnUpdateFound. CLI-driven.
    [switch]$BatchMode,
    [string[]]$Apps,
    [ValidateSet('Report','Stage','StageAndPackage')][string]$OnUpdateFound = 'Report',
    [string]$LogPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Assembly loading (must happen before XAML parse)
# =============================================================================
# AppPackagerCommon is imported unconditionally -- both modes need its helpers.
Import-Module (Join-Path $PSScriptRoot 'Packagers\AppPackagerCommon.psm1') -Force -DisableNameChecking -ErrorAction SilentlyContinue

if (-not $BatchMode) {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

    $libDir = Join-Path $PSScriptRoot 'Lib'

    # Auto-unblock: if the tree was copied from a remote share (Copy-Item
    # -ToSession, browser download, etc.) Windows stamps MOTW on every file,
    # which makes LoadFrom fail with a misleading "cannot find file specified"
    # error. Silently strip MOTW from everything in Lib\ before loading.
    Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue

    [System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll')) | Out-Null
    [System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll')) | Out-Null
    [System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll')) | Out-Null
}

# =============================================================================
# Helpers (carried over from WinForms version)
# =============================================================================
function Get-PreferencesPath {
    Join-Path $PSScriptRoot "AppPackager.preferences.json"
}

function Read-Preferences {
    $defaults = [pscustomobject]@{
        SiteCode             = "MCM"
        ProviderMachineName  = ""
        FileShareRoot        = "\\fileserver\sccm$"
        ContentLayout        = "Nested"
        DownloadRoot         = "C:\temp\ap"
        EstimatedRuntimeMins = 15
        MaximumRuntimeMins   = 30
        CompanyName          = ""
        M365Channel          = "MonthlyEnterprise"
        M365DeployMode       = "Managed"
        M365ExcludeApps      = @('Groove','Lync','OneDrive','Teams','Bing')
        SSMSInstallOptions   = [pscustomobject]@{
            UIMode             = "Quiet"
            DownloadThenInstall = $true
            NoUpdateInstaller  = $false
            IncludeRecommended = $false
            IncludeOptional    = $false
            RemoveOos          = $true
            ForceClose         = $false
            InstallPath        = ""
        }
        HiddenApplications   = @()
        AppFlow              = [pscustomobject]@{
            Tracked          = @()
            Action           = 'Report'
            CadenceOverrides = [pscustomobject]@{}
            ForceOnLaunch    = $false
        }
        DetectedTools        = [pscustomobject]@{
            ConfigMgrConsole = [pscustomobject]@{
                Found           = $false
                DisplayName     = ''
                DisplayVersion  = ''
                InstallLocation = ''
                ModulePath      = ''
                DetectedAt      = ''
            }
            SevenZipCli      = [pscustomobject]@{
                Found           = $false
                DisplayName     = ''
                DisplayVersion  = ''
                InstallLocation = ''
                ExePath         = ''
                DetectedAt      = ''
            }
            IntuneWinAppUtil = [pscustomobject]@{
                Found          = $false
                DisplayVersion = ''
                ExePath        = ''
                DetectedAt     = ''
            }
        }
        ContentDistribution  = [pscustomobject]@{
            AutoDistribute                = $false
            DPGroupName                   = ''
            DeployToTestCollection        = $false
            TestCollectionName            = ''
            CreateTestCollectionIfMissing = $false
        }
        Intune               = [pscustomobject]@{
            CreateIntuneWin = $false
        }
        DeploymentConditions = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
        CommandOverrides = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
    }

    $path = Get-PreferencesPath
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($null -ne $data.SiteCode)             { $defaults.SiteCode             = [string]$data.SiteCode }
        if ($null -ne $data.ProviderMachineName)  { $defaults.ProviderMachineName  = [string]$data.ProviderMachineName }
        elseif ($data.MECM -and $null -ne $data.MECM.ServerFQDN) {
            $defaults.ProviderMachineName = [string]$data.MECM.ServerFQDN
        }
        if ($null -ne $data.FileShareRoot)         { $defaults.FileShareRoot        = [string]$data.FileShareRoot }
        if ([string]$data.ContentLayout -in @('Nested','Flat')) { $defaults.ContentLayout = [string]$data.ContentLayout }
        if ($null -ne $data.DownloadRoot)          { $defaults.DownloadRoot         = [string]$data.DownloadRoot }
        if ($null -ne $data.EstimatedRuntimeMins)  { $defaults.EstimatedRuntimeMins = [int]$data.EstimatedRuntimeMins }
        if ($null -ne $data.MaximumRuntimeMins)    { $defaults.MaximumRuntimeMins   = [int]$data.MaximumRuntimeMins }
        if ($null -ne $data.CompanyName)            { $defaults.CompanyName          = [string]$data.CompanyName }

        # M365Channel: validate against current set; migrate legacy SemiAnnual
        # and SemiAnnualPreview to MonthlyEnterprise (SAEC retired from the UI).
        # Unknown values fall back to the default rather than trip the
        # packager's [ValidateSet] and fail staging.
        if ($null -ne $data.M365Channel) {
            $chanRaw = [string]$data.M365Channel
            switch -Regex ($chanRaw) {
                '^(MonthlyEnterprise|Current)$' { $defaults.M365Channel = $chanRaw }
                '^SemiAnnual(Preview)?$'        { $defaults.M365Channel = 'MonthlyEnterprise' }
                default                         { $defaults.M365Channel = 'MonthlyEnterprise' }
            }
        }

        # M365DeployMode: same guard against unknown values.
        if ($null -ne $data.M365DeployMode) {
            $modeRaw = [string]$data.M365DeployMode
            if ($modeRaw -in @('Managed','Online')) { $defaults.M365DeployMode = $modeRaw }
            else { $defaults.M365DeployMode = 'Managed' }
        }

        if ($null -ne $data.M365ExcludeApps) {
            # Filter to only documented ExcludeApp IDs (plus "Bing" which is
            # accepted historically). Unknown values are dropped silently.
            $validExcludes = @('Access','Excel','Groove','Lync','OneDrive','OneNote','Outlook','OutlookForWindows','PowerPoint','Publisher','Teams','Word','Bing')
            $defaults.M365ExcludeApps = @($data.M365ExcludeApps | Where-Object { $_ -in $validExcludes })
        }

        if ($null -ne $data.SSMSInstallOptions) {
            $ssms = $data.SSMSInstallOptions
            if ($null -ne $ssms.UIMode) {
                $modeRaw = [string]$ssms.UIMode
                if ($modeRaw -in @('Quiet','Passive')) { $defaults.SSMSInstallOptions.UIMode = $modeRaw }
            }
            foreach ($prop in @('DownloadThenInstall','NoUpdateInstaller','IncludeRecommended','IncludeOptional','RemoveOos','ForceClose')) {
                if ($null -ne $ssms.$prop) {
                    try { $defaults.SSMSInstallOptions.$prop = [bool]$ssms.$prop } catch { }
                }
            }
            if ($null -ne $ssms.InstallPath) { $defaults.SSMSInstallOptions.InstallPath = [string]$ssms.InstallPath }
        }

        if ($null -ne $data.HiddenApplications)    { $defaults.HiddenApplications  = @($data.HiddenApplications) }

        # AppFlow: 1-click Full Run settings. Schema is additive; missing key
        # keeps the defaults above so older prefs files from v1.0 still load.
        if ($null -ne $data.AppFlow) {
            $af = $data.AppFlow

            if ($null -ne $af.Tracked) {
                $defaults.AppFlow.Tracked = @(
                    $af.Tracked |
                        Where-Object { $_ -is [string] -and $_ -match '^package-' } |
                        ForEach-Object { [string]$_ }
                )
            }

            if ($null -ne $af.Action) {
                $actionRaw = [string]$af.Action
                if ($actionRaw -in @('Report','Stage','StageAndPackage')) {
                    $defaults.AppFlow.Action = $actionRaw
                }
            }

            if ($null -ne $af.CadenceOverrides) {
                $overrideProps = [ordered]@{}
                foreach ($prop in $af.CadenceOverrides.PSObject.Properties) {
                    if ($prop.Name -notmatch '^package-') { continue }
                    $days = 0
                    if ([int]::TryParse([string]$prop.Value, [ref]$days) -and $days -ge 1) {
                        $overrideProps[$prop.Name] = $days
                    }
                }
                $defaults.AppFlow.CadenceOverrides = [pscustomobject]$overrideProps
            }

            if ($null -ne $af.ForceOnLaunch) {
                try { $defaults.AppFlow.ForceOnLaunch = [bool]$af.ForceOnLaunch } catch { }
            }
        }

        # ContentDistribution: auto-distribute-to-DP-group settings.
        if ($null -ne $data.ContentDistribution) {
            $cd = $data.ContentDistribution
            if ($null -ne $cd.AutoDistribute) {
                try { $defaults.ContentDistribution.AutoDistribute = [bool]$cd.AutoDistribute } catch { }
            }
            if ($null -ne $cd.DPGroupName) {
                $defaults.ContentDistribution.DPGroupName = [string]$cd.DPGroupName
            }
            if ($null -ne $cd.DeployToTestCollection) {
                try { $defaults.ContentDistribution.DeployToTestCollection = [bool]$cd.DeployToTestCollection } catch { }
            }
            if ($null -ne $cd.TestCollectionName) {
                $defaults.ContentDistribution.TestCollectionName = [string]$cd.TestCollectionName
            }
            if ($null -ne $cd.CreateTestCollectionIfMissing) {
                try { $defaults.ContentDistribution.CreateTestCollectionIfMissing = [bool]$cd.CreateTestCollectionIfMissing } catch { }
            }
        }

        # Intune: .intunewin production during Package.
        if ($null -ne $data.Intune -and $null -ne $data.Intune.CreateIntuneWin) {
            try { $defaults.Intune.CreateIntuneWin = [bool]$data.Intune.CreateIntuneWin } catch { }
        }

        # DeploymentConditions: per-app requirement rule selections applied
        # at Package time. Entries that reduce to no conditions are dropped;
        # invalid values fall back silently like the other sections.
        if ($null -ne $data.DeploymentConditions -and $null -ne $data.DeploymentConditions.Apps) {
            $condProps = [ordered]@{}
            foreach ($prop in $data.DeploymentConditions.Apps.PSObject.Properties) {
                if ($prop.Name -notmatch '^package-') { continue }
                $entry = $prop.Value
                $arch = 'Any'
                if ([string]$entry.Architecture -in @('x64', 'ARM64')) { $arch = [string]$entry.Architecture }
                $network = 'Any'
                if ([string]$entry.Network -in @('VpnOnly', 'OnSiteOnly')) { $network = [string]$entry.Network }
                $langs = @()
                if ($null -ne $entry.Languages) {
                    $langs = @($entry.Languages |
                        Where-Object { [string]$_ -match '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,2}$' } |
                        ForEach-Object { [string]$_ })
                }
                $split = 'None'
                if ([string]$entry.Split -in @('Architecture', 'Language', 'Network')) { $split = [string]$entry.Split }
                if ($arch -eq 'Any' -and $network -eq 'Any' -and $langs.Count -eq 0 -and $split -eq 'None') { continue }
                $condProps[$prop.Name] = [pscustomobject]@{
                    Architecture = $arch
                    Languages    = $langs
                    Network      = $network
                    Split        = $split
                }
            }
            $defaults.DeploymentConditions.Apps = [pscustomobject]$condProps
        }

        # CommandOverrides: per-app install/uninstall command replacements.
        # Entries with neither command are dropped; whitespace trims away.
        if ($null -ne $data.CommandOverrides -and $null -ne $data.CommandOverrides.Apps) {
            $cmdProps = [ordered]@{}
            foreach ($prop in $data.CommandOverrides.Apps.PSObject.Properties) {
                if ($prop.Name -notmatch '^package-') { continue }
                $entry = $prop.Value
                $inst = ([string]$entry.Install).Trim()
                $uninst = ([string]$entry.Uninstall).Trim()
                if (-not $inst -and -not $uninst) { continue }
                $cmdProps[$prop.Name] = [pscustomobject]@{
                    Install   = $inst
                    Uninstall = $uninst
                }
            }
            $defaults.CommandOverrides.Apps = [pscustomobject]$cmdProps
        }

        # DetectedTools: last known detection results. Refreshed on launch
        # but persists across sessions so we have something to show before
        # the first detection completes.
        if ($null -ne $data.DetectedTools -and $null -ne $data.DetectedTools.ConfigMgrConsole) {
            $cm = $data.DetectedTools.ConfigMgrConsole
            $stored = [pscustomobject]@{
                Found           = $false
                DisplayName     = ''
                DisplayVersion  = ''
                InstallLocation = ''
                ModulePath      = ''
                DetectedAt      = ''
            }
            if ($null -ne $cm.Found)           { try { $stored.Found = [bool]$cm.Found } catch { } }
            if ($null -ne $cm.DisplayName)     { $stored.DisplayName     = [string]$cm.DisplayName }
            if ($null -ne $cm.DisplayVersion)  { $stored.DisplayVersion  = [string]$cm.DisplayVersion }
            if ($null -ne $cm.InstallLocation) { $stored.InstallLocation = [string]$cm.InstallLocation }
            if ($null -ne $cm.ModulePath)      { $stored.ModulePath      = [string]$cm.ModulePath }
            if ($null -ne $cm.DetectedAt)      { $stored.DetectedAt      = [string]$cm.DetectedAt }
            $defaults.DetectedTools.ConfigMgrConsole = $stored
        }
        if ($null -ne $data.DetectedTools -and $null -ne $data.DetectedTools.SevenZipCli) {
            $sz = $data.DetectedTools.SevenZipCli
            $stored = [pscustomobject]@{
                Found           = $false
                DisplayName     = ''
                DisplayVersion  = ''
                InstallLocation = ''
                ExePath         = ''
                DetectedAt      = ''
            }
            if ($null -ne $sz.Found)           { try { $stored.Found = [bool]$sz.Found } catch { } }
            if ($null -ne $sz.DisplayName)     { $stored.DisplayName     = [string]$sz.DisplayName }
            if ($null -ne $sz.DisplayVersion)  { $stored.DisplayVersion  = [string]$sz.DisplayVersion }
            if ($null -ne $sz.InstallLocation) { $stored.InstallLocation = [string]$sz.InstallLocation }
            if ($null -ne $sz.ExePath)         { $stored.ExePath         = [string]$sz.ExePath }
            if ($null -ne $sz.DetectedAt)      { $stored.DetectedAt      = [string]$sz.DetectedAt }
            $defaults.DetectedTools.SevenZipCli = $stored
        }
        if ($null -ne $data.DetectedTools -and $null -ne $data.DetectedTools.IntuneWinAppUtil) {
            $iw = $data.DetectedTools.IntuneWinAppUtil
            $stored = [pscustomobject]@{
                Found          = $false
                DisplayVersion = ''
                ExePath        = ''
                DetectedAt     = ''
            }
            if ($null -ne $iw.Found)          { try { $stored.Found = [bool]$iw.Found } catch { } }
            if ($null -ne $iw.DisplayVersion) { $stored.DisplayVersion = [string]$iw.DisplayVersion }
            if ($null -ne $iw.ExePath)        { $stored.ExePath        = [string]$iw.ExePath }
            if ($null -ne $iw.DetectedAt)     { $stored.DetectedAt     = [string]$iw.DetectedAt }
            $defaults.DetectedTools.IntuneWinAppUtil = $stored
        }
    }
    catch { }

    return $defaults
}

function Save-Preferences {
    param([Parameter(Mandatory)][pscustomobject]$Prefs)

    $path = Get-PreferencesPath
    $json = $Prefs | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8

    $pkgPrefsPath = Join-Path (Join-Path $PSScriptRoot "Packagers") "packager-preferences.json"
    try {
        $pkgPrefs = @{}
        if (Test-Path -LiteralPath $pkgPrefsPath) {
            $existing = Get-Content -LiteralPath $pkgPrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($prop in $existing.PSObject.Properties) {
                $pkgPrefs[$prop.Name] = $prop.Value
            }
        }
        $pkgPrefs["CompanyName"]     = $Prefs.CompanyName
        $pkgPrefs["M365ExcludeApps"] = @($Prefs.M365ExcludeApps)
        $pkgPrefs["SSMSInstallOptions"] = $Prefs.SSMSInstallOptions
        $pkgPrefs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $pkgPrefsPath -Encoding UTF8
    }
    catch {
        Write-Warning ("Failed to save packager preferences ({0}): {1}" -f $pkgPrefsPath, $_.Exception.Message)
    }
}

function Invoke-DetectConfigMgrConsole {
    # Detects the ConfigMgr Console (AdminUI). Combines three signals:
    # 1. ARP registry: gets DisplayName + DisplayVersion (InstallLocation is
    #    often empty for this product, so the path alone isn't reliable).
    # 2. $env:SMS_ADMIN_UI_PATH: set by AdminUI install. Points at
    #    ...\AdminConsole\bin\i386; module lives at ...\AdminConsole\bin.
    # 3. Well-known install paths as a last resort.
    # Found = true only when ConfigurationManager.psd1 resolves on disk.
    $result = [pscustomobject]@{
        Found           = $false
        DisplayName     = ''
        DisplayVersion  = ''
        InstallLocation = ''
        ModulePath      = ''
        DetectedAt      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    # 1. ARP scan for display metadata
    $hives = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($hive in $hives) {
        if (-not (Test-Path $hive)) { continue }
        $matchEntry = Get-ChildItem -LiteralPath $hive -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop } catch { }
            } |
            Where-Object { $_.DisplayName -and $_.DisplayName -match 'Configuration Manager Console' } |
            Select-Object -First 1
        if ($matchEntry) {
            $result.DisplayName    = [string]$matchEntry.DisplayName
            $result.DisplayVersion = [string]$matchEntry.DisplayVersion
            if ($matchEntry.InstallLocation) {
                $result.InstallLocation = [string]$matchEntry.InstallLocation
            }
            break
        }
    }

    # 2. Build a list of candidate bin paths and resolve the module.
    $candidates = @()
    if ($env:SMS_ADMIN_UI_PATH) {
        # Env var points at ...\AdminConsole\bin\i386; parent is ...\AdminConsole\bin
        $candidates += (Split-Path -Parent $env:SMS_ADMIN_UI_PATH)
    }
    if ($result.InstallLocation) {
        $candidates += (Join-Path $result.InstallLocation 'bin')
    }
    $candidates += @(
        'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin',
        'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin',
        'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin',
        'C:\Program Files\Microsoft Endpoint Manager\AdminConsole\bin'
    )

    foreach ($c in ($candidates | Select-Object -Unique)) {
        if (-not $c) { continue }
        $mod = Join-Path $c 'ConfigurationManager.psd1'
        if (Test-Path -LiteralPath $mod) {
            $result.ModulePath = $mod
            if (-not $result.InstallLocation) {
                $result.InstallLocation = (Split-Path -Parent $c)
            }
            $result.Found = $true
            break
        }
    }

    return $result
}

function Invoke-DetectSevenZipCli {
    # Detects 7-Zip CLI (7z.exe). Used by package-adobereader.ps1 to extract
    # the Adobe enterprise installer and by package-teamviewerhost.ps1 to
    # read ProductVersion from an unsigned EXE's PE header. Supporting
    # non-default install paths (not just Program Files\7-Zip) makes the
    # tool work on workstations where an admin relocated it.
    # Detection signals, in order:
    #   1. ARP registry: DisplayName matches "7-Zip", InstallLocation points
    #      at the install dir.
    #   2. Well-known install paths (Program Files / Program Files x86).
    # Found = true only when 7z.exe resolves on disk.
    $result = [pscustomobject]@{
        Found           = $false
        DisplayName     = ''
        DisplayVersion  = ''
        InstallLocation = ''
        ExePath         = ''
        DetectedAt      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    # 1. ARP scan
    $hives = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($hive in $hives) {
        if (-not (Test-Path $hive)) { continue }
        $matchEntry = Get-ChildItem -LiteralPath $hive -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop } catch { }
            } |
            Where-Object { $_.DisplayName -and $_.DisplayName -match '^7-Zip' } |
            Select-Object -First 1
        if ($matchEntry) {
            $result.DisplayName    = [string]$matchEntry.DisplayName
            $result.DisplayVersion = [string]$matchEntry.DisplayVersion
            if ($matchEntry.InstallLocation) {
                $result.InstallLocation = [string]$matchEntry.InstallLocation
            }
            break
        }
    }

    # 2. Candidate paths: ARP InstallLocation wins; fall back to defaults.
    $candidates = @()
    if ($result.InstallLocation) {
        $candidates += $result.InstallLocation
    }
    $candidates += @(
        (Join-Path $env:ProgramFiles '7-Zip'),
        (Join-Path ${env:ProgramFiles(x86)} '7-Zip')
    )

    foreach ($c in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        $exe = Join-Path $c '7z.exe'
        if (Test-Path -LiteralPath $exe) {
            $result.ExePath = $exe
            if (-not $result.InstallLocation) {
                $result.InstallLocation = $c
            }
            $result.Found = $true
            break
        }
    }

    return $result
}

function Get-IntuneWinToolCachePath {
    return (Join-Path $env:LOCALAPPDATA 'AppPackager\Tools')
}

function Invoke-DetectIntuneWinAppUtil {
    # Detects the Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe).
    # The tool ships as a bare executable with no installer, so there is no
    # ARP entry to scan. Detection signals, in order:
    #   1. The path stored in preferences (keeps a manually placed copy).
    #   2. The AppPackager tool cache under LOCALAPPDATA, the download-on-
    #      first-use target.
    #   3. PATH via Get-Command.
    # Found = true only when IntuneWinAppUtil.exe resolves on disk.
    param([string]$KnownPath = '')

    $result = [pscustomobject]@{
        Found          = $false
        DisplayVersion = ''
        ExePath        = ''
        DetectedAt     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($KnownPath)) { $candidates += $KnownPath }
    $candidates += (Join-Path (Get-IntuneWinToolCachePath) 'IntuneWinAppUtil.exe')
    $cmd = Get-Command -Name 'IntuneWinAppUtil.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { $candidates += $cmd.Source }

    foreach ($c in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $c) {
            $result.ExePath = [string]$c
            try {
                $result.DisplayVersion = [string][System.Diagnostics.FileVersionInfo]::GetVersionInfo($c).FileVersion
            } catch { }
            $result.Found = $true
            break
        }
    }

    return $result
}

$script:Prefs = Read-Preferences

# Refresh tool detection once per launch. Persists into the same
# preferences JSON so the status is available immediately on next start.
try {
    $script:Prefs.DetectedTools.ConfigMgrConsole = Invoke-DetectConfigMgrConsole
    $script:Prefs.DetectedTools.SevenZipCli      = Invoke-DetectSevenZipCli
    $script:Prefs.DetectedTools.IntuneWinAppUtil = Invoke-DetectIntuneWinAppUtil -KnownPath ([string]$script:Prefs.DetectedTools.IntuneWinAppUtil.ExePath)
    Save-Preferences -Prefs $script:Prefs
} catch { }

if ([string]::IsNullOrWhiteSpace($script:Prefs.CompanyName)) {
    $pkgPrefsPath = Join-Path (Join-Path $PSScriptRoot "Packagers") "packager-preferences.json"
    if (Test-Path -LiteralPath $pkgPrefsPath) {
        try {
            $pkgData = Get-Content -LiteralPath $pkgPrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($pkgData.CompanyName) { $script:Prefs.CompanyName = [string]$pkgData.CompanyName }
        }
        catch { }
    }
}

if ($PSBoundParameters.ContainsKey('SiteCode')) {
    $script:Prefs.SiteCode = $SiteCode
}
if ($PSBoundParameters.ContainsKey('ProviderMachineName')) {
    $script:Prefs.ProviderMachineName = $ProviderMachineName
}

function Get-PackagerMetadata {
    param([Parameter(Mandatory)][string]$Path)

    $meta = [ordered]@{
        Vendor            = $null
        App               = $null
        CMName            = $null
        VendorUrl         = $null
        CPE               = $null
        ReleaseNotesUrl   = $null
        DownloadPageUrl   = $null
        Description       = $null
        UpdateCadenceDays = $null
        SupportsVariants  = @()
    }

    $lines = Get-Content -LiteralPath $Path -TotalCount 200 -ErrorAction Stop

    $inSynopsis = $false
    foreach ($line in $lines) {
        $l = $line.TrimStart([char]0xFEFF)

        if (-not $meta.Vendor    -and $l -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$')    { $meta.Vendor    = $Matches[1].Trim(); continue }
        if (-not $meta.App       -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')       { $meta.App       = $Matches[1].Trim(); continue }
        if (-not $meta.CMName    -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$')    { $meta.CMName    = $Matches[1].Trim(); continue }
        if (-not $meta.VendorUrl       -and $l -match '^\s*(?:#\s*)?VendorUrl\s*:\s*(.+?)\s*$')       { $meta.VendorUrl       = $Matches[1].Trim(); continue }
        if (-not $meta.CPE             -and $l -match '^\s*(?:#\s*)?CPE\s*:\s*(.+?)\s*$')             { $meta.CPE             = $Matches[1].Trim(); continue }
        if (-not $meta.ReleaseNotesUrl -and $l -match '^\s*(?:#\s*)?ReleaseNotesUrl\s*:\s*(.+?)\s*$') { $meta.ReleaseNotesUrl = $Matches[1].Trim(); continue }
        if (-not $meta.DownloadPageUrl -and $l -match '^\s*(?:#\s*)?DownloadPageUrl\s*:\s*(.+?)\s*$') { $meta.DownloadPageUrl = $Matches[1].Trim(); continue }
        if ($meta.SupportsVariants.Count -eq 0 -and $l -match '^\s*(?:#\s*)?SupportsVariants\s*:\s*(.+?)\s*$') {
            $meta.SupportsVariants = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -in @('Architecture', 'Language', 'Network') })
            continue
        }
        if ($null -eq $meta.UpdateCadenceDays -and $l -match '^\s*(?:#\s*)?UpdateCadenceDays\s*:\s*(\d+)\s*$') {
            $days = [int]$Matches[1]
            if ($days -ge 1) { $meta.UpdateCadenceDays = $days }
            continue
        }
        if (-not $meta.App             -and $l -match '^\s*(?:#\s*)?Application\s*:\s*(.+?)\s*$')     { $meta.App             = $Matches[1].Trim(); continue }

        if (-not $meta.Description -and $l -match '^\s*\.SYNOPSIS\s*$') { $inSynopsis = $true; continue }
        if ($inSynopsis -and -not $meta.Description) {
            $trimmed = $l.Trim()
            if ($trimmed.Length -gt 0) { $meta.Description = $trimmed; $inSynopsis = $false }
            continue
        }
    }

    if (-not $meta.CMName) { $meta.CMName = $meta.App }

    return [pscustomobject]@{
        Vendor            = $meta.Vendor
        Application       = $meta.App
        CMName            = $meta.CMName
        VendorUrl         = $meta.VendorUrl
        CPE               = $meta.CPE
        ReleaseNotesUrl   = $meta.ReleaseNotesUrl
        DownloadPageUrl   = $meta.DownloadPageUrl
        Description       = $meta.Description
        UpdateCadenceDays = $meta.UpdateCadenceDays
        SupportsVariants  = @($meta.SupportsVariants)
        Script            = (Split-Path -Leaf $Path)
        FullPath          = $Path
    }
}

function Get-Packagers {
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) { return @() }

    $files = Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop |
        Where-Object { $_.Name -match '^package-.*\.(?:ps1|notps1)$' } |
        Sort-Object Name

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        try {
            $m = Get-PackagerMetadata -Path $f.FullName

            $status = "Ready"
            if ($f.Extension -ieq ".notps1") { $status = "Not runnable (.notps1)" }
            if (-not $m.Vendor -or -not $m.Application) { $status = "Missing metadata (Vendor/App)" }

            $items.Add([pscustomobject]@{
                Selected          = $false
                Vendor            = $m.Vendor
                Application       = $m.Application
                CMName            = $m.CMName
                VendorUrl         = $m.VendorUrl
                Description       = $m.Description
                Script            = $m.Script
                FullPath          = $m.FullPath
                UpdateCadenceDays = $m.UpdateCadenceDays
                CurrentVersion    = ""
                LatestVersion     = ""
                Status            = $status
            })
        }
        catch {
            $items.Add([pscustomobject]@{
                Selected          = $false
                Vendor            = ""
                Application       = ""
                CMName            = ""
                VendorUrl         = ""
                Description       = ""
                Script            = $f.Name
                FullPath          = $f.FullName
                UpdateCadenceDays = $null
                CurrentVersion    = ""
                LatestVersion     = ""
                Status            = ("Read error: " + $_.Exception.Message)
            })
        }
    }
    return $items
}

function Test-PackagerSupportsFileServerPath {
    param([Parameter(Mandatory)][string]$PackagerPath)
    try {
        $head = Get-Content -LiteralPath $PackagerPath -TotalCount 120 -ErrorAction Stop | Out-String
        return ($head -match '\$FileServerPath')
    }
    catch { return $false }
}

function ConvertTo-ProcessArgument {
    param([AllowNull()][string]$Argument)

    if ($null -eq $Argument -or $Argument.Length -eq 0) { return '""' }
    if ($Argument -notmatch '[\s"]') { return $Argument }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    $backslashes = 0

    foreach ($ch in $Argument.ToCharArray()) {
        if ($ch -eq '\') {
            $backslashes++
            continue
        }
        if ($ch -eq '"') {
            if ($backslashes -gt 0) { [void]$sb.Append(('\' * ($backslashes * 2))) }
            [void]$sb.Append('\"')
            $backslashes = 0
            continue
        }
        if ($backslashes -gt 0) {
            [void]$sb.Append(('\' * $backslashes))
            $backslashes = 0
        }
        [void]$sb.Append($ch)
    }

    if ($backslashes -gt 0) { [void]$sb.Append(('\' * ($backslashes * 2))) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Set-ProcessStartInfoArgumentList {
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Arguments
    )

    $argumentListProperty = $StartInfo.GetType().GetProperty('ArgumentList')
    if ($argumentListProperty) {
        try { $StartInfo.ArgumentList.Clear() } catch { }
        foreach ($arg in $Arguments) {
            [void]$StartInfo.ArgumentList.Add($arg)
        }
    }
    else {
        $StartInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join ' ')
    }
}

function Get-PackagerFolderInfo {
    param([Parameter(Mandatory)][string]$ScriptPath)

    $info = @{ DownloadSubfolder = $null; VendorFolder = $null; AppFolder = $null }
    try {
        # Stream the file and stop as soon as all three vars are found. Avoids
        # the prior -TotalCount 120 cutoff that missed packagers (e.g.
        # package-teamviewerhost.ps1) where the declarations sit past line 120.
        foreach ($line in Get-Content -LiteralPath $ScriptPath -ErrorAction Stop) {
            if (-not $info.DownloadSubfolder -and $line -match '\$BaseDownloadRoot\s*=\s*Join-Path\s+\$DownloadRoot\s+"([^"]+)"') {
                $info.DownloadSubfolder = $matches[1]
            }
            if (-not $info.VendorFolder -and $line -match '^\s*\$VendorFolder\s*=\s*"([^"]+)"') {
                $info.VendorFolder = $matches[1]
            }
            if (-not $info.AppFolder -and $line -match '^\s*\$AppFolder\s*=\s*"([^"]+)"') {
                $info.AppFolder = $matches[1]
            }
            if ($info.DownloadSubfolder -and $info.VendorFolder -and $info.AppFolder) { break }
        }
    }
    catch { }
    return $info
}

function Get-PackagerLoggedPath {
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory)][string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $pattern = [regex]::Escape($Label) + '\s*:\s*(.+?)\s*$'
    foreach ($line in @($Text -split "`r?`n")) {
        if ($line -match $pattern) {
            return $Matches[1].Trim()
        }
    }
    return $null
}

function Find-NewestStageManifestForPackager {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [string]$DownloadRoot = $null
    )

    if ([string]::IsNullOrWhiteSpace($DownloadRoot)) { return $null }

    $info = Get-PackagerFolderInfo -ScriptPath $PackagerPath
    $searchRoot = $DownloadRoot
    if ($info.DownloadSubfolder) {
        $candidateRoot = Join-Path $DownloadRoot $info.DownloadSubfolder
        if (Test-Path -LiteralPath $candidateRoot) {
            $searchRoot = $candidateRoot
        }
    }

    if (-not (Test-Path -LiteralPath $searchRoot)) { return $null }
    $manifest = Get-ChildItem -LiteralPath $searchRoot -Filter 'stage-manifest.json' -Recurse -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1

    if ($manifest) { return $manifest.FullName }
    return $null
}

function Get-StageFileHashComparisonMessage {
    param([Parameter(Mandatory)]$Comparison)

    if ($Comparison.Pass) { return 'integrity verified' }
    if ($Comparison.Skipped) { return [string]$Comparison.Reason }

    $parts = New-Object System.Collections.Generic.List[string]
    if ($Comparison.Missing.Count -gt 0) {
        $sample = @($Comparison.Missing | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("missing {0}: {1}" -f $Comparison.Missing.Count, $sample))
    }
    if ($Comparison.Mismatches.Count -gt 0) {
        $sample = @($Comparison.Mismatches | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("mismatched {0}: {1}" -f $Comparison.Mismatches.Count, $sample))
    }
    if ($Comparison.Extra.Count -gt 0) {
        $sample = @($Comparison.Extra | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("extra {0}: {1}" -f $Comparison.Extra.Count, $sample))
    }
    if ($parts.Count -eq 0 -and $Comparison.Reason) { $parts.Add([string]$Comparison.Reason) }
    return ($parts.ToArray() -join '; ')
}

function Assert-PackagerStageIntegrity {
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [string]$DownloadRoot = $null
    )

    if ($Result.ExitCode -ne 0) { return }

    $stagePath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Stage complete'
    $manifestPath = $null
    if (-not [string]::IsNullOrWhiteSpace($stagePath)) {
        $manifestPath = Join-Path $stagePath 'stage-manifest.json'
    }
    if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
        $manifestPath = Find-NewestStageManifestForPackager -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot
    }
    if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
        throw "Stage integrity verification could not find stage-manifest.json."
    }

    $manifest = Read-StageManifest -Path $manifestPath
    $root = Split-Path -Path $manifestPath -Parent
    $comparison = Compare-StageFileHashes -Root $root -Expected $manifest.FileHashes
    if (-not $comparison.Pass) {
        throw ("Stage integrity verification failed: {0}" -f (Get-StageFileHashComparisonMessage -Comparison $comparison))
    }
}

function Assert-PackagerPackageIntegrity {
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$FileServerPath,
        [string]$DownloadRoot = $null,
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested'
    )

    if ($Result.ExitCode -ne 0) { return }

    $manifestPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Read stage manifest'
    if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
        $manifestPath = Find-NewestStageManifestForPackager -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot
    }
    if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
        throw "Package integrity verification could not find stage-manifest.json."
    }

    $manifest = Read-StageManifest -Path $manifestPath
    $networkContentPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Network content path'
    if ([string]::IsNullOrWhiteSpace($networkContentPath)) {
        $info = Get-PackagerFolderInfo -ScriptPath $PackagerPath
        if (-not $info.VendorFolder -or -not $info.AppFolder) {
            throw "Package integrity verification could not resolve the network content path."
        }
        if ($ContentLayout -eq 'Flat') {
            $networkContentPath = Join-Path (Join-Path $FileServerPath 'Applications') ('{0}-{1}-{2}' -f $info.VendorFolder, $info.AppFolder, $manifest.SoftwareVersion)
        }
        else {
            $networkContentPath = Join-Path (Join-Path (Join-Path $FileServerPath 'Applications') $info.VendorFolder) $info.AppFolder
            $networkContentPath = Join-Path $networkContentPath $manifest.SoftwareVersion
        }
    }

    $comparison = Compare-StageFileHashes -Root $networkContentPath -Expected $manifest.FileHashes
    if (-not $comparison.Pass) {
        throw ("Package integrity verification failed: {0}" -f (Get-StageFileHashComparisonMessage -Comparison $comparison))
    }
}

function Invoke-PackagerIntuneWinPostStep {
    # Produces <AppFolder>-<Version>.intunewin from the staged content and
    # copies it beside the network content version folder. The artifact
    # lands in the parent of both version folders, never inside them:
    # stage hash verification fails on any file added to verified content.
    # Failures never fail the package run - the MECM application already
    # exists when this executes - so the returned note carries Ok/Message
    # for the caller to surface.
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$FileServerPath,
        [string]$DownloadRoot = $null,
        [string]$ToolPath = '',
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested'
    )

    $note = [pscustomobject]@{
        Ok          = $false
        Message     = ''
        LocalPath   = ''
        NetworkPath = ''
    }

    try {
        if ([string]::IsNullOrWhiteSpace($ToolPath) -or -not (Test-Path -LiteralPath $ToolPath)) {
            $note.Message = 'IntuneWinAppUtil.exe not available; skipped. Configure it in MECM Preferences.'
            return $note
        }

        $manifestPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Read stage manifest'
        if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
            $manifestPath = Find-NewestStageManifestForPackager -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot
        }
        if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
            $note.Message = 'stage-manifest.json not found; skipped.'
            return $note
        }

        $manifest      = Read-StageManifest -Path $manifestPath
        $contentFolder = Split-Path -Path $manifestPath -Parent
        $version       = [string]$manifest.SoftwareVersion

        $info = Get-PackagerFolderInfo -ScriptPath $PackagerPath
        $baseName = if ($info.AppFolder) { [string]$info.AppFolder } else {
            ([IO.Path]::GetFileNameWithoutExtension($PackagerPath)) -replace '^package-', ''
        }
        $outputName = ((('{0}-{1}' -f $baseName, $version) -replace '[\\/:*?"<>|]', '_') + '.intunewin')

        $pkg = New-IntuneWinPackage `
            -ToolPath $ToolPath `
            -ContentFolder $contentFolder `
            -SetupFile 'install.bat' `
            -OutputFolder (Split-Path -Path $contentFolder -Parent) `
            -OutputName $outputName
        $note.LocalPath = [string]$pkg.IntuneWinPath

        $networkContentPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Network content path'
        if ([string]::IsNullOrWhiteSpace($networkContentPath) -and $info.VendorFolder -and $info.AppFolder) {
            if ($ContentLayout -eq 'Flat') {
                $networkContentPath = Join-Path (Join-Path $FileServerPath 'Applications') ('{0}-{1}-{2}' -f $info.VendorFolder, $info.AppFolder, $version)
            }
            else {
                $networkContentPath = Join-Path (Join-Path (Join-Path $FileServerPath 'Applications') $info.VendorFolder) $info.AppFolder
                $networkContentPath = Join-Path $networkContentPath $version
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($networkContentPath)) {
            $networkTarget = Join-Path (Split-Path -Path $networkContentPath -Parent) $outputName
            Copy-Item -LiteralPath $pkg.IntuneWinPath -Destination $networkTarget -Force -ErrorAction Stop
            $note.NetworkPath = $networkTarget
        }

        $note.Ok = $true
        $note.Message = ('created {0} ({1:N1} MB, SHA256 {2})' -f $outputName, ($pkg.SizeBytes / 1MB), $pkg.Sha256.Substring(0, 12))
        return $note
    }
    catch {
        $note.Message = ('creation failed: {0}' -f $_.Exception.Message)
        return $note
    }
}

function Compare-SemVer {
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B
    )
    try {
        $va = [version]($A -replace '[+-].*$', '')
        $vb = [version]($B -replace '[+-].*$', '')

        # Significant-part counts. Unset Build/Revision on [version] is -1.
        $aCount = 2
        if ($va.Build -ge 0) { $aCount = 3 }
        if ($va.Revision -ge 0) { $aCount = 4 }
        $bCount = 2
        if ($vb.Build -ge 0) { $bCount = 3 }
        if ($vb.Revision -ge 0) { $bCount = 4 }

        # Compare only the parts both sides actually provide. If one side has
        # extra trailing parts (e.g., MSI "26.2.2.2" vs vendor "26.2.2"), we
        # treat the extra parts as non-significant. This handles LibreOffice
        # and mRemoteNG where the MSI adds internal build numbers the vendor
        # doesn't publish as the version.
        $minCount = [Math]::Min($aCount, $bCount)
        $aParts = @($va.Major, $va.Minor, [Math]::Max($va.Build, 0), [Math]::Max($va.Revision, 0))
        $bParts = @($vb.Major, $vb.Minor, [Math]::Max($vb.Build, 0), [Math]::Max($vb.Revision, 0))

        for ($i = 0; $i -lt $minCount; $i++) {
            if ($aParts[$i] -lt $bParts[$i]) { return -1 }
            if ($aParts[$i] -gt $bParts[$i]) { return  1 }
        }
        return 0
    }
    catch { return 0 }
}

function Invoke-PackagerGetLatestVersion {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$FileServerPath = $null,
        [string]$DownloadRoot = $null,
        [string]$M365Channel = $null,
        [string]$M365DeployMode = $null
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.WorkingDirectory = Split-Path -Parent $PackagerPath
    $argsBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PackagerPath, '-SiteCode', $SiteCode, '-GetLatestVersionOnly')
    if ($FileServerPath -and (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath)) {
        $argsBase += @('-FileServerPath', $FileServerPath)
    }
    if ($DownloadRoot) { $argsBase += @('-DownloadRoot', $DownloadRoot) }
    if ($M365Channel) { $argsBase += @('-M365Channel', $M365Channel) }
    if ($M365DeployMode) { $argsBase += @('-M365DeployMode', $M365DeployMode) }
    Set-ProcessStartInfoArgumentList -StartInfo $psi -Arguments $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $p = New-Object System.Diagnostics.Process
    try {
        $p.StartInfo = $psi
        $null = $p.Start()
        $stdoutTask = $p.StandardOutput.ReadToEndAsync()
        $stderrTask = $p.StandardError.ReadToEndAsync()

        if (-not $p.WaitForExit(30000)) {
            try { $p.Kill() } catch {}
            throw "Packager timed out after 30 seconds."
        }

        $stdout = if ($stdoutTask.Wait(5000)) { $stdoutTask.Result } else { '' }
        $stderr = if ($stderrTask.Wait(5000)) { $stderrTask.Result } else { '' }

        if ($p.ExitCode -ne 0) {
            $msg = $stderr
            if ([string]::IsNullOrWhiteSpace($msg)) { $msg = $stdout }
            if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Packager returned exit code $($p.ExitCode)." }
            throw $msg.Trim()
        }

        $lines = @($stdout -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if (-not $lines -or $lines.Count -lt 1) { throw "No version output received." }

        $version = ([string]$lines[0]).Trim()
        if ($version -notmatch '^\d+(\.\d+){1,3}([+-]\d+)?$') {
            throw ("Unexpected version string: '{0}'" -f $version)
        }
        return $version
    }
    finally {
        if ($p) { try { $p.Dispose() } catch { } }
    }
}

function Get-MecmCurrentVersionByCMName {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = $null,
        [Parameter(Mandatory)][string]$CMName
    )

    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        try {
            if ($env:SMS_ADMIN_UI_PATH) {
                $cmModule = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
                if (Test-Path -LiteralPath $cmModule) {
                    Import-Module $cmModule -Force -ErrorAction Stop
                }
            }
        } catch { }
    }
    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        throw "ConfigMgr PowerShell cmdlets not available in this session."
    }

    # Capture the caller's location BEFORE touching the site drive so the
    # finally block restores it instead of leaving the shell parked on the
    # site drive.
    $savedLocation = Get-Location

    $existingDrive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue
    $connected = $false

    if ($existingDrive) {
        try {
            Set-Location "${SiteCode}:" -ErrorAction Stop
            $connected = $true
        }
        catch {
            # Existing drive won't enter (dead provider connection, e.g.
            # provider restart). Tear it down and rebuild from the provider.
            Remove-PSDrive -Name $SiteCode -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not $connected) {
        $providerRoot = $null
        if (-not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
            $providerRoot = $ProviderMachineName.Trim()
        }
        elseif ($existingDrive) {
            $providerRoot = [string]$existingDrive.Root
        }

        if ([string]::IsNullOrWhiteSpace($providerRoot)) {
            Set-Location $savedLocation -ErrorAction SilentlyContinue
            throw ("Failed to connect to CM site PSDrive '{0}:'. Open the ConfigMgr console once on this machine, or set Provider Machine in Options > MECM Preferences (the ProviderMachineName value from the AdminUI connect script)." -f $SiteCode)
        }

        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $providerRoot -ErrorAction Stop | Out-Null
            Set-Location "${SiteCode}:" -ErrorAction Stop
        }
        catch {
            Set-Location $savedLocation -ErrorAction SilentlyContinue
            throw ("Failed to connect to CM site PSDrive '{0}:' via provider '{1}': {2}" -f $SiteCode, $providerRoot, $_.Exception.Message)
        }
    }

    try {
        $apps = @(Get-CMApplication -Name $CMName -ErrorAction SilentlyContinue)
        if (-not $apps -or $apps.Count -eq 0) {
            $apps = @(Get-CMApplication -Name ("{0}*" -f $CMName) -ErrorAction SilentlyContinue)
        }

        if (-not $apps -or $apps.Count -eq 0) {
            return [pscustomobject]@{ Found = $false; DisplayName = $null; SoftwareVersion = $null; MatchCount = 0 }
        }

        $exact = $apps | Where-Object { $_.LocalizedDisplayName -eq $CMName -or $_.Name -eq $CMName }
        if ($exact -and $exact.Count -gt 0) {
            $chosen = $exact | Select-Object -First 1
        }
        else {
            $parsable = @()
            $nonParsable = @()
            foreach ($a in $apps) {
                try { $null = [version]$a.SoftwareVersion; $parsable += $a }
                catch { $nonParsable += $a }
            }
            if ($parsable.Count -gt 0) {
                $chosen = $parsable | Sort-Object { [version]$_.SoftwareVersion } -Descending | Select-Object -First 1
            }
            else {
                $chosen = $nonParsable | Sort-Object Name -Descending | Select-Object -First 1
            }
        }

        return [pscustomobject]@{
            Found           = $true
            DisplayName     = $chosen.LocalizedDisplayName
            SoftwareVersion = $chosen.SoftwareVersion
            MatchCount      = $apps.Count
        }
    }
    finally {
        Set-Location $savedLocation -ErrorAction SilentlyContinue
    }
}

function Invoke-ProcessWithStreaming {
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [Parameter(Mandatory)][string]$OutLog,
        [Parameter(Mandatory)][string]$ErrLog,
        [string]$StructuredLog = '',
        [System.Windows.Controls.TextBox]$LogTextBox = $null,
        # Idle timeout: kill the child if no stdout line arrives for this long.
        # Default 30 min covers slow MSIs without leaving silently-hung processes
        # wedging the GUI forever.
        [int]$IdleTimeoutSeconds = 1800
    )

    $p = New-Object System.Diagnostics.Process
    try {
        $p.StartInfo = $StartInfo
        $null = $p.Start()

        $outLines = New-Object System.Collections.Generic.List[string]
        $errTask = $p.StandardError.ReadToEndAsync()
        $reader   = $p.StandardOutput
        $lineTask = $reader.ReadLineAsync()

        $lastActivity = [DateTime]::UtcNow
        $timedOut = $false

        while ($true) {
            if ($lineTask.IsCompleted) {
                $line = $lineTask.Result
                if ($null -eq $line) { break }
                $outLines.Add($line)
                $lastActivity = [DateTime]::UtcNow

                if ($LogTextBox) {
                    $displayLine = $line -replace '^\[[\d: -]+\] \[\w+\s*\] ', ''
                    if ($displayLine.Trim()) {
                        Add-LogLine -Message ("  {0}" -f $displayLine)
                    }
                }

                $lineTask = $reader.ReadLineAsync()
            }

            # Idle timeout: child has stopped emitting stdout. Don't wait forever.
            if ($IdleTimeoutSeconds -gt 0 -and ([DateTime]::UtcNow - $lastActivity).TotalSeconds -gt $IdleTimeoutSeconds) {
                $timedOut = $true
                break
            }

            # WPF dispatcher pump
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )
            Start-Sleep -Milliseconds 50
        }

        if ($timedOut) {
            try { $p.Kill() } catch { }
            $p.WaitForExit(5000)
            $outLines.Add("[ERROR] Packager idle for $IdleTimeoutSeconds seconds; killed by Invoke-ProcessWithStreaming.")
        }
        elseif (-not $p.WaitForExit(15000)) {
            try { $p.Kill() } catch { }
            $p.WaitForExit(5000)
        }

        $stdout = ($outLines -join "`r`n")
        $stderr = if ($errTask.IsCompleted) { $errTask.Result } else { "" }

        Set-Content -LiteralPath $OutLog -Value $stdout -Encoding UTF8
        Set-Content -LiteralPath $ErrLog -Value $stderr -Encoding UTF8

        return [pscustomobject]@{
            ExitCode      = $p.ExitCode
            OutLog        = $OutLog
            ErrLog        = $ErrLog
            StructuredLog = $StructuredLog
            StdOut        = $stdout
            StdErr        = $stderr
        }
    }
    finally {
        if ($p) { try { $p.Dispose() } catch { } }
    }
}

function Invoke-PackagerStage {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [string]$M365Channel = $null,
        [string]$M365DeployMode = $null,
        [string]$SevenZipPath = '',
        [System.Windows.Controls.TextBox]$LogTextBox = $null
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-stage-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-stage-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-stage-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.WorkingDirectory = Split-Path -Parent $PackagerPath
    $argsBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PackagerPath, '-StageOnly', '-LogPath', $structuredLog)
    if ($DownloadRoot) { $argsBase += @('-DownloadRoot', $DownloadRoot) }
    if ($M365Channel) { $argsBase += @('-M365Channel', $M365Channel) }
    if ($M365DeployMode) { $argsBase += @('-M365DeployMode', $M365DeployMode) }
    Set-ProcessStartInfoArgumentList -StartInfo $psi -Arguments $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    # Pass detected-tool paths to the packager via env vars so it can use
    # non-default install locations. Packagers that need these tools check
    # the env var first and fall back to Program Files\<tool> defaults.
    Set-PackagerEnvironment -StartInfo $psi -SevenZipPath $SevenZipPath

    $result = Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
    Assert-PackagerStageIntegrity -Result $result -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot
    return $result
}

function Invoke-PackagerPackage {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = '',
        [AllowEmptyString()][string]$Comment = '',
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [string]$M365Channel = $null,
        [string]$M365DeployMode = $null,
        [int]$EstimatedRuntimeMins = 0,
        [int]$MaximumRuntimeMins = 0,
        [string]$SevenZipPath = '',
        [switch]$CreateIntuneWin,
        [string]$IntuneWinToolPath = '',
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested',
        [string]$RequirementsJson = '',
        [string]$VariantsJson = '',
        [string]$CommandsJson = '',
        [System.Windows.Controls.TextBox]$LogTextBox = $null
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-package-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-package-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-package-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.WorkingDirectory = Split-Path -Parent $PackagerPath
    $argsBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PackagerPath, '-PackageOnly', '-SiteCode', $SiteCode, '-Comment', $Comment, '-LogPath', $structuredLog)
    if (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath) {
        $argsBase += @('-FileServerPath', $FileServerPath)
    }
    if ($DownloadRoot) { $argsBase += @('-DownloadRoot', $DownloadRoot) }
    if ($ContentLayout) { $argsBase += @('-ContentLayout', $ContentLayout) }
    if ($M365Channel) { $argsBase += @('-M365Channel', $M365Channel) }
    if ($M365DeployMode) { $argsBase += @('-M365DeployMode', $M365DeployMode) }
    if ($EstimatedRuntimeMins -gt 0) { $argsBase += @('-EstimatedRuntimeMins', [string]$EstimatedRuntimeMins) }
    if ($MaximumRuntimeMins -gt 0) { $argsBase += @('-MaximumRuntimeMins', [string]$MaximumRuntimeMins) }
    Set-ProcessStartInfoArgumentList -StartInfo $psi -Arguments $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    Set-PackagerEnvironment -StartInfo $psi -SevenZipPath $SevenZipPath -ProviderMachineName $ProviderMachineName -RequirementsJson $RequirementsJson -VariantsJson $VariantsJson -CommandsJson $CommandsJson

    $result = Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
    Assert-PackagerPackageIntegrity -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ContentLayout $ContentLayout

    # Optional post-step: produce a .intunewin beside the network content.
    # Runs only after integrity passes; failures ride on the result for the
    # caller to surface, never thrown - the MECM application already exists
    # by this point.
    if ($CreateIntuneWin -and $result.ExitCode -eq 0) {
        $intuneNote = Invoke-PackagerIntuneWinPostStep -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ToolPath $IntuneWinToolPath -ContentLayout $ContentLayout
        if ($intuneNote) {
            $result | Add-Member -NotePropertyName IntuneWin -NotePropertyValue $intuneNote -Force
        }
    }
    return $result
}

function Set-PackagerEnvironment {
    # Forwards DetectedTools paths to the packager child process via env
    # vars so the packager can resolve tools without hardcoded paths.
    # The child inherits the parent's environment block unless we set
    # $psi.EnvironmentVariables here. Accepts the 7-Zip path as a
    # parameter so this function and its callers work identically in the
    # main runspace and in the background pipeline's STA runspace (which
    # has its own $script: session state and cannot see $script:Prefs).
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [string]$SevenZipPath,
        [string]$ProviderMachineName,
        [string]$RequirementsJson,
        [string]$VariantsJson,
        [string]$CommandsJson
    )
    if (-not [string]::IsNullOrWhiteSpace($SevenZipPath)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_SEVENZIP'] = [string]$SevenZipPath
    }
    if (-not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_CM_PROVIDER'] = [string]$ProviderMachineName
    }
    if (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_REQUIREMENTS'] = [string]$RequirementsJson
    }
    if (-not [string]::IsNullOrWhiteSpace($VariantsJson)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_VARIANTS'] = [string]$VariantsJson
    }
    if (-not [string]::IsNullOrWhiteSpace($CommandsJson)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_COMMANDS'] = [string]$CommandsJson
    }
}

function ConvertTo-RequirementsJson {
    # Maps one Deployment Conditions prefs entry onto the
    # APP_PACKAGER_REQUIREMENTS JSON that New-DeploymentTypeRequirementRules
    # consumes. Returns '' when the entry asks for nothing so callers can
    # skip setting the env var.
    param($Entry)

    if (-not $Entry) { return '' }
    $rules = @()
    if ([string]$Entry.Architecture -in @('x64', 'ARM64')) {
        $rules += @{ ConditionId = 'cpu-arch'; Value = [string]$Entry.Architecture }
    }
    $langs = @()
    if ($null -ne $Entry.Languages) {
        $langs = @($Entry.Languages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    }
    if ($langs.Count -gt 0) {
        $rules += @{ ConditionId = 'os-language'; Cultures = $langs }
    }
    switch ([string]$Entry.Network) {
        'VpnOnly'    { $rules += @{ ConditionId = 'vpn-connected'; Value = $true } }
        'OnSiteOnly' { $rules += @{ ConditionId = 'vpn-connected'; Value = $false } }
    }
    if ($rules.Count -eq 0) { return '' }
    return (@{ SchemaVersion = 1; Rules = $rules } | ConvertTo-Json -Depth 4 -Compress)
}

function ConvertTo-VariantsJson {
    # Maps one Deployment Conditions prefs entry onto the
    # APP_PACKAGER_VARIANTS JSON a SupportsVariants packager consumes via
    # Get-RequestedPackagerVariants. Returns '' when no split is selected.
    param($Entry)

    if (-not $Entry -or -not $Entry.PSObject.Properties['Split']) { return '' }
    $split = [string]$Entry.Split
    if ($split -notin @('Architecture', 'Language', 'Network')) { return '' }
    $doc = @{ SchemaVersion = 1; Split = $split }
    if ($split -eq 'Language' -and $null -ne $Entry.Languages) {
        $doc['Languages'] = @($Entry.Languages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    }
    return ($doc | ConvertTo-Json -Depth 4 -Compress)
}

function ConvertTo-CommandsJson {
    # Maps one CommandOverrides prefs entry onto the APP_PACKAGER_COMMANDS
    # JSON that Get-RequestedCommandOverrides consumes. Returns '' when
    # the entry carries no command.
    param($Entry)

    if (-not $Entry) { return '' }
    $inst = ([string]$Entry.Install).Trim()
    $uninst = ([string]$Entry.Uninstall).Trim()
    if (-not $inst -and -not $uninst) { return '' }
    $doc = @{ SchemaVersion = 1 }
    if ($inst) { $doc['Install'] = $inst }
    if ($uninst) { $doc['Uninstall'] = $uninst }
    return ($doc | ConvertTo-Json -Depth 3 -Compress)
}

function Get-CommandsMapForContext {
    # Prebuilt on the UI thread, same reason as Get-RequirementsMapForContext.
    $map = @{}
    try {
        $apps = $script:Prefs.CommandOverrides.Apps
        if ($apps) {
            foreach ($prop in $apps.PSObject.Properties) {
                $json = ConvertTo-CommandsJson -Entry $prop.Value
                if ($json) { $map[$prop.Name] = $json }
            }
        }
    } catch { }
    return $map
}

function Get-VariantsMapForContext {
    # Prebuilt on the UI thread, same reason as Get-RequirementsMapForContext.
    $map = @{}
    try {
        $apps = $script:Prefs.DeploymentConditions.Apps
        if ($apps) {
            foreach ($prop in $apps.PSObject.Properties) {
                $json = ConvertTo-VariantsJson -Entry $prop.Value
                if ($json) { $map[$prop.Name] = $json }
            }
        }
    } catch { }
    return $map
}

function Get-RequirementsMapForContext {
    # Prebuilt on the UI thread because the background STA runspace has its
    # own session state and cannot read $script:Prefs.
    $map = @{}
    try {
        $apps = $script:Prefs.DeploymentConditions.Apps
        if ($apps) {
            foreach ($prop in $apps.PSObject.Properties) {
                $json = ConvertTo-RequirementsJson -Entry $prop.Value
                if ($json) { $map[$prop.Name] = $json }
            }
        }
    } catch { }
    return $map
}

function Get-SevenZipPathForContext {
    # Resolves the detected 7-Zip path from prefs, safe against the
    # first-run case where DetectedTools or SevenZipCli may not be
    # populated yet. Used when building the Context hashtable passed
    # into Invoke-MultiAppPipeline.
    try {
        if ($script:Prefs -and $script:Prefs.DetectedTools -and $script:Prefs.DetectedTools.SevenZipCli -and $script:Prefs.DetectedTools.SevenZipCli.Found) {
            return [string]$script:Prefs.DetectedTools.SevenZipCli.ExePath
        }
    } catch { }
    return ''
}

function Get-IntuneWinToolPathForContext {
    # Same prefs-safe resolution as Get-SevenZipPathForContext, for the
    # Win32 Content Prep Tool.
    try {
        if ($script:Prefs -and $script:Prefs.DetectedTools -and $script:Prefs.DetectedTools.IntuneWinAppUtil -and $script:Prefs.DetectedTools.IntuneWinAppUtil.Found) {
            return [string]$script:Prefs.DetectedTools.IntuneWinAppUtil.ExePath
        }
    } catch { }
    return ''
}

function Select-OnlyUpdateAvailable {
    # Clear every row first, then select within the visible (filtered) set
    # only: a row hidden by the grid filter must never keep or gain
    # Selected=true, or a later Stage/Package would run on rows the user
    # cannot see.
    foreach ($item in $script:PackagerData) { $item.Selected = $false }
    foreach ($item in @($dataGrid.ItemsSource)) {
        $item.Selected = ($item.Status -eq "Update available")
    }
}

# =============================================================================
# Log helper (WPF version)
# =============================================================================
$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:MaxLogLines = 4000
$script:LogTrimBatch = 500

function Add-LogEntry {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Line)

    if ($null -eq $script:LogLines) {
        $script:LogLines = New-Object System.Collections.Generic.List[string]
    }

    [void]$script:LogLines.Add($Line)

    if ($script:LogLines.Count -eq 1) {
        $txtLog.AppendText($Line)
    }
    else {
        $txtLog.AppendText([Environment]::NewLine + $Line)
    }

    if ($script:LogLines.Count -gt $script:MaxLogLines) {
        $overflow = $script:LogLines.Count - $script:MaxLogLines
        $removeCount = [Math]::Min(
            $script:LogLines.Count,
            [Math]::Max($script:LogTrimBatch, $overflow)
        )

        if ($removeCount -gt 0) {
            $script:LogLines.RemoveRange(0, $removeCount)
            $txtLog.Text = [string]::Join([Environment]::NewLine, $script:LogLines.ToArray())
        }
    }

    $txtLog.ScrollToEnd()
}

function Add-LogSeparator {
    Add-LogEntry -Line ''
}

function Add-LogLine {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    Add-LogEntry -Line $line
}

# =============================================================================
# Window state persistence
# =============================================================================
function Get-WindowStatePath {
    Join-Path $PSScriptRoot "AppPackager.windowstate.json"
}


# =============================================================================
# CWA Switches (carried over)
# =============================================================================
function Get-CwaSwitchesPath {
    Join-Path (Join-Path $PSScriptRoot "Packagers") "citrix-workspace-switches.json"
}

function Read-CwaSwitches {
    $defaults = [pscustomobject]@{
        Store = [pscustomobject]@{ Name = ""; Url = "" }
        Installation = [pscustomobject]@{
            CleanInstall     = $true
            IncludeSSON      = $true
            EnableSSON       = $true
            AppProtection    = $false
            SessionPreLaunch = $false
            SelfServiceMode  = $true
        }
        Plugins = [pscustomobject]@{
            MSTeamsPlugin        = $true
            ZoomPlugin           = $true
            WebExPlugin          = $false
            UberAgent            = $false
            UberAgentSkipUpgrade = $false
            EPAClient            = $true
            SessionRecording     = $false
        }
        UpdateAndTelemetry = [pscustomobject]@{
            AutoUpdateCheck = "disabled"
            EnableCEIP      = $false
            EnableTracing   = $false
        }
        StorePolicy = [pscustomobject]@{
            AllowAddStore = "S"
            AllowSavePwd  = "S"
        }
        Components = [pscustomobject]@{
            Customize      = $false
            ReceiverInside = $true
            ICA_Client     = $true
            AM             = $true
            SelfService    = $true
            DesktopViewer  = $true
            WebHelper      = $true
            BCR_Client     = $true
            USB            = $false
            SSON           = $false
        }
    }

    $path = Get-CwaSwitchesPath
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($null -ne $data.Store) {
            if ($null -ne $data.Store.Name) { $defaults.Store.Name = [string]$data.Store.Name }
            if ($null -ne $data.Store.Url)  { $defaults.Store.Url  = [string]$data.Store.Url }
        }
        foreach ($prop in @('CleanInstall','IncludeSSON','EnableSSON','AppProtection','SessionPreLaunch','SelfServiceMode')) {
            if ($null -ne $data.Installation.$prop) { $defaults.Installation.$prop = [bool]$data.Installation.$prop }
        }
        foreach ($prop in @('MSTeamsPlugin','ZoomPlugin','WebExPlugin','UberAgent','UberAgentSkipUpgrade','EPAClient','SessionRecording')) {
            if ($null -ne $data.Plugins.$prop) { $defaults.Plugins.$prop = [bool]$data.Plugins.$prop }
        }
        if ($null -ne $data.UpdateAndTelemetry) {
            if ($null -ne $data.UpdateAndTelemetry.AutoUpdateCheck) { $defaults.UpdateAndTelemetry.AutoUpdateCheck = [string]$data.UpdateAndTelemetry.AutoUpdateCheck }
            if ($null -ne $data.UpdateAndTelemetry.EnableCEIP)      { $defaults.UpdateAndTelemetry.EnableCEIP      = [bool]$data.UpdateAndTelemetry.EnableCEIP }
            if ($null -ne $data.UpdateAndTelemetry.EnableTracing)   { $defaults.UpdateAndTelemetry.EnableTracing   = [bool]$data.UpdateAndTelemetry.EnableTracing }
        }
        if ($null -ne $data.StorePolicy) {
            if ($null -ne $data.StorePolicy.AllowAddStore) { $defaults.StorePolicy.AllowAddStore = [string]$data.StorePolicy.AllowAddStore }
            if ($null -ne $data.StorePolicy.AllowSavePwd)  { $defaults.StorePolicy.AllowSavePwd  = [string]$data.StorePolicy.AllowSavePwd }
        }
        foreach ($prop in @('Customize','ReceiverInside','ICA_Client','AM','SelfService','DesktopViewer','WebHelper','BCR_Client','USB','SSON')) {
            if ($null -ne $data.Components.$prop) { $defaults.Components.$prop = [bool]$data.Components.$prop }
        }
    }
    catch { }

    return $defaults
}

function Save-CwaSwitches {
    param([Parameter(Mandatory)][pscustomobject]$Switches)
    $path = Get-CwaSwitchesPath
    $json = $Switches | ConvertTo-Json -Depth 3
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

# ----- TeamViewer Host mass-deployment config -----
function Get-TvHostConfigPath {
    Join-Path (Join-Path $PSScriptRoot "Packagers") "teamviewer-host-config.json"
}

function Read-TvHostConfig {
    $defaults = [pscustomobject]@{
        ApiToken              = ""
        CustomConfigId        = ""
        AssignmentOptions     = ""
        RemoveDesktopShortcut = $true
    }

    $path = Get-TvHostConfigPath
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($null -ne $data.ApiToken)              { $defaults.ApiToken              = [string]$data.ApiToken }
        if ($null -ne $data.CustomConfigId)        { $defaults.CustomConfigId        = [string]$data.CustomConfigId }
        if ($null -ne $data.AssignmentOptions)     { $defaults.AssignmentOptions     = [string]$data.AssignmentOptions }
        if ($null -ne $data.RemoveDesktopShortcut) { $defaults.RemoveDesktopShortcut = [bool]$data.RemoveDesktopShortcut }
    }
    catch { }

    return $defaults
}

function Save-TvHostConfig {
    param([Parameter(Mandatory)][pscustomobject]$Config)
    $path = Get-TvHostConfigPath
    $json = $Config | ConvertTo-Json -Depth 2
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

# =============================================================================
# Batch-mode dispatcher (headless; no WPF)
# =============================================================================
function Invoke-BatchUpdate {
    <#
    .SYNOPSIS
        CLI-mode batch driver for Full Run. Used ONLY by the -BatchMode
        command-line entry point (Write-Log / stdout logging, no WPF).

    .DESCRIPTION
        Kept as the CLI codepath for scheduled / headless invocations. The
        GUI's Full Run button uses Invoke-MultiAppPipeline instead (bg
        runspace + progress overlay + DispatcherTimer polling). The two
        paths intentionally diverge: the GUI path streams per-app status
        through the overlay, while the CLI path serializes Write-Log
        lines to stdout/file for tail / grep-friendly CI output.

        Don't unify the two without agreeing on a common streaming shape
        first - the GUI path's dependency on the UI dispatcher and the
        CLI path's dependency on plain stdout are not trivially merged.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagersRoot,
        [Parameter(Mandatory)][string[]]$Apps,
        [ValidateSet('Report','Stage','StageAndPackage')][string]$OnUpdateFound = 'Report',
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = '',
        [string]$FileServerPath,
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested',
        [string]$DownloadRoot,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30,
        [string]$Comment = '',
        [string]$SevenZipPath = '',
        [pscustomobject]$CadenceOverrides,
        [pscustomobject]$ConditionApps = $null,
        [pscustomobject]$CommandApps = $null,
        [switch]$Force
    )

    $defaultCadenceDays = 7
    $results = @()
    foreach ($appKey in $Apps) {
        $baseName = if ($appKey -like 'package-*') { $appKey } else { "package-$appKey" }
        $scriptPath = Join-Path $PackagersRoot ("{0}.ps1" -f $baseName)
        if (-not (Test-Path -LiteralPath $scriptPath)) {
            Write-Log ("[batch] Packager not found: {0}" -f $scriptPath) -Level ERROR
            $results += [pscustomobject]@{ Name = $baseName; Action = 'NotFound'; OldVersion = $null; NewVersion = $null; Reason = 'script missing' }
            continue
        }

        $history = Read-PackagerHistory
        $lastKnown    = $null
        $lastChecked  = $null
        $lastStaged   = $null
        $lastPackaged = $null
        if ($history.ContainsKey($baseName)) {
            $entry = $history[$baseName]
            if ($entry -is [hashtable]) {
                $lastKnown    = $entry['LastKnownVersion']
                $lastChecked  = $entry['LastChecked']
                $lastStaged   = $entry['LastStaged']
                $lastPackaged = $entry['LastPackaged']
            } else {
                $lastKnown    = $entry.LastKnownVersion
                $lastChecked  = $entry.LastChecked
                $lastStaged   = $entry.LastStaged
                $lastPackaged = $entry.LastPackaged
            }
        }

        # Cadence gate applies to Report only. Stage / StageAndPackage
        # always run when the user clicks Full Run - the cadence is for
        # throttling vendor queries, not for blocking explicit packaging.
        if ($OnUpdateFound -eq 'Report' -and -not $Force -and $lastChecked) {
            $cadenceDays = $defaultCadenceDays
            $cadenceFromOverride = $false
            if ($CadenceOverrides) {
                $overrideProp = $CadenceOverrides.PSObject.Properties[$baseName]
                if ($overrideProp) {
                    $parsedOverride = 0
                    if ([int]::TryParse([string]$overrideProp.Value, [ref]$parsedOverride) -and $parsedOverride -ge 1) {
                        $cadenceDays = $parsedOverride
                        $cadenceFromOverride = $true
                    }
                }
            }
            if (-not $cadenceFromOverride) {
                try {
                    $meta = Get-PackagerMetadata -Path $scriptPath
                    if ($meta.UpdateCadenceDays -and [int]$meta.UpdateCadenceDays -ge 1) {
                        $cadenceDays = [int]$meta.UpdateCadenceDays
                    }
                } catch { }
            }
            if ($cadenceDays -lt 1) { $cadenceDays = $defaultCadenceDays }

            try {
                $lastCheckedDt = [datetime]$lastChecked
                $nextDue = $lastCheckedDt.ToUniversalTime().AddDays($cadenceDays)
                $nowUtc  = (Get-Date).ToUniversalTime()
                if ($nextDue -gt $nowUtc) {
                    $daysRemaining = [int][math]::Ceiling(($nextDue - $nowUtc).TotalDays)
                    Write-Log ("[batch] [Skipped] {0}: cadence {1}d, next check in {2}d" -f $baseName, $cadenceDays, $daysRemaining) -Level INFO
                    $results += [pscustomobject]@{ Name = $baseName; Action = 'Skipped'; OldVersion = $lastKnown; NewVersion = $null; Reason = ("cadence {0}d, {1}d remaining" -f $cadenceDays, $daysRemaining) }
                    continue
                }
            } catch { }
        }

        # 1. Discover vendor's current version. On failure, don't update
        # LastChecked - next run should retry, not wait for cadence.
        $latest = $null
        try {
            $latest = Invoke-PackagerGetLatestVersion -PackagerPath $scriptPath -SiteCode $SiteCode -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot
        } catch {
            Write-Log ("[batch] {0}: check failed: {1}" -f $baseName, $_.Exception.Message) -Level WARN
            $results += [pscustomobject]@{ Name = $baseName; Action = 'CheckFailed'; OldVersion = $lastKnown; NewVersion = $null; Reason = $_.Exception.Message }
            continue
        }

        # 2. Decide whether to act. NoChange short-circuit applies only
        # when there's nothing new to do: same version AND (for Stage /
        # StageAndPackage) we've already staged/packaged that version
        # at least once. First-time stages and Force=on always run.
        $versionChanged = (-not $lastKnown) -or ($lastKnown -ne $latest)
        $neverStaged    = ($OnUpdateFound -eq 'Stage'           -and -not $lastStaged)
        $neverPackaged  = ($OnUpdateFound -eq 'StageAndPackage' -and -not $lastPackaged)
        $shouldAct      = $versionChanged -or $Force -or $neverStaged -or $neverPackaged

        if (-not $shouldAct) {
            Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result NoChange
            Write-Log ("[batch] [NoChange] {0}: {1}" -f $baseName, $latest) -Level INFO
            $results += [pscustomobject]@{ Name = $baseName; Action = 'NoChange'; OldVersion = $lastKnown; NewVersion = $latest }
            continue
        }

        # 3. Taking action
        $oldDisplay = if ($lastKnown) { $lastKnown } else { '(none)' }
        $label = if ($versionChanged) { 'Updated' } else { 'Forced' }
        Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result Updated
        Write-Log ("[batch] [{0}] {1}: {2} -> {3} (action={4})" -f $label, $baseName, $oldDisplay, $latest, $OnUpdateFound) -Level INFO

        if ($OnUpdateFound -eq 'Report') {
            $results += [pscustomobject]@{ Name = $baseName; Action = 'Reported'; OldVersion = $lastKnown; NewVersion = $latest }
            continue
        }

        # 4. Invoke the packager for Stage / StageAndPackage
        $pkgArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-SiteCode', $SiteCode)
        if ($FileServerPath)       { $pkgArgs += @('-FileServerPath',       $FileServerPath)       }
        if ($ContentLayout)        { $pkgArgs += @('-ContentLayout',        $ContentLayout)        }
        if ($DownloadRoot)         { $pkgArgs += @('-DownloadRoot',         $DownloadRoot)         }
        if ($EstimatedRuntimeMins) { $pkgArgs += @('-EstimatedRuntimeMins', $EstimatedRuntimeMins) }
        if ($MaximumRuntimeMins)   { $pkgArgs += @('-MaximumRuntimeMins',   $MaximumRuntimeMins)   }
        if ($Comment)              { $pkgArgs += @('-Comment',              $Comment)              }
        if ($OnUpdateFound -eq 'Stage') { $pkgArgs += '-StageOnly' }

        $requirementsJson = ''
        $variantsJson = ''
        $commandsJson = ''
        if ($ConditionApps) {
            $condProp = $ConditionApps.PSObject.Properties[$baseName]
            if ($condProp) {
                $requirementsJson = ConvertTo-RequirementsJson -Entry $condProp.Value
                $variantsJson = ConvertTo-VariantsJson -Entry $condProp.Value
            }
        }
        if ($CommandApps) {
            $cmdProp = $CommandApps.PSObject.Properties[$baseName]
            if ($cmdProp) { $commandsJson = ConvertTo-CommandsJson -Entry $cmdProp.Value }
        }

        try {
            $restoreSevenZipEnv = $false
            $previousSevenZipEnv = $null
            $restoreProviderEnv = $false
            $previousProviderEnv = $null
            $restoreRequirementsEnv = $false
            $previousRequirementsEnv = $null
            $restoreVariantsEnv = $false
            $previousVariantsEnv = $null
            $restoreCommandsEnv = $false
            $previousCommandsEnv = $null
            $packagerWorkingDirectory = Split-Path -Parent $scriptPath
            $pushedPackagerLocation = $false
            try {
                if (-not [string]::IsNullOrWhiteSpace($SevenZipPath)) {
                    $restoreSevenZipEnv = $true
                    $previousSevenZipEnv = $env:APP_PACKAGER_SEVENZIP
                    $env:APP_PACKAGER_SEVENZIP = $SevenZipPath
                }
                if (-not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
                    $restoreProviderEnv = $true
                    $previousProviderEnv = $env:APP_PACKAGER_CM_PROVIDER
                    $env:APP_PACKAGER_CM_PROVIDER = $ProviderMachineName
                }
                if (-not [string]::IsNullOrWhiteSpace($requirementsJson)) {
                    $restoreRequirementsEnv = $true
                    $previousRequirementsEnv = $env:APP_PACKAGER_REQUIREMENTS
                    $env:APP_PACKAGER_REQUIREMENTS = $requirementsJson
                }
                if (-not [string]::IsNullOrWhiteSpace($variantsJson)) {
                    $restoreVariantsEnv = $true
                    $previousVariantsEnv = $env:APP_PACKAGER_VARIANTS
                    $env:APP_PACKAGER_VARIANTS = $variantsJson
                }
                if (-not [string]::IsNullOrWhiteSpace($commandsJson)) {
                    $restoreCommandsEnv = $true
                    $previousCommandsEnv = $env:APP_PACKAGER_COMMANDS
                    $env:APP_PACKAGER_COMMANDS = $commandsJson
                }

                Push-Location -LiteralPath $packagerWorkingDirectory
                $pushedPackagerLocation = $true
                & powershell.exe @pkgArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    if ($line) { Write-Log ("[batch:{0}] {1}" -f $baseName, $line) -Level INFO }
                }
            }
            finally {
                if ($pushedPackagerLocation) {
                    try { Pop-Location } catch { }
                }
                if ($restoreSevenZipEnv) {
                    if ($null -ne $previousSevenZipEnv) {
                        $env:APP_PACKAGER_SEVENZIP = $previousSevenZipEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_SEVENZIP -ErrorAction SilentlyContinue
                    }
                }
                if ($restoreProviderEnv) {
                    if ($null -ne $previousProviderEnv) {
                        $env:APP_PACKAGER_CM_PROVIDER = $previousProviderEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_CM_PROVIDER -ErrorAction SilentlyContinue
                    }
                }
                if ($restoreCommandsEnv) {
                    if ($null -ne $previousCommandsEnv) {
                        $env:APP_PACKAGER_COMMANDS = $previousCommandsEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_COMMANDS -ErrorAction SilentlyContinue
                    }
                }
                if ($restoreVariantsEnv) {
                    if ($null -ne $previousVariantsEnv) {
                        $env:APP_PACKAGER_VARIANTS = $previousVariantsEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_VARIANTS -ErrorAction SilentlyContinue
                    }
                }
                if ($restoreRequirementsEnv) {
                    if ($null -ne $previousRequirementsEnv) {
                        $env:APP_PACKAGER_REQUIREMENTS = $previousRequirementsEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_REQUIREMENTS -ErrorAction SilentlyContinue
                    }
                }
            }
            $rc = $LASTEXITCODE
            if ($rc -ne 0 -and $rc -ne 3010) { throw "Packager exited with code $rc" }

            Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $latest -Result Updated
            if ($OnUpdateFound -eq 'StageAndPackage') {
                Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $latest -Result Updated
            }
            $results += [pscustomobject]@{ Name = $baseName; Action = $OnUpdateFound; OldVersion = $lastKnown; NewVersion = $latest }
        } catch {
            # Don't update LastChecked on failure - next run should retry
            # immediately, not wait for cadence to expire.
            Write-Log ("[batch] {0}: action failed: {1}" -f $baseName, $_.Exception.Message) -Level ERROR
            $results += [pscustomobject]@{ Name = $baseName; Action = 'Failed'; OldVersion = $lastKnown; NewVersion = $latest; Reason = $_.Exception.Message }
        }
    }

    return ,$results
}

# =============================================================================
# Batch-mode entry point (exits before WPF)
# =============================================================================
if ($BatchMode) {
    if ($LogPath) { Initialize-Logging -LogPath $LogPath }

    if (-not $Apps -or $Apps.Count -eq 0) {
        Write-Log "[batch] -BatchMode requires -Apps <list>. Aborting." -Level ERROR
        exit 2
    }
    # Child-process arg passing collapses string[] to a single comma-joined
    # string; split it back out if that happened.
    if ($Apps.Count -eq 1 -and $Apps[0] -match ',') {
        $Apps = @($Apps[0] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $prefs = if (Test-Path (Get-PreferencesPath)) { Read-Preferences } else { $null }
    $fileServerPath   = if ($prefs -and $prefs.FileShareRoot)             { $prefs.FileShareRoot }             else { $null }
    $contentLayout    = if ($prefs -and $prefs.ContentLayout)             { [string]$prefs.ContentLayout }     else { 'Nested' }
    $downloadRoot     = if ($prefs -and $prefs.DownloadRoot)              { $prefs.DownloadRoot }              else { $null }
    $providerForBatch = if ($script:Prefs -and $script:Prefs.ProviderMachineName) { [string]$script:Prefs.ProviderMachineName } else { $null }
    $cadenceOverrides = if ($prefs -and $prefs.AppFlow.CadenceOverrides)  { $prefs.AppFlow.CadenceOverrides }  else { $null }
    $conditionApps    = if ($prefs -and $prefs.DeploymentConditions -and $prefs.DeploymentConditions.Apps) { $prefs.DeploymentConditions.Apps } else { $null }
    $commandApps      = if ($prefs -and $prefs.CommandOverrides -and $prefs.CommandOverrides.Apps) { $prefs.CommandOverrides.Apps } else { $null }
    $sevenZipPath     = $null
    if ($prefs -and $prefs.DetectedTools -and $prefs.DetectedTools.SevenZipCli -and $prefs.DetectedTools.SevenZipCli.Found) {
        $sevenZipPath = [string]$prefs.DetectedTools.SevenZipCli.ExePath
    }
    if ([string]::IsNullOrWhiteSpace($sevenZipPath)) {
        try {
            $sevenZipProbe = Invoke-DetectSevenZipCli
            if ($sevenZipProbe -and $sevenZipProbe.Found) {
                $sevenZipPath = [string]$sevenZipProbe.ExePath
            }
        } catch { }
    }

    Write-Log ("[batch] Starting: {0} app(s), OnUpdateFound={1}" -f $Apps.Count, $OnUpdateFound) -Level INFO

    $summary = Invoke-BatchUpdate `
        -PackagersRoot     $PackagersRoot `
        -Apps              $Apps `
        -OnUpdateFound     $OnUpdateFound `
        -SiteCode          $SiteCode `
        -ProviderMachineName $providerForBatch `
        -FileServerPath    $fileServerPath `
        -ContentLayout     $contentLayout `
        -DownloadRoot      $downloadRoot `
        -SevenZipPath      $sevenZipPath `
        -CadenceOverrides  $cadenceOverrides `
        -ConditionApps     $conditionApps `
        -CommandApps       $commandApps `
        -Force:$Force

    Write-Log "" -Level INFO
    Write-Log "[batch] Summary:" -Level INFO
    foreach ($r in $summary) {
        $ov = if ($r.OldVersion) { $r.OldVersion } else { '(none)' }
        $nv = if ($r.NewVersion) { $r.NewVersion } else { '(n/a)'  }
        Write-Log ("[batch]   {0,-30} {1,-15} {2} -> {3}" -f $r.Name, $r.Action, $ov, $nv) -Level INFO
    }

    $failed = @($summary | Where-Object { $_.Action -in @('Failed','CheckFailed','NotFound') })
    if ($failed.Count -gt 0) { exit 1 } else { exit 0 }
}

# =============================================================================
# Data model - ObservableCollection of PSCustomObjects
# =============================================================================
$script:PackagerData = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]

# =============================================================================
# Parse XAML and create window
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# =============================================================================
# Title-bar drag fallback. PS51-WPF-033. SuiteCommon owns the hook and its
# state; wire on every MetroWindow (main window and every modal popup).
# =============================================================================
Install-TitleBarDragFallback -Window $window

# No window icon - the old .ico lacks transparency and doesn't fit the MahApps theme.
# Taskbar shows the PowerShell icon (PS5.1 WPF limitation without a compiled Application).

# =============================================================================
# Find named controls
# =============================================================================
$txtAppTitle     = $window.FindName('txtAppTitle')
$toggleTheme     = $window.FindName('toggleTheme')
$txtThemeLabel   = $window.FindName('txtThemeLabel')
$btnCheckLatest  = $window.FindName('btnCheckLatest')
$btnCheckMECM    = $window.FindName('btnCheckMECM')
$btnStage        = $window.FindName('btnStage')
$btnPackage      = $window.FindName('btnPackage')
$btnFullRun      = $window.FindName('btnFullRun')
$btnOptions      = $window.FindName('btnOptions')
$toggleDebugCols = $window.FindName('toggleDebugCols')
$txtGridFilter = $window.FindName('txtGridFilter')
$txtComment      = $window.FindName('txtComment')
$dataGrid        = $window.FindName('dataGrid')
$colSelected     = $window.FindName('colSelected')
$txtLog          = $window.FindName('txtLog')
$lblLogOutput    = $window.FindName('lblLogOutput')
$txtStatus       = $window.FindName('txtStatus')
$colCMName       = $window.FindName('colCMName')
$colScript       = $window.FindName('colScript')
$colVendorURL    = $window.FindName('colVendorURL')
$colLastChecked  = $window.FindName('colLastChecked')
$progressOverlay  = $window.FindName('progressOverlay')
$txtProgressTitle = $window.FindName('txtProgressTitle')
$txtProgressStep  = $window.FindName('txtProgressStep')
$btnPausePipeline = $window.FindName('btnPausePipeline')
$btnCancelPipeline = $window.FindName('btnCancelPipeline')

# =============================================================================
# Theme toggle
# =============================================================================
# Apply Dark.Steel theme explicitly at startup so the title bar gets the correct
# grey color on first render (XAML resource dict alone doesn't fully apply until
# ThemeManager touches the window).
[ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, "Dark.Steel")

# Palette and button/label theming come from SuiteCommon.
$script:WorkflowButtons = @($btnFullRun, $btnCheckLatest, $btnCheckMECM, $btnStage, $btnPackage)
$script:OptionsButtons  = @($btnOptions)

Initialize-SuiteTheme -Window $window `
    -IsDarkGetter { $toggleTheme.IsOn -eq $true } `
    -WorkflowButtons $script:WorkflowButtons `
    -OptionsButtons $script:OptionsButtons `
    -LogLabel $lblLogOutput


function Set-DialogChromeFromOwner {
    # Applies the owner's theme to a child dialog and copies title bar +
    # glow brushes across, including into the NonActive slots so the dialog
    # does not fall back to default grey when it loses focus.
    param(
        [Parameter(Mandatory)]$Dialog,
        [Parameter(Mandatory)]$Owner
    )
    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, $theme) }
    $Dialog.Owner = $Owner
    try {
        $Dialog.WindowTitleBrush          = $Owner.WindowTitleBrush
        $Dialog.NonActiveWindowTitleBrush = $Owner.WindowTitleBrush
        $Dialog.GlowBrush                 = $Owner.GlowBrush
        $Dialog.NonActiveGlowBrush        = $Owner.GlowBrush
    } catch { }
}


$toggleTheme.Add_Toggled({
    if ($toggleTheme.IsOn) {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, "Dark.Steel")
        $txtThemeLabel.Text = "Dark Theme"
        Set-ButtonTheme -IsDark $true
    }
    else {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, "Light.Blue")
        $txtThemeLabel.Text = "Light Theme"
        Set-ButtonTheme -IsDark $false
    }
    Update-TitleBarBrushes
})

# =============================================================================
# DataGrid binding + filter
# =============================================================================
$dataGrid.ItemsSource = $script:PackagerData

function Update-GridFilter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes the grid ItemsSource only.')]
    param()

    $needle = ([string]$txtGridFilter.Text).Trim().ToLowerInvariant()
    if (-not $needle) {
        # Re-bind only when a filter was active: keeping the
        # ObservableCollection itself as ItemsSource is what lets
        # Invoke-RefreshGrid's Clear()/Add() render live.
        if (-not [object]::ReferenceEquals($dataGrid.ItemsSource, $script:PackagerData)) {
            $dataGrid.ItemsSource = $script:PackagerData
        }
        return
    }
    $dataGrid.ItemsSource = @($script:PackagerData | Where-Object {
        ([string]$_.Application).ToLowerInvariant().Contains($needle) -or
        ([string]$_.Vendor).ToLowerInvariant().Contains($needle) -or
        ([string]$_.Status).ToLowerInvariant().Contains($needle) -or
        ([string]$_.CMName).ToLowerInvariant().Contains($needle)
    })
}
$txtGridFilter.Add_TextChanged({ Update-GridFilter })

# Ctrl+Click on a row opens the vendor URL
$dataGrid.Add_PreviewMouseLeftButtonUp({
    param($s, $e)
    if ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) {
        $row = $dataGrid.SelectedItem
        if ($row) {
            $url = [string]$row.VendorURL
            if (-not [string]::IsNullOrWhiteSpace($url)) {
                Start-Process $url
            }
        }
    }
})

# =============================================================================
# Context menu on DataGrid
# =============================================================================
$contextMenu = New-Object System.Windows.Controls.ContextMenu

$menuOpenLogFolder = New-Object System.Windows.Controls.MenuItem
$menuOpenLogFolder.Header = "Open Log Folder"

$menuOpenStagedFolder = New-Object System.Windows.Controls.MenuItem
$menuOpenStagedFolder.Header = "Open Staged Folder"

$menuOpenNetworkShare = New-Object System.Windows.Controls.MenuItem
$menuOpenNetworkShare.Header = "Open Network Share"

$menuSep1 = New-Object System.Windows.Controls.Separator

$menuCopyLatestVersion = New-Object System.Windows.Controls.MenuItem
$menuCopyLatestVersion.Header = "Copy Latest Version"

$contextMenu.Items.Add($menuOpenLogFolder) | Out-Null
$contextMenu.Items.Add($menuOpenStagedFolder) | Out-Null
$contextMenu.Items.Add($menuOpenNetworkShare) | Out-Null
$contextMenu.Items.Add($menuSep1) | Out-Null
$contextMenu.Items.Add($menuCopyLatestVersion) | Out-Null

$dataGrid.ContextMenu = $contextMenu

$menuOpenLogFolder.Add_Click({
    $logFolder = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path -LiteralPath $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
    }
    Start-Process "explorer.exe" -ArgumentList $logFolder
})

$menuOpenStagedFolder.Add_Click({
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    $dlRoot = $script:Prefs.DownloadRoot

    if ([string]::IsNullOrWhiteSpace($dlRoot)) {
        Add-LogLine -Message "Download Root is not set. Open Preferences to configure."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row.FullPath)
    if ($info.DownloadSubfolder) {
        $targetPath = Join-Path $dlRoot $info.DownloadSubfolder

        $version = [string]$row.LatestVersion
        if (-not [string]::IsNullOrWhiteSpace($version)) {
            $versionPath = Join-Path $targetPath $version
            if (Test-Path -LiteralPath $versionPath) {
                Start-Process "explorer.exe" -ArgumentList $versionPath
                return
            }
        }

        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList $targetPath
            return
        }
    }

    if (Test-Path -LiteralPath $dlRoot) {
        Start-Process "explorer.exe" -ArgumentList $dlRoot
    }
    else {
        Add-LogLine -Message ("Folder not found: {0}" -f $dlRoot)
    }
})

$menuOpenNetworkShare.Add_Click({
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    $fsPath = $script:Prefs.FileShareRoot

    if ([string]::IsNullOrWhiteSpace($fsPath)) {
        Add-LogLine -Message "File Share Root is not set. Open Preferences to configure."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row.FullPath)
    if ($info.VendorFolder -and $info.AppFolder) {
        $targetPath = Join-Path (Join-Path (Join-Path $fsPath "Applications") $info.VendorFolder) $info.AppFolder
        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList $targetPath
            return
        }
    }

    $appsRoot = Join-Path $fsPath "Applications"
    if (Test-Path -LiteralPath $appsRoot) {
        Start-Process "explorer.exe" -ArgumentList $appsRoot
    }
    else {
        Add-LogLine -Message ("Network path not accessible: {0}" -f $appsRoot)
    }
})

$menuCopyLatestVersion.Add_Click({
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    $version = [string]$row.LatestVersion
    if (-not [string]::IsNullOrWhiteSpace($version)) {
        [System.Windows.Clipboard]::SetText($version)
        Add-LogLine -Message ("Copied version to clipboard: {0}" -f $version)
    }
})

# =============================================================================
# Helper: enable/disable all action buttons
# =============================================================================
function Set-ActionButtonsEnabled {
    param([bool]$Enabled)
    $btnCheckLatest.IsEnabled = $Enabled
    $btnCheckMECM.IsEnabled   = $Enabled
    $btnStage.IsEnabled       = $Enabled
    $btnPackage.IsEnabled     = $Enabled
    $btnFullRun.IsEnabled     = $Enabled
    $btnOptions.IsEnabled     = $Enabled
}

$btnPausePipeline.Add_Click({
    if (-not $script:BgState -or $script:BgState.Done) { return }

    if ([bool]$script:BgState.Paused) {
        $script:BgState.Paused = $false
        $btnPausePipeline.Content = 'Pause'
        Add-LogLine -Message 'Resume requested.'
        $txtProgressStep.Text = 'Resuming...'
    }
    else {
        $script:BgState.Paused = $true
        $btnPausePipeline.Content = 'Resume'
        Add-LogLine -Message 'Pause requested. Current app will finish before the run pauses.'
        $txtProgressStep.Text = 'Pause pending...'
    }
})

$btnCancelPipeline.Add_Click({
    if (-not $script:BgState -or $script:BgState.Done) { return }

    $script:BgState.CancelRequested = $true
    $script:BgState.Paused = $false
    $btnPausePipeline.Content = 'Pause'
    $btnPausePipeline.IsEnabled = $false
    $btnCancelPipeline.IsEnabled = $false
    Add-LogLine -Message 'Cancel requested. Current app will finish before the run stops.'
    $txtProgressStep.Text = 'Cancel pending...'
})

function Get-SelectedRows {
    $selected = @()
    foreach ($item in $script:PackagerData) {
        if ($item.Selected -eq $true) { $selected += $item }
    }
    return $selected
}

# =============================================================================
# Sidebar button handlers
# =============================================================================

# --- Row selection cycle on checkbox column header click ---
# The column header is a tri-state symbol that reflects the CURRENT bulk
# selection state: empty circle = nothing selected, filled circle = all
# selected, half circle = updates only. Clicking cycles:
#   none -> all -> updates only -> none ...
# Freed three sidebar buttons worth of vertical space without losing any
# functionality. Sorting is disabled on this column only (CanUserSort="False"
# in XAML); all other columns keep their sort behavior.
#
# Unicode glyphs, all from the "Geometric Shapes" block so they share an
# em-size in Segoe UI Symbol (unlike U+25CF BLACK CIRCLE, which renders as a
# bullet dot and is too small to match the other two).
#   \u25C9 = fisheye (filled with outline),
#   \u25D0 = circle with left half black,
#   \u25CB = white circle.
$script:SelCycleSymbolAll  = [string][char]0x25C9
$script:SelCycleSymbolUpd  = [string][char]0x25D0
$script:SelCycleSymbolNone = [string][char]0x25CB
$script:SelCycleState = 0  # 0 = nothing selected (header shows empty)
                           # 1 = all selected      (header shows filled)
                           # 2 = updates only      (header shows half)

$dataGrid.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($snd, $e)
        $src = $e.OriginalSource
        if (-not ($src -is [System.Windows.Controls.Primitives.DataGridColumnHeader])) { return }
        if ($src.Column -ne $colSelected) { return }
        $e.Handled = $true

        # Commit any pending cell/row edits before mutating Selected on every
        # row and calling Items.Refresh(). Without this, if the user had just
        # toggled a checkbox individually, the DataGrid still has an open
        # edit scope on that row; the bulk mutation + Refresh tears down the
        # row mid-edit and WPF's commit state machine deadlocks.
        [void]$dataGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$dataGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)

        switch ($script:SelCycleState) {
            0 {
                # Same visible-set rule as Select-OnlyUpdateAvailable: clear
                # everywhere, select only what the filter shows.
                foreach ($item in $script:PackagerData) { $item.Selected = $false }
                foreach ($item in @($dataGrid.ItemsSource)) { $item.Selected = $true }
                Add-LogLine -Message "Selected all visible rows."
                $colSelected.Header.Text = $script:SelCycleSymbolAll
                $script:SelCycleState = 1
            }
            1 {
                Select-OnlyUpdateAvailable
                Add-LogLine -Message "Selected rows with 'Update available' status."
                $colSelected.Header.Text = $script:SelCycleSymbolUpd
                $script:SelCycleState = 2
            }
            2 {
                foreach ($item in $script:PackagerData) { $item.Selected = $false }
                Add-LogLine -Message "Deselected all rows."
                $colSelected.Header.Text = $script:SelCycleSymbolNone
                $script:SelCycleState = 0
            }
        }
        $dataGrid.Items.Refresh()
    }
)

# --- Space-bar toggles Selected on the focused row ---
# With DataGridTemplateColumn + CheckBox, the cell gets focus but the inner
# CheckBox does not receive keyboard input until tabbed/clicked into. Hook
# the DataGrid's PreviewKeyDown to toggle the Selected column's CheckBox
# when Space is pressed while a row is focused.
#
# We toggle the CheckBox's IsChecked (which drives the binding and updates
# the underlying data property via the two-way PropertyChanged binding)
# rather than mutating the pscustomobject + Items.Refresh(). The Refresh
# approach destroys the focused row/cell, breaking keyboard navigation.
$dataGrid.Add_PreviewKeyDown({
    param($snd, $e)
    if ($e.Key -ne [System.Windows.Input.Key]::Space) { return }

    # Ignore Space when a text-input control has focus (e.g., filter textbox).
    $focused = [System.Windows.Input.Keyboard]::FocusedElement
    if ($focused -is [System.Windows.Controls.TextBox]) { return }

    $row = $dataGrid.CurrentItem
    if (-not $row) { return }
    if (-not ($row.PSObject.Properties['Selected'])) { return }

    # Find the CheckBox in the Selected column's cell. GetCellContent returns
    # the root visual produced by the CellTemplate (the CheckBox itself).
    $cellContent = $colSelected.GetCellContent($row)
    if (-not $cellContent) { return }

    $checkBox = $null
    if ($cellContent -is [System.Windows.Controls.CheckBox]) {
        $checkBox = $cellContent
    }
    else {
        # Walk children if wrapped in a panel
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($cellContent)
        for ($i = 0; $i -lt $count; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($cellContent, $i)
            if ($child -is [System.Windows.Controls.CheckBox]) { $checkBox = $child; break }
        }
    }
    if (-not $checkBox) { return }

    $checkBox.IsChecked = -not [bool]$checkBox.IsChecked
    $e.Handled = $true
})

# --- Debug Columns toggle (pill at sidebar bottom) ---
$toggleDebugCols.Add_Toggled({
    $vis = if ($toggleDebugCols.IsOn) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $colCMName.Visibility      = $vis
    $colScript.Visibility      = $vis
    $colVendorURL.Visibility   = $vis
    $colLastChecked.Visibility = $vis
    Add-LogLine -Message ("Debug columns {0}." -f $(if ($toggleDebugCols.IsOn) { 'shown' } else { 'hidden' }))
})

# --- Options (single unified window) ---
$btnOptions.Add_Click({
    Show-OptionsDialog -Owner $window
})

# =============================================================================
# Dialog windows (MahApps MetroWindow versions)
# =============================================================================

# =============================================================================
# Themed message dialog (brand-cohesive replacement for System.Windows.MessageBox)
# -----------------------------------------------------------------------------
# Pass Title, Message, optional Buttons ('OK' | 'YesNo') and Icon ('Info' |
# 'Warning' | 'Error' | 'Question'). Returns 'OK' | 'Yes' | 'No' | 'Cancel'.
# Inherits the parent window's theme.
# =============================================================================

# =============================================================================
# Options dialog - single master window with left-nav + right content pattern
# (Discord / VS Code style). Replaces the four individual Show-XxxDialog
# functions. Each panel is built by a factory returning { Name, Element,
# Commit }; master OK runs every panel's Commit then Save-Preferences once.
# =============================================================================
function New-PanelStub {
    param([string]$Name, [string]$Message = 'Panel not yet migrated to the Options window.')
    $xaml = @"
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <TextBlock Text="$Message" TextWrapping="Wrap" FontSize="13" VerticalAlignment="Top" Margin="0,20,0,0"
               Foreground="{DynamicResource MahApps.Brushes.Gray3}"/>
</Grid>
"@
    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)
    return @{ Name = $Name; Element = $element; Commit = { } }
}

function New-MecmPreferencesPanel {
    $xaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="140"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Text="Site Code:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtSC" Width="80" FontSize="13" HorizontalAlignment="Left" MaxLength="5" Margin="0,0,0,8" ToolTip="ConfigMgr site code PSDrive name (e.g., MCM)"/>

    <TextBlock Grid.Row="1" Grid.Column="0" Text="Provider Machine:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtProvider" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="SMS Provider server from the ConfigMgr AdminUI connect script's ProviderMachineName value"/>

    <TextBlock Grid.Row="2" Grid.Column="0" Text="File Share Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtFS" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="UNC path to the SCCM content file share"/>

    <TextBlock Grid.Row="3" Grid.Column="0" Text="Content Layout:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Share folder layout for packaged content. Applies to future Package runs; existing content stays where it is."/>
    <ComboBox  Grid.Row="3" Grid.Column="1" x:Name="cboLayout" Width="360" FontSize="13" HorizontalAlignment="Left" Margin="0,0,0,8" ToolTip="Nested keeps an app's versions adjacent for easy retention pruning; Flat is one folder per package.">
        <ComboBoxItem Content="Nested - Applications\Vendor\App\Version"/>
        <ComboBoxItem Content="Flat - Applications\Vendor-App-Version"/>
    </ComboBox>

    <TextBlock Grid.Row="4" Grid.Column="0" Text="Download Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="4" Grid.Column="1" x:Name="txtDL" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="Local folder where installers are downloaded during staging"/>

    <TextBlock Grid.Row="5" Grid.Column="0" Text="Est. Runtime:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="5" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBox x:Name="txtEst" Width="60" FontSize="13" MaxLength="4" ToolTip="Estimated install runtime in minutes"/>
        <TextBlock Text=" mins" FontSize="13" VerticalAlignment="Center" Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
    </StackPanel>

    <TextBlock Grid.Row="6" Grid.Column="0" Text="Max Runtime:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="6" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBox x:Name="txtMax" Width="60" FontSize="13" MaxLength="4" ToolTip="Maximum allowed install runtime in minutes"/>
        <TextBlock Text=" mins" FontSize="13" VerticalAlignment="Center" Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
    </StackPanel>

    <TextBlock Grid.Row="7" Grid.Column="0" Text="Auto-distribute:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="When enabled, the Package phase calls Start-CMContentDistribution after creating each MECM Application."/>
    <CheckBox  Grid.Row="7" Grid.Column="1" x:Name="chkAutoDist" Content="Start-CMContentDistribution after Package" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="8" Grid.Column="0" Text="DP Group:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Exact name of the Distribution Point Group to target."/>
    <TextBox   Grid.Row="8" Grid.Column="1" x:Name="txtDPGroup" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="Distribution Point Group display name (e.g. 'All DPs')"/>

    <TextBlock Grid.Row="9" Grid.Column="0" Text="Test deployment:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Requires Auto-distribute enabled and a DP Group name. After content distribution, deploys the application (Available, immediately, default options) to the test collection."/>
    <CheckBox  Grid.Row="9" Grid.Column="1" x:Name="chkTestDeploy" Content="Deploy to test collection after distribution" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="10" Grid.Column="0" Text="Test collection:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Exact device collection name that receives the Available test deployment."/>
    <TextBox   Grid.Row="10" Grid.Column="1" x:Name="txtTestCollection" FontSize="13" MaxLength="255" Margin="0,0,0,8" ToolTip="Device collection display name (e.g. 'App Test Devices')"/>

    <TextBlock Grid.Row="11" Grid.Column="0" Text="" Margin="0,0,0,8"/>
    <CheckBox  Grid.Row="11" Grid.Column="1" x:Name="chkCreateTestColl" Content="Create collection if it does not exist" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal" ToolTip="Creates an empty direct-membership device collection limited to All Systems when the named collection is missing."/>

    <TextBlock Grid.Row="12" Grid.Column="0" Text="Console:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="Configuration Manager Console (AdminUI) detection status. Checked once per launch."/>
    <TextBlock Grid.Row="12" Grid.Column="1" x:Name="txtConsoleStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,6,0,0"/>

    <TextBlock Grid.Row="13" Grid.Column="0" Text="7-Zip CLI:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="7-Zip command-line (7z.exe) detection status. Required by Adobe Reader + TeamViewer Host packagers."/>
    <TextBlock Grid.Row="13" Grid.Column="1" x:Name="txtSevenZipStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,6,0,0"/>

    <TextBlock Grid.Row="14" Grid.Column="0" Text="Content Prep:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) detection status. Downloaded on first use, or place the exe on PATH."/>
    <StackPanel Grid.Row="14" Grid.Column="1" Orientation="Horizontal" Margin="0,6,0,0">
        <TextBlock x:Name="txtIntuneWinStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/>
        <Button x:Name="btnIntuneWinDownload" Content="Download" FontSize="11" Margin="10,0,0,0" Padding="10,2" Visibility="Collapsed"/>
    </StackPanel>

    <TextBlock Grid.Row="15" Grid.Column="0" Text="Intunewin:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="When enabled, a successful Package also produces an .intunewin from the staged content and stores it beside the network content version folder."/>
    <CheckBox  Grid.Row="15" Grid.Column="1" x:Name="chkIntuneWin" Content="Create .intunewin during Package" FontSize="13" VerticalAlignment="Center" Margin="0,6,0,0" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
</Grid>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $txtSC  = $element.FindName('txtSC')
    $txtProvider = $element.FindName('txtProvider')
    $txtFS  = $element.FindName('txtFS')
    $cboLayout = $element.FindName('cboLayout')
    $txtDL  = $element.FindName('txtDL')
    $txtEst = $element.FindName('txtEst')
    $txtMax = $element.FindName('txtMax')
    $chkAutoDist = $element.FindName('chkAutoDist')
    $txtDPGroup  = $element.FindName('txtDPGroup')
    $chkTestDeploy     = $element.FindName('chkTestDeploy')
    $txtTestCollection = $element.FindName('txtTestCollection')
    $chkCreateTestColl = $element.FindName('chkCreateTestColl')
    $txtConsoleStatus  = $element.FindName('txtConsoleStatus')
    $txtSevenZipStatus = $element.FindName('txtSevenZipStatus')
    $txtIntuneWinStatus   = $element.FindName('txtIntuneWinStatus')
    $btnIntuneWinDownload = $element.FindName('btnIntuneWinDownload')
    $chkIntuneWin         = $element.FindName('chkIntuneWin')

    $txtSC.Text  = [string]$script:Prefs.SiteCode
    $txtProvider.Text = [string]$script:Prefs.ProviderMachineName
    $txtFS.Text  = [string]$script:Prefs.FileShareRoot
    $cboLayout.SelectedIndex = if ([string]$script:Prefs.ContentLayout -eq 'Flat') { 1 } else { 0 }
    $txtDL.Text  = [string]$script:Prefs.DownloadRoot
    $txtEst.Text = [string]$script:Prefs.EstimatedRuntimeMins
    $txtMax.Text = [string]$script:Prefs.MaximumRuntimeMins
    $chkAutoDist.IsChecked = [bool]$script:Prefs.ContentDistribution.AutoDistribute
    $txtDPGroup.Text       = [string]$script:Prefs.ContentDistribution.DPGroupName
    $chkTestDeploy.IsChecked     = [bool]$script:Prefs.ContentDistribution.DeployToTestCollection
    $txtTestCollection.Text      = [string]$script:Prefs.ContentDistribution.TestCollectionName
    $chkCreateTestColl.IsChecked = [bool]$script:Prefs.ContentDistribution.CreateTestCollectionIfMissing

    # Test-deployment controls require auto-distribute + DP group: the
    # deployment only runs after successful content distribution.
    $updateTestDeployState = {
        $distReady = [bool]$chkAutoDist.IsChecked -and -not [string]::IsNullOrWhiteSpace($txtDPGroup.Text)
        $chkTestDeploy.IsEnabled     = $distReady
        $txtTestCollection.IsEnabled = $distReady -and [bool]$chkTestDeploy.IsChecked
        $chkCreateTestColl.IsEnabled = $distReady -and [bool]$chkTestDeploy.IsChecked
    }.GetNewClosure()
    & $updateTestDeployState
    $chkAutoDist.Add_Click($updateTestDeployState)
    $chkAutoDist.Add_Unchecked($updateTestDeployState)
    $txtDPGroup.Add_TextChanged($updateTestDeployState)
    $chkTestDeploy.Add_Click($updateTestDeployState)

    $cm = $script:Prefs.DetectedTools.ConfigMgrConsole
    if ($cm -and $cm.Found) {
        $txtConsoleStatus.Text = ([char]0x2713 + " Detected  -  {0} v{1}" -f $cm.DisplayName, $cm.DisplayVersion)
        $txtConsoleStatus.ToolTip = ("Module: {0}" -f $cm.ModulePath)
    } else {
        $txtConsoleStatus.Text = ([char]0x2717 + " Not detected  -  install the ConfigMgr Console (AdminUI) and reboot")
        $txtConsoleStatus.ToolTip = "Detected once per launch via registry ARP + SMS_ADMIN_UI_PATH + well-known install paths"
    }

    $sz = $script:Prefs.DetectedTools.SevenZipCli
    if ($sz -and $sz.Found) {
        $txtSevenZipStatus.Text = ([char]0x2713 + " Detected  -  {0} v{1}" -f $sz.DisplayName, $sz.DisplayVersion)
        $txtSevenZipStatus.ToolTip = ("7z.exe: {0}" -f $sz.ExePath)
    } else {
        $txtSevenZipStatus.Text = ([char]0x2717 + " Not detected  -  Adobe Reader + TeamViewer Host packagers need 7-Zip CLI")
        $txtSevenZipStatus.ToolTip = "Detected once per launch via registry ARP + Program Files\7-Zip"
    }

    $chkIntuneWin.IsChecked = [bool]$script:Prefs.Intune.CreateIntuneWin
    $prefsRefIw = $script:Prefs
    $updateIntuneWinState = {
        $iw = $prefsRefIw.DetectedTools.IntuneWinAppUtil
        if ($iw -and $iw.Found) {
            $verText = if ([string]::IsNullOrWhiteSpace([string]$iw.DisplayVersion)) { '' } else { (" v{0}" -f $iw.DisplayVersion) }
            $txtIntuneWinStatus.Text = ([char]0x2713 + " Detected  -  IntuneWinAppUtil{0}" -f $verText)
            $txtIntuneWinStatus.ToolTip = ("IntuneWinAppUtil.exe: {0}" -f $iw.ExePath)
            $btnIntuneWinDownload.Visibility = 'Collapsed'
            $chkIntuneWin.IsEnabled = $true
        } else {
            $txtIntuneWinStatus.Text = ([char]0x2717 + " Not detected  -  download it here or place IntuneWinAppUtil.exe on PATH")
            $txtIntuneWinStatus.ToolTip = "Checked once per launch: preferences path, LOCALAPPDATA tool cache, PATH"
            $btnIntuneWinDownload.Visibility = 'Visible'
            $chkIntuneWin.IsEnabled = $false
        }
    }.GetNewClosure()
    & $updateIntuneWinState

    $btnIntuneWinDownload.Add_Click({
        $btnIntuneWinDownload.IsEnabled = $false
        $txtIntuneWinStatus.Text = "Downloading Microsoft Win32 Content Prep Tool..."
        $downloadOk = $false
        try {
            $exePath = Install-IntuneWinAppUtil -DestinationFolder (Get-IntuneWinToolCachePath)
            $prefsRefIw.DetectedTools.IntuneWinAppUtil = Invoke-DetectIntuneWinAppUtil -KnownPath $exePath
            Save-Preferences -Prefs $prefsRefIw
            $downloadOk = $true
        } catch {
            $txtIntuneWinStatus.Text = ([char]0x2717 + " Download failed  -  {0}" -f $_.Exception.Message)
        } finally {
            $btnIntuneWinDownload.IsEnabled = $true
        }
        if ($downloadOk) { & $updateIntuneWinState }
    }.GetNewClosure())

    # Closure captures panel-local controls by value. Prefs ref is captured too
    # so the commit writes to the live $script:Prefs without needing $script:
    # scope resolution from inside the closure (which can be unreliable).
    $prefsRef = $script:Prefs
    $commit = {
        $estVal = 15; $maxVal = 30
        if (-not [int]::TryParse($txtEst.Text.Trim(), [ref]$estVal)) { $estVal = 15 }
        if (-not [int]::TryParse($txtMax.Text.Trim(), [ref]$maxVal)) { $maxVal = 30 }

        $prefsRef.SiteCode             = $txtSC.Text.Trim()
        $prefsRef.ProviderMachineName  = $txtProvider.Text.Trim()
        $prefsRef.FileShareRoot        = $txtFS.Text.Trim()
        $prefsRef.ContentLayout        = if ($cboLayout.SelectedIndex -eq 1) { 'Flat' } else { 'Nested' }
        $prefsRef.DownloadRoot         = $txtDL.Text.Trim()
        $prefsRef.EstimatedRuntimeMins = $estVal
        $prefsRef.MaximumRuntimeMins   = $maxVal
        $prefsRef.ContentDistribution.AutoDistribute = [bool]$chkAutoDist.IsChecked
        $prefsRef.ContentDistribution.DPGroupName    = $txtDPGroup.Text.Trim()
        $prefsRef.ContentDistribution.DeployToTestCollection        = [bool]$chkTestDeploy.IsChecked
        $prefsRef.ContentDistribution.TestCollectionName            = $txtTestCollection.Text.Trim()
        $prefsRef.ContentDistribution.CreateTestCollectionIfMissing = [bool]$chkCreateTestColl.IsChecked
        $prefsRef.Intune.CreateIntuneWin = [bool]$chkIntuneWin.IsChecked
    }.GetNewClosure()

    return @{ Name = 'MECM Preferences'; Element = $element; Commit = $commit }
}

function New-AppFlowPanel {
    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" FontSize="12"
               Foreground="{DynamicResource MahApps.Brushes.Gray3}" Margin="0,0,0,12"
               Text="One Click runs a Check (and optionally Stage / Package) against the apps you track here. Apps are skipped when the last check is still within their cadence unless you enable Force on launch."/>
    <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Action on update:" VerticalAlignment="Center" FontSize="12" Margin="0,0,8,0"/>
        <ComboBox Grid.Column="1" x:Name="cboAction" Width="190" VerticalAlignment="Center">
            <ComboBoxItem Content="Report only"/>
            <ComboBoxItem Content="Stage"/>
            <ComboBoxItem Content="Stage and Package"/>
        </ComboBox>
        <Controls:ToggleSwitch Grid.Column="3" x:Name="toggleForce" IsOn="False"
                                Header="Force on launch (ignore cadence)"
                                OnContent="" OffContent="" MinWidth="0"
                                VerticalAlignment="Center"/>
    </Grid>
    <DataGrid x:Name="dgApps" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
              GridLinesVisibility="Horizontal" HeadersVisibility="Column" RowHeaderWidth="0" BorderThickness="0"
              IsTextSearchEnabled="True" TextSearch.TextPath="Application">
        <DataGrid.Columns>
            <DataGridTemplateColumn Header="Track" Width="56" CanUserSort="True" SortMemberPath="Tracked">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <CheckBox IsChecked="{Binding Tracked, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}"
                                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn Header="Application" Width="*" Binding="{Binding Application}" IsReadOnly="True"/>
            <DataGridTextColumn Header="Vendor" Width="160" Binding="{Binding Vendor}" IsReadOnly="True"/>
            <DataGridTextColumn Header="Cadence (days)" Width="120" Binding="{Binding CadenceDisplay, UpdateSourceTrigger=LostFocus, Mode=TwoWay}"/>
        </DataGrid.Columns>
    </DataGrid>
</DockPanel>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $cboAction   = $element.FindName('cboAction')
    $toggleForce = $element.FindName('toggleForce')
    $dgApps      = $element.FindName('dgApps')

    $currentPrefs = $script:Prefs.AppFlow
    $trackedSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($currentPrefs.Tracked),
        [System.StringComparer]::OrdinalIgnoreCase)

    $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $packagers = Get-Packagers -Root $PackagersRoot | Sort-Object Vendor, Application
    foreach ($p in $packagers) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($p.Script)
        $headerDays = $null
        if ($p.UpdateCadenceDays) { $headerDays = [int]$p.UpdateCadenceDays }
        $effective = 7
        if ($headerDays) { $effective = $headerDays }
        $overrideProp = $null
        if ($currentPrefs.CadenceOverrides) {
            $overrideProp = $currentPrefs.CadenceOverrides.PSObject.Properties[$base]
        }
        if ($overrideProp) { $effective = [int]$overrideProp.Value }

        $rows.Add([pscustomobject]@{
            Packager       = $base
            Application    = $p.Application
            Vendor         = $p.Vendor
            Tracked        = $trackedSet.Contains($base)
            CadenceDisplay = [string]$effective
            HeaderDays     = $headerDays
        })
    }

    $dgApps.ItemsSource = $rows
    $cboAction.SelectedIndex = switch ($currentPrefs.Action) {
        'Report'          { 0 }
        'Stage'           { 1 }
        'StageAndPackage' { 2 }
        default           { 0 }
    }
    $toggleForce.IsOn = [bool]$currentPrefs.ForceOnLaunch

    $prefsRef = $script:Prefs
    $commit = {
        [void]$dgApps.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$dgApps.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,  $true)

        $newTracked = @($rows | Where-Object { $_.Tracked } | ForEach-Object { $_.Packager })
        $newAction = switch ($cboAction.SelectedIndex) {
            0 { 'Report' }
            1 { 'Stage' }
            2 { 'StageAndPackage' }
            default { 'Report' }
        }

        $overrideProps = [ordered]@{}
        foreach ($row in $rows) {
            $parsed = 0
            if (-not [int]::TryParse([string]$row.CadenceDisplay, [ref]$parsed)) { continue }
            if ($parsed -lt 1) { continue }
            $headerDefault = if ($row.HeaderDays) { [int]$row.HeaderDays } else { 7 }
            if ($parsed -ne $headerDefault) { $overrideProps[$row.Packager] = $parsed }
        }

        $prefsRef.AppFlow.Tracked          = $newTracked
        $prefsRef.AppFlow.Action           = $newAction
        $prefsRef.AppFlow.CadenceOverrides = [pscustomobject]$overrideProps
        $prefsRef.AppFlow.ForceOnLaunch    = [bool]$toggleForce.IsOn
    }.GetNewClosure()

    return @{ Name = 'One Click Settings'; Element = $element; Commit = $commit }
}

function Show-PreviewDialog {
    # Themed read-only preview window (monospaced, scrollable, Copy / Close).
    # Used by the Packager Preferences panel's CWA and M365 preview buttons.
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Content,
        [int]$Width = 780,
        [int]$Height = 500
    )
    $xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="$Title"
    Width="$Width" Height="$Height"
    MinWidth="480" MinHeight="260"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <DockPanel Margin="12">
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnCopy"  Content="Copy"  MinWidth="90" Height="32" Margin="0,0,8,0" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnClose" Content="Close" MinWidth="90" Height="32" IsDefault="True" IsCancel="True" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
        <TextBox x:Name="txtContent"
                 IsReadOnly="True"
                 TextWrapping="NoWrap"
                 AcceptsReturn="True"
                 VerticalScrollBarVisibility="Auto"
                 HorizontalScrollBarVisibility="Auto"
                 FontFamily="Cascadia Code, Consolas, Courier New"
                 FontSize="11"/>
    </DockPanel>
</Controls:MetroWindow>
"@
    [xml]$xmlDoc = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $pvWin = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $pvWin
    Set-DialogChromeFromOwner -Dialog $pvWin -Owner $Owner

    $txt   = $pvWin.FindName('txtContent')
    $copy  = $pvWin.FindName('btnCopy')
    $close = $pvWin.FindName('btnClose')
    $txt.Text = $Content

    $copy.Add_Click({
        try { [System.Windows.Clipboard]::SetText($txt.Text) } catch { }
    }.GetNewClosure())
    $close.Add_Click({ $pvWin.Close() }.GetNewClosure())

    [void]$pvWin.ShowDialog()
}

function New-ProductFilterPanel {
    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Select which applications appear in the main grid. Uncheck to hide."
                   FontSize="12" Foreground="{DynamicResource MahApps.Brushes.Gray3}" VerticalAlignment="Center" TextWrapping="Wrap"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Button x:Name="btnSelAll"  Content="Select All"  MinWidth="90" Height="28" Margin="6,0,0,0" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnSelNone" Content="Select None" MinWidth="90" Height="28" Margin="6,0,0,0" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </Grid>
    <TreeView x:Name="treeApps" />
</DockPanel>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $treeApps  = $element.FindName('treeApps')
    $btnSelAll = $element.FindName('btnSelAll')
    $btnSelNone= $element.FindName('btnSelNone')

    $hiddenSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($script:Prefs.HiddenApplications),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $checkboxes = @{}
    $items = Get-Packagers -Root $PackagersRoot
    $vendors = $items | Group-Object Vendor | Sort-Object Name

    foreach ($group in $vendors) {
        $vendorItem = New-Object System.Windows.Controls.TreeViewItem
        $vendorCB = New-Object System.Windows.Controls.CheckBox
        $vendorCB.Content = if ($group.Name) { $group.Name } else { "(No Vendor)" }
        $vendorCB.FontWeight = [System.Windows.FontWeights]::Bold
        $vendorItem.Header = $vendorCB

        $allChecked = $true
        foreach ($app in ($group.Group | Sort-Object Application)) {
            $appItem = New-Object System.Windows.Controls.TreeViewItem
            $appCB = New-Object System.Windows.Controls.CheckBox
            $appCB.Content = $app.Application
            $appCB.Tag = $app.Script
            $isHidden = $hiddenSet.Contains($app.Script)
            $appCB.IsChecked = (-not $isHidden)
            if ($isHidden) { $allChecked = $false }
            $appItem.Header = $appCB
            [void]$vendorItem.Items.Add($appItem)
            $checkboxes[$app.Script] = $appCB
        }

        $vendorCB.IsChecked = $allChecked
        $vendorCB.Tag = $vendorItem

        $vendorCB.Add_Checked({
            param($s, $e)
            $vi = $s.Tag
            foreach ($child in $vi.Items) { $child.Header.IsChecked = $true }
        })
        $vendorCB.Add_Unchecked({
            param($s, $e)
            $vi = $s.Tag
            foreach ($child in $vi.Items) { $child.Header.IsChecked = $false }
        })

        $vendorItem.IsExpanded = $true
        [void]$treeApps.Items.Add($vendorItem)
    }

    $btnSelAll.Add_Click({
        foreach ($kv in $checkboxes.GetEnumerator()) { $kv.Value.IsChecked = $true }
        foreach ($vi in $treeApps.Items) { $vi.Header.IsChecked = $true }
    }.GetNewClosure())

    $btnSelNone.Add_Click({
        foreach ($kv in $checkboxes.GetEnumerator()) { $kv.Value.IsChecked = $false }
        foreach ($vi in $treeApps.Items) { $vi.Header.IsChecked = $false }
    }.GetNewClosure())

    $prefsRef = $script:Prefs
    $commit = {
        $hidden = New-Object System.Collections.Generic.List[string]
        foreach ($kv in $checkboxes.GetEnumerator()) {
            if ($kv.Value.IsChecked -ne $true) {
                $hidden.Add([string]$kv.Key)
            }
        }
        $prefsRef.HiddenApplications = $hidden.ToArray()
    }.GetNewClosure()

    return @{ Name = 'Product Filter'; Element = $element; Commit = $commit }
}

function New-PackagerPreferencesPanel {
    $sw = Read-CwaSwitches
    $tv = Read-TvHostConfig
    $ssms = $script:Prefs.SSMSInstallOptions

    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,4">
        <Button x:Name="btnM365Preview" Content="M365 Preview" MinWidth="120" Height="30" Margin="0,0,8,0"
                Style="{DynamicResource MahApps.Styles.Button.Square}"
                Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        <Button x:Name="btnCwaPreview"  Content="CWA Preview"  MinWidth="120" Height="30"
                Style="{DynamicResource MahApps.Styles.Button.Square}"
                Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
    </StackPanel>
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
        <StackPanel x:Name="panelContent" Margin="0,0,4,0"/>
    </ScrollViewer>
</DockPanel>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $panelContent   = $element.FindName('panelContent')
    $btnCwaPreview  = $element.FindName('btnCwaPreview')
    $btnM365Preview = $element.FindName('btnM365Preview')

    # --- Helpers (local to factory; close over $panelContent) ---
    $addHeader = {
        param([string]$Text)
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $Text
        $tb.FontSize = 13
        $tb.FontWeight = [System.Windows.FontWeights]::Bold
        $tb.Margin = New-Object System.Windows.Thickness(0, 14, 0, 6)
        [void]$panelContent.Children.Add($tb)
    }
    $addDivider = {
        $div = New-Object System.Windows.Controls.Border
        $div.Height = 1
        $div.Margin = New-Object System.Windows.Thickness(0, 16, 0, 0)
        $div.SetResourceReference(
            [System.Windows.Controls.Border]::BackgroundProperty,
            'MahApps.Brushes.Control.Border'
        )
        [void]$panelContent.Children.Add($div)
    }
    $addLabelRow = {
        param([string]$Label, [System.Windows.UIElement]$Control, [string]$Tooltip = '')
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $sp.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $Label
        $lbl.Width = 130
        $lbl.FontSize = 13
        $lbl.FontWeight = [System.Windows.FontWeights]::Bold
        $lbl.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$sp.Children.Add($lbl)
        if ($Tooltip) { $Control.ToolTip = $Tooltip }
        [void]$sp.Children.Add($Control)
        [void]$panelContent.Children.Add($sp)
    }
    $addCheckBox = {
        param([string]$Text, [bool]$Checked, [string]$Tooltip = '', [double]$LeftMargin = 0)
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $Text
        $cb.FontSize = 12
        $cb.IsChecked = $Checked
        $cb.Margin = New-Object System.Windows.Thickness($LeftMargin, 2, 0, 2)
        if ($Tooltip) { $cb.ToolTip = $Tooltip }
        [void]$panelContent.Children.Add($cb)
        return $cb
    }

    # =============================================
    # M365: ODT SETTINGS
    # =============================================
    & $addHeader "M365: ODT Settings"

    $txtCN = New-Object System.Windows.Controls.TextBox
    $txtCN.Text = $script:Prefs.CompanyName
    $txtCN.FontSize = 13
    $txtCN.MaxLength = 100
    $txtCN.Width = 250
    $txtCN.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Company Name:" $txtCN "Organization name embedded in Office deployment XML and other packager configs"

    $cmbCH = New-Object System.Windows.Controls.ComboBox
    $cmbCH.FontSize = 13
    $cmbCH.Width = 220
    $cmbCH.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    [void]$cmbCH.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{Content='Monthly Enterprise Channel'}))
    [void]$cmbCH.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{Content='Current Channel'}))
    $channelDisplayMap = @{ 'MonthlyEnterprise' = 'Monthly Enterprise Channel'; 'Current' = 'Current Channel' }
    $currentDisplay = $channelDisplayMap[$script:Prefs.M365Channel]
    if (-not $currentDisplay) { $currentDisplay = 'Monthly Enterprise Channel' }
    foreach ($item in $cmbCH.Items) { if ($item.Content -eq $currentDisplay) { $cmbCH.SelectedItem = $item; break } }
    & $addLabelRow "M365 Channel:" $cmbCH "Office 365 update channel for M365 Apps, Project, and Visio packagers"

    $cmbDM = New-Object System.Windows.Controls.ComboBox
    $cmbDM.FontSize = 13
    $cmbDM.Width = 220
    $cmbDM.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    [void]$cmbDM.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{Content='Managed (Offline)'}))
    [void]$cmbDM.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{Content='Online (CDN)'}))
    $deployModeDisplayMap = @{ 'Managed' = 'Managed (Offline)'; 'Online' = 'Online (CDN)' }
    $currentDM = $deployModeDisplayMap[$script:Prefs.M365DeployMode]
    if (-not $currentDM) { $currentDM = 'Managed (Offline)' }
    foreach ($item in $cmbDM.Items) { if ($item.Content -eq $currentDM) { $cmbDM.SelectedItem = $item; break } }
    & $addLabelRow "M365 Deploy Mode:" $cmbDM "Managed: download Office source (~2.3 GB), pin version. Online: CDN-direct install, always latest."

    # --- Exclude apps (ExcludeApp IDs injected into ODT XML) ---
    $tbExcl = New-Object System.Windows.Controls.TextBlock
    $tbExcl.Text = "Exclude apps from install:"
    $tbExcl.FontSize = 13
    $tbExcl.FontWeight = [System.Windows.FontWeights]::Bold
    $tbExcl.Margin = New-Object System.Windows.Thickness(0, 6, 0, 4)
    [void]$panelContent.Children.Add($tbExcl)

    $exclGrid = New-Object System.Windows.Controls.Grid
    $exclGrid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
    $ec1 = New-Object System.Windows.Controls.ColumnDefinition; $ec1.Width = [System.Windows.GridLength]::Auto
    $ec2 = New-Object System.Windows.Controls.ColumnDefinition; $ec2.Width = [System.Windows.GridLength]::Auto
    [void]$exclGrid.ColumnDefinitions.Add($ec1)
    [void]$exclGrid.ColumnDefinitions.Add($ec2)

    $excludeDefs = @(
        @{Id='Access';            Label='Access';                          Tip="Exclude Microsoft Access from install. Safe to exclude in environments that don't use Access databases."}
        @{Id='Excel';             Label='Excel';                           Tip="Exclude Excel. Rarely used in real deployments; excluding Excel usually breaks user expectations."}
        @{Id='Groove';            Label='OneDrive for Business (Groove)';  Tip="Exclude the legacy OneDrive for Business sync client (ExcludeApp ID 'Groove'). ODT docs: 'For OneDrive, use Groove.' Recommended exclude - the modern OneDrive client is a separate install."}
        @{Id='Lync';              Label='Skype for Business (Lync)';       Tip="Exclude Skype for Business (ExcludeApp ID 'Lync'). Skype for Business Online retired 2021; almost always safe to exclude."}
        @{Id='OneDrive';          Label='OneDrive (modern)';               Tip="Exclude the modern per-user OneDrive client that Office auto-installs. Exclude if you deploy OneDrive separately (Intune, machine-wide installer, etc)."}
        @{Id='OneNote';           Label='OneNote';                         Tip="Exclude OneNote. Most orgs keep OneNote installed."}
        @{Id='Outlook';           Label='Outlook (classic)';               Tip="Exclude classic Outlook. Rarely excluded."}
        @{Id='OutlookForWindows'; Label='Outlook for Windows (new)';       Tip="Exclude the new Outlook for Windows app that Office 365 installs alongside classic Outlook. Typical exclude until users have migrated."}
        @{Id='PowerPoint';        Label='PowerPoint';                      Tip="Exclude PowerPoint. Rarely excluded."}
        @{Id='Publisher';         Label='Publisher';                       Tip="Exclude Publisher. Publisher support ends October 2026; good candidate to exclude in new deployments."}
        @{Id='Teams';             Label='Teams';                           Tip="Exclude the auto-bundled Teams installer. Recommended exclude when deploying Teams via Intune or machine-wide MSI separately."}
        @{Id='Word';              Label='Word';                            Tip="Exclude Word. Rarely excluded."}
        @{Id='Bing';              Label='Microsoft Search in Bing';        Tip="Exclude the Microsoft Search in Bing browser extension (ExcludeApp ID 'Bing'). Not in current ODT docs but historically accepted."}
    )

    $excludeCBs = @{}
    $currentExcludes = @($script:Prefs.M365ExcludeApps)
    for ($i = 0; $i -lt $excludeDefs.Count; $i++) {
        $def = $excludeDefs[$i]
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $def.Label
        $cb.FontSize = 12
        $cb.IsChecked = ($currentExcludes -contains $def.Id)
        $cb.ToolTip = $def.Tip
        $cb.Margin = New-Object System.Windows.Thickness(0, 2, 18, 2)
        $col = $i % 2
        $row = [int]([math]::Floor($i / 2))
        while ($exclGrid.RowDefinitions.Count -le $row) {
            $rd = New-Object System.Windows.Controls.RowDefinition
            $rd.Height = [System.Windows.GridLength]::Auto
            [void]$exclGrid.RowDefinitions.Add($rd)
        }
        [System.Windows.Controls.Grid]::SetColumn($cb, $col)
        [System.Windows.Controls.Grid]::SetRow($cb, $row)
        [void]$exclGrid.Children.Add($cb)
        $excludeCBs[$def.Id] = $cb
    }
    [void]$panelContent.Children.Add($exclGrid)

    # =============================================
    # SSMS: SILENT INSTALL OPTIONS
    # =============================================
    & $addDivider
    & $addHeader "SSMS: Silent Install Options"

    $cmbSsmsUiMode = New-Object System.Windows.Controls.ComboBox
    $cmbSsmsUiMode.FontSize = 13
    $cmbSsmsUiMode.Width = 120
    $cmbSsmsUiMode.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    foreach ($val in @("Quiet", "Passive")) { [void]$cmbSsmsUiMode.Items.Add($val) }
    $cmbSsmsUiMode.SelectedItem = if ($ssms.UIMode -in @('Quiet','Passive')) { $ssms.UIMode } else { 'Quiet' }
    & $addLabelRow "UI Mode:" $cmbSsmsUiMode "Quiet adds --quiet for a fully hidden install. Passive adds --passive for progress-only UI and is less suitable for required MECM deployments."

    $txtSsmsInstallPath = New-Object System.Windows.Controls.TextBox
    $txtSsmsInstallPath.Text = [string]$ssms.InstallPath
    $txtSsmsInstallPath.FontSize = 13
    $txtSsmsInstallPath.Width = 350
    $txtSsmsInstallPath.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Install Path:" $txtSsmsInstallPath "Optional --installPath value. Leave blank for Microsoft's default SSMS 22 path. If set, the same path is used for detection and uninstall."

    $chkSsmsDownloadThenInstall = & $addCheckBox "Download all packages before install (--downloadThenInstall)" ([bool]$ssms.DownloadThenInstall) "Forces SSMS setup to download required packages before starting installation. Mutually exclusive with --installWhileDownloading, which is the Microsoft default."
    $chkSsmsNoUpdateInstaller  = & $addCheckBox "Do not update Visual Studio Installer (--noUpdateInstaller)" ([bool]$ssms.NoUpdateInstaller) "Prevents installer self-update when quiet is specified. Microsoft documents that setup can fail if an installer update is required."
    $chkSsmsRecommended        = & $addCheckBox "Include recommended components (--includeRecommended)" ([bool]$ssms.IncludeRecommended) "Adds recommended components for selected SSMS workloads. Leave off for the lean default SSMS install."
    $chkSsmsOptional           = & $addCheckBox "Include optional components (--includeOptional)" ([bool]$ssms.IncludeOptional) "Adds optional components for selected SSMS workloads. This can increase install size and duration."
    $chkSsmsRemoveOos          = & $addCheckBox "Remove out-of-support components (--removeOos true)" ([bool]$ssms.RemoveOos) "Tells the installer to remove components that have transitioned out of support during this install or update."
    $chkSsmsForceClose         = & $addCheckBox "Force close SSMS if in use (--force)" ([bool]$ssms.ForceClose) "Allows setup to close running SSMS processes. This can cause loss of unsaved query windows, so use deliberately."

    # =============================================
    # TEAMVIEWER HOST
    # =============================================
    & $addDivider
    & $addHeader "TeamViewer Host"

    $txtTvApiToken = New-Object System.Windows.Controls.TextBox
    $txtTvApiToken.Text = $tv.ApiToken
    $txtTvApiToken.FontSize = 13
    $txtTvApiToken.Width = 350
    $txtTvApiToken.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "API Token:" $txtTvApiToken "TeamViewer script token that authorizes automatic device assignment. Management Console -> Company Administration -> Advanced -> Create script token. Leave blank to skip auto-assignment."

    $txtTvConfigId = New-Object System.Windows.Controls.TextBox
    $txtTvConfigId.Text = $tv.CustomConfigId
    $txtTvConfigId.FontSize = 13
    $txtTvConfigId.Width = 250
    $txtTvConfigId.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Custom Config ID:" $txtTvConfigId "Identifier of a custom Host module from Management Console -> Design & Deploy. Leave blank for default Host."

    $txtTvAssignOpts = New-Object System.Windows.Controls.TextBox
    $txtTvAssignOpts.Text = $tv.AssignmentOptions
    $txtTvAssignOpts.FontSize = 13
    $txtTvAssignOpts.Width = 350
    $txtTvAssignOpts.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Assignment Options:" $txtTvAssignOpts "Quoted string of flags passed during enrollment (--grant-easy-access, --alias %COMPUTERNAME%, --reassign, --group <name>). Passed to msiexec as ASSIGNMENTOPTIONS=`"...`""

    $chkTvRemoveShortcut = & $addCheckBox "Remove desktop shortcut after install" ([bool]$tv.RemoveDesktopShortcut) "Adds REMOVE=f.DesktopShortcut to the msiexec command so the install does not place a TeamViewer shortcut on the desktop."

    # =============================================
    # CWA: STORE CONFIGURATION
    # =============================================
    & $addDivider
    & $addHeader "CWA: Store Configuration"

    $txtStoreName = New-Object System.Windows.Controls.TextBox
    $txtStoreName.Text = $sw.Store.Name
    $txtStoreName.FontSize = 13
    $txtStoreName.Width = 200
    $txtStoreName.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Store Name:" $txtStoreName "Friendly name for the StoreFront store (STORE0 parameter)"

    $txtStoreUrl = New-Object System.Windows.Controls.TextBox
    $txtStoreUrl.Text = $sw.Store.Url
    $txtStoreUrl.FontSize = 13
    $txtStoreUrl.Width = 350
    $txtStoreUrl.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    & $addLabelRow "Store URL:" $txtStoreUrl "StoreFront base URL (e.g. https://storefront.company.com/Citrix/Store). /discovery is appended automatically."

    # =============================================
    # CWA: INSTALLATION OPTIONS
    # =============================================
    & $addHeader "CWA: Installation Options"

    $chkClean     = & $addCheckBox "Clean Install (/CleanInstall)" ([bool]$sw.Installation.CleanInstall) "Removes leftover configuration and registry data from any prior installation before installing."
    $chkSSOn      = & $addCheckBox "Single Sign-On (/includeSSON + ENABLE_SSON)" ([bool]$sw.Installation.IncludeSSON) "Installs the SSO component and activates domain pass-through authentication."
    $chkAppProt   = & $addCheckBox "App Protection (/includeappprotection)" ([bool]$sw.Installation.AppProtection) "Installs anti-keylogging and anti-screen capture protection for Citrix sessions."
    $chkPreLaunch = & $addCheckBox "Session Pre-Launch (ENABLEPRELAUNCH)" ([bool]$sw.Installation.SessionPreLaunch) "Pre-launches a Citrix session at logon for faster application startup."
    $chkSelfSvc   = & $addCheckBox "Self-Service Mode (SELFSERVICEMODE)" ([bool]$sw.Installation.SelfServiceMode) "Shows the Citrix Workspace self-service app window."

    # =============================================
    # CWA: PLUGINS AND ADD-ONS
    # =============================================
    & $addHeader "CWA: Plugins and Add-ons"

    $chkTeams    = & $addCheckBox "MS Teams VDI Plugin (default on 2508+)" ([bool]$sw.Plugins.MSTeamsPlugin) "Installs MsTeamsPluginCitrix for Teams VDI optimization."
    $chkZoom     = & $addCheckBox "Zoom VDI Plugin (default on 2511+)" ([bool]$sw.Plugins.ZoomPlugin) "Installs 64-bit Zoom VDI plugin."
    $chkWebEx    = & $addCheckBox "WebEx VDI Plugin (ADDONS=WebexVDIPlugin)" ([bool]$sw.Plugins.WebExPlugin) "Installs the WebEx VDI plugin engine."
    $chkUber     = & $addCheckBox "uberAgent Monitoring (/InstallUberAgent)" ([bool]$sw.Plugins.UberAgent) "Installs or upgrades the uberAgent monitoring/diagnostics plugin."
    $chkUberSkip = & $addCheckBox "Skip upgrade if present (/SkipUberAgentUpgrade)" ([bool]$sw.Plugins.UberAgentSkipUpgrade) "Installs uberAgent only if not already present; skips upgrade." 20
    $chkUberSkip.IsEnabled = [bool]$sw.Plugins.UberAgent
    $chkUber.Add_Checked({   $chkUberSkip.IsEnabled = $true }.GetNewClosure())
    $chkUber.Add_Unchecked({ $chkUberSkip.IsEnabled = $false }.GetNewClosure())
    $chkEPA = & $addCheckBox "EPA Client (default on 2508+)" ([bool]$sw.Plugins.EPAClient) "Endpoint Analysis client for Device Posture checks."
    $chkSR  = & $addCheckBox "Session Recording (/InstallSRAgent, 2511+)" ([bool]$sw.Plugins.SessionRecording) "Installs the Session Recording agent for endpoint device session monitoring."

    # =============================================
    # CWA: UPDATE AND TELEMETRY
    # =============================================
    & $addHeader "CWA: Update and Telemetry"

    $cmbAutoUpd = New-Object System.Windows.Controls.ComboBox
    $cmbAutoUpd.FontSize = 13
    $cmbAutoUpd.Width = 120
    $cmbAutoUpd.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    foreach ($val in @("auto", "manual", "disabled")) { [void]$cmbAutoUpd.Items.Add($val) }
    $cmbAutoUpd.SelectedItem = $sw.UpdateAndTelemetry.AutoUpdateCheck
    if ($cmbAutoUpd.SelectedIndex -lt 0) { $cmbAutoUpd.SelectedIndex = 2 }
    & $addLabelRow "Auto-Update:" $cmbAutoUpd "Controls automatic update checking: auto, manual, disabled."

    $chkCEIP  = & $addCheckBox "CEIP / Telemetry (EnableCEIP)" ([bool]$sw.UpdateAndTelemetry.EnableCEIP) "Citrix Customer Experience Improvement Program."
    $chkTrace = & $addCheckBox "Always-On Tracing (EnableTracing)" ([bool]$sw.UpdateAndTelemetry.EnableTracing) "Enables always-on diagnostic tracing."

    # =============================================
    # CWA: STORE POLICY
    # =============================================
    & $addHeader "CWA: Store Policy"

    $cmbAddStore = New-Object System.Windows.Controls.ComboBox
    $cmbAddStore.FontSize = 13
    $cmbAddStore.Width = 60
    $cmbAddStore.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    foreach ($val in @("S", "A", "N")) { [void]$cmbAddStore.Items.Add($val) }
    $cmbAddStore.SelectedItem = $sw.StorePolicy.AllowAddStore
    if ($cmbAddStore.SelectedIndex -lt 0) { $cmbAddStore.SelectedIndex = 0 }
    & $addLabelRow "Allow Add Store:" $cmbAddStore "S = Secure/HTTPS only, A = All protocols, N = None."

    $cmbSavePwd = New-Object System.Windows.Controls.ComboBox
    $cmbSavePwd.FontSize = 13
    $cmbSavePwd.Width = 60
    $cmbSavePwd.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    foreach ($val in @("S", "A", "N")) { [void]$cmbSavePwd.Items.Add($val) }
    $cmbSavePwd.SelectedItem = $sw.StorePolicy.AllowSavePwd
    if ($cmbSavePwd.SelectedIndex -lt 0) { $cmbSavePwd.SelectedIndex = 0 }
    & $addLabelRow "Allow Save Pwd:" $cmbSavePwd "S = Secure only, A = All, N = Never cache credentials."

    # =============================================
    # CWA: COMPONENTS (ADDLOCAL)
    # =============================================
    & $addHeader "CWA: Components (ADDLOCAL)"

    $chkCustomize = & $addCheckBox "Customize (specify ADDLOCAL explicitly)" ([bool]$sw.Components.Customize) "When unchecked, ADDLOCAL is omitted; CWA installs default components."

    $compDefs = @(
        @{ Name = 'ReceiverInside'; Label = 'ReceiverInside (Core SDK)'; Tip = 'Core Workspace SDK services. Required.'; Required = $true },
        @{ Name = 'ICA_Client';     Label = 'ICA_Client (HDX Engine)';   Tip = 'Session launch and ICA protocol handling. Required.'; Required = $true },
        @{ Name = 'AM';             Label = 'AM (Authentication)';       Tip = 'User authentication manager. Required.'; Required = $true },
        @{ Name = 'SelfService';    Label = 'SelfService (Self-Service UI)'; Tip = 'Native application launch and self-service plugin.' },
        @{ Name = 'DesktopViewer';  Label = 'DesktopViewer (Virtual Desktop)'; Tip = 'Virtual desktop UI framework.' },
        @{ Name = 'WebHelper';      Label = 'WebHelper (Browser Helper)'; Tip = 'Browser-to-application connectivity.' },
        @{ Name = 'BCR_Client';     Label = 'BCR_Client (Browser Content Redir.)'; Tip = 'Redirects browser content rendering to the client device.' },
        @{ Name = 'USB';            Label = 'USB (USB Redirection)'; Tip = 'USB device passthrough to virtual sessions.' },
        @{ Name = 'SSON';           Label = 'SSON (SSO Component)'; Tip = 'Single Sign-On GINA/credential provider.' }
    )

    $compCBs = @{}
    foreach ($def in $compDefs) {
        $isChecked = ($sw.Components.($def.Name) -eq $true)
        if ($def.Required) { $isChecked = $true }
        $cb = & $addCheckBox $def.Label $isChecked $def.Tip 20
        $cb.IsEnabled = [bool]$sw.Components.Customize
        $cb.Tag = $def.Name
        $compCBs[$def.Name] = $cb
    }

    $chkCustomize.Add_Checked({
        foreach ($kv in $compCBs.GetEnumerator()) { $kv.Value.IsEnabled = $true }
    }.GetNewClosure())
    $chkCustomize.Add_Unchecked({
        foreach ($kv in $compCBs.GetEnumerator()) { $kv.Value.IsEnabled = $false }
        foreach ($req in @('ReceiverInside', 'ICA_Client', 'AM')) {
            $compCBs[$req].IsChecked = $true
        }
    }.GetNewClosure())

    # Bottom spacer
    $spacer = New-Object System.Windows.Controls.TextBlock
    $spacer.Height = 8
    [void]$panelContent.Children.Add($spacer)

    # =============================================
    # Preview buttons
    # =============================================
    $btnCwaPreview.Add_Click({
        $previewArgs = @('/silent', '/noreboot')
        if ($chkClean.IsChecked)    { $previewArgs += '/CleanInstall' }
        if ($chkSSOn.IsChecked)     { $previewArgs += '/includeSSON'; $previewArgs += 'ENABLE_SSON=Yes' }
        if ($chkAppProt.IsChecked)  { $previewArgs += '/includeappprotection' }
        if ($chkPreLaunch.IsChecked){ $previewArgs += 'ENABLEPRELAUNCH=True' }
        if ($chkSelfSvc.IsChecked)  { $previewArgs += 'SELFSERVICEMODE=True' } else { $previewArgs += 'SELFSERVICEMODE=False' }

        if (-not [string]::IsNullOrWhiteSpace($txtStoreUrl.Text)) {
            $sn = if ([string]::IsNullOrWhiteSpace($txtStoreName.Text)) { 'Store' } else { $txtStoreName.Text.Trim() }
            $su = $txtStoreUrl.Text.Trim().TrimEnd('/')
            if ($su -notlike '*/discovery') { $su = "$su/discovery" }
            $previewArgs += ('STORE0="{0};{1};On;{0}"' -f $sn, $su)
        }

        if (-not $chkTeams.IsChecked)  { $previewArgs += 'InstallMSTeamsPlugin=N' }
        if (-not $chkZoom.IsChecked)   { $previewArgs += 'Installzoomplugin=N' }
        if ($chkWebEx.IsChecked)       { $previewArgs += 'ADDONS=WebexVDIPlugin' }
        if ($chkUber.IsChecked)        { $previewArgs += '/InstallUberAgent'; if ($chkUberSkip.IsChecked) { $previewArgs += '/SkipUberAgentUpgrade' } }
        if (-not $chkEPA.IsChecked)    { $previewArgs += 'InstallEPAClient=N' }
        if ($chkSR.IsChecked)          { $previewArgs += '/InstallSRAgent' }

        $previewArgs += ('AutoUpdateCheck={0}' -f $cmbAutoUpd.SelectedItem)
        if (-not $chkCEIP.IsChecked)   { $previewArgs += 'EnableCEIP=False' }
        if (-not $chkTrace.IsChecked)  { $previewArgs += 'EnableTracing=false' }

        $previewArgs += ('ALLOWADDSTORE={0}' -f $cmbAddStore.SelectedItem)
        $previewArgs += ('ALLOWSAVEPWD={0}' -f $cmbSavePwd.SelectedItem)

        if ($chkCustomize.IsChecked) {
            $cl = @()
            foreach ($kv in $compCBs.GetEnumerator()) { if ($kv.Value.IsChecked) { $cl += $kv.Key } }
            if ($cl.Count -gt 0) { $previewArgs += ('ADDLOCAL={0}' -f ($cl -join ',')) }
        }

        $cmdLine = "CitrixWorkspaceApp.exe " + ($previewArgs -join " ")
        $ownerWin = [System.Windows.Window]::GetWindow($element)
        Show-PreviewDialog -Owner $ownerWin -Title "CWA Preview" -Content $cmdLine -Width 820 -Height 360
    }.GetNewClosure())

    $btnM365Preview.Add_Click({
        try {
            $channelReverseMap = @{ 'Monthly Enterprise Channel' = 'MonthlyEnterprise'; 'Current Channel' = 'Current' }
            $chanRaw = $null
            if ($cmbCH.SelectedItem) { $chanRaw = $channelReverseMap[$cmbCH.SelectedItem.Content] }
            if (-not $chanRaw) { $chanRaw = 'MonthlyEnterprise' }

            $companyName = $txtCN.Text.Trim()

            $excludeList = @()
            foreach ($kv in $excludeCBs.GetEnumerator()) {
                if ($kv.Value.IsChecked -eq $true) { $excludeList += $kv.Key }
            }

            if (-not (Get-Command -Name New-OdtConfigXml -ErrorAction SilentlyContinue)) {
                $ownerWin = [System.Windows.Window]::GetWindow($element)
                Show-PreviewDialog -Owner $ownerWin -Title "M365 Preview (error)" -Content "New-OdtConfigXml not available. Ensure Packagers\AppPackagerCommon.psm1 is importable." -Width 600 -Height 220
                return
            }

            $sb = [System.Text.StringBuilder]::new()
            $products = @(
                @{ Label = 'M365 Apps for Enterprise (x64)'; Edition = '64'; Ids = @('O365ProPlusRetail') },
                @{ Label = 'M365 Apps for Enterprise (x86)'; Edition = '32'; Ids = @('O365ProPlusRetail') },
                @{ Label = 'Project Pro (x64)';              Edition = '64'; Ids = @('ProjectProRetail')   },
                @{ Label = 'Visio Pro (x64)';                Edition = '64'; Ids = @('VisioProRetail')     }
            )
            foreach ($p in $products) {
                [void]$sb.AppendLine(('# ===== {0} =====' -f $p.Label))
                $xml = New-OdtConfigXml -OfficeClientEdition $p.Edition -ProductIds $p.Ids -Channel $chanRaw -CompanyName $companyName -ExcludeApps $excludeList
                [void]$sb.AppendLine($xml)
                [void]$sb.AppendLine('')
            }

            $ownerWin = [System.Windows.Window]::GetWindow($element)
            Show-PreviewDialog -Owner $ownerWin -Title "M365 Preview (install.xml)" -Content $sb.ToString() -Width 820 -Height 640
        } catch {
            $ownerWin = [System.Windows.Window]::GetWindow($element)
            Show-PreviewDialog -Owner $ownerWin -Title "M365 Preview (error)" -Content ("Failed to build preview:`r`n{0}" -f $_.Exception.Message) -Width 600 -Height 260
        }
    }.GetNewClosure())

    # =============================================
    # Commit closure: mutate $sw, $tv, $prefsRef. Master OK handles saves.
    # =============================================
    $prefsRef = $script:Prefs
    $commit = {
        $channelReverseMap = @{ 'Monthly Enterprise Channel' = 'MonthlyEnterprise'; 'Current Channel' = 'Current' }
        $selectedChannel = $channelReverseMap[$cmbCH.SelectedItem.Content]
        if (-not $selectedChannel) { $selectedChannel = 'MonthlyEnterprise' }

        $deployModeReverseMap = @{ 'Managed (Offline)' = 'Managed'; 'Online (CDN)' = 'Online' }
        $selectedDM = $deployModeReverseMap[$cmbDM.SelectedItem.Content]
        if (-not $selectedDM) { $selectedDM = 'Managed' }

        $prefsRef.CompanyName    = $txtCN.Text.Trim()
        $prefsRef.M365Channel    = $selectedChannel
        $prefsRef.M365DeployMode = $selectedDM

        $selectedExcludes = @()
        foreach ($kv in $excludeCBs.GetEnumerator()) {
            if ($kv.Value.IsChecked -eq $true) { $selectedExcludes += $kv.Key }
        }
        $prefsRef.M365ExcludeApps = $selectedExcludes

        if (-not $prefsRef.SSMSInstallOptions) {
            $prefsRef.SSMSInstallOptions = [pscustomobject]@{
                UIMode              = "Quiet"
                DownloadThenInstall = $true
                NoUpdateInstaller   = $false
                IncludeRecommended  = $false
                IncludeOptional     = $false
                RemoveOos           = $true
                ForceClose          = $false
                InstallPath         = ""
            }
        }
        $selectedSsmsUiMode = [string]$cmbSsmsUiMode.SelectedItem
        if ($selectedSsmsUiMode -notin @('Quiet','Passive')) { $selectedSsmsUiMode = 'Quiet' }
        $prefsRef.SSMSInstallOptions.UIMode              = $selectedSsmsUiMode
        $prefsRef.SSMSInstallOptions.DownloadThenInstall = ($chkSsmsDownloadThenInstall.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.NoUpdateInstaller   = ($chkSsmsNoUpdateInstaller.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.IncludeRecommended  = ($chkSsmsRecommended.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.IncludeOptional     = ($chkSsmsOptional.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.RemoveOos           = ($chkSsmsRemoveOos.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.ForceClose          = ($chkSsmsForceClose.IsChecked -eq $true)
        $prefsRef.SSMSInstallOptions.InstallPath         = [string]$txtSsmsInstallPath.Text.Trim()

        $sw.Store.Name = $txtStoreName.Text.Trim()
        $sw.Store.Url  = $txtStoreUrl.Text.Trim()

        $sw.Installation.CleanInstall     = ($chkClean.IsChecked -eq $true)
        $sw.Installation.IncludeSSON      = ($chkSSOn.IsChecked -eq $true)
        $sw.Installation.EnableSSON       = ($chkSSOn.IsChecked -eq $true)
        $sw.Installation.AppProtection    = ($chkAppProt.IsChecked -eq $true)
        $sw.Installation.SessionPreLaunch = ($chkPreLaunch.IsChecked -eq $true)
        $sw.Installation.SelfServiceMode  = ($chkSelfSvc.IsChecked -eq $true)

        $sw.Plugins.MSTeamsPlugin        = ($chkTeams.IsChecked -eq $true)
        $sw.Plugins.ZoomPlugin           = ($chkZoom.IsChecked -eq $true)
        $sw.Plugins.WebExPlugin          = ($chkWebEx.IsChecked -eq $true)
        $sw.Plugins.UberAgent            = ($chkUber.IsChecked -eq $true)
        $sw.Plugins.UberAgentSkipUpgrade = ($chkUberSkip.IsChecked -eq $true)
        $sw.Plugins.EPAClient            = ($chkEPA.IsChecked -eq $true)
        $sw.Plugins.SessionRecording     = ($chkSR.IsChecked -eq $true)

        $sw.UpdateAndTelemetry.AutoUpdateCheck = [string]$cmbAutoUpd.SelectedItem
        $sw.UpdateAndTelemetry.EnableCEIP      = ($chkCEIP.IsChecked -eq $true)
        $sw.UpdateAndTelemetry.EnableTracing   = ($chkTrace.IsChecked -eq $true)

        $sw.StorePolicy.AllowAddStore = [string]$cmbAddStore.SelectedItem
        $sw.StorePolicy.AllowSavePwd  = [string]$cmbSavePwd.SelectedItem

        $sw.Components.Customize = ($chkCustomize.IsChecked -eq $true)
        foreach ($kv in $compCBs.GetEnumerator()) {
            $sw.Components.($kv.Key) = ($kv.Value.IsChecked -eq $true)
        }

        $tv.ApiToken              = [string]$txtTvApiToken.Text
        $tv.CustomConfigId        = [string]$txtTvConfigId.Text
        $tv.AssignmentOptions     = [string]$txtTvAssignOpts.Text
        $tv.RemoveDesktopShortcut = ($chkTvRemoveShortcut.IsChecked -eq $true)
    }.GetNewClosure()

    return @{
        Name        = 'Packager Preferences'
        Element     = $element
        Commit      = $commit
        CwaSwitches = $sw
        TvConfig    = $tv
    }
}

function Show-CommandOverrideDialog {
    # Modal editor for one app's install/uninstall command overrides.
    # Returns $null on cancel, otherwise @{ Install; Uninstall } with
    # trimmed values (both empty = revert to the shipped commands).
    param(
        [Parameter(Mandatory)][string]$AppLabel,
        [AllowEmptyString()][string]$Install = '',
        [AllowEmptyString()][string]$Uninstall = '',
        [Parameter(Mandatory)]$Owner
    )
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="560" SizeToContent="Height" MinWidth="440"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ResizeMode="NoResize" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="110"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="110"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtIntro" Grid.Row="0" TextWrapping="Wrap" FontSize="12" Margin="0,4,0,12"/>
        <TextBlock Grid.Row="1" Text="Install command:" FontSize="12" Margin="0,0,0,4"/>
        <TextBox   x:Name="txtInstall" Grid.Row="2" FontSize="12" FontFamily="Consolas" Margin="0,0,0,10"
                   Controls:TextBoxHelper.Watermark="install.bat (packager default)"/>
        <TextBlock Grid.Row="3" Text="Uninstall command:" FontSize="12" Margin="0,0,0,4"/>
        <TextBox   x:Name="txtUninstall" Grid.Row="4" FontSize="12" FontFamily="Consolas" Margin="0,0,0,16"
                   Controls:TextBoxHelper.Watermark="uninstall.bat (packager default)"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnRevert" Content="Revert to default" Style="{StaticResource DialogButton}"/>
            <Button x:Name="btnSave"   Content="Save"              Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel"            Style="{StaticResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $Owner
    $dlg.Title = "Command overrides - $AppLabel"
    Install-TitleBarDragFallback -Window $dlg
    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.NonActiveWindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.NonActiveGlowBrush
    } catch { $null = $_ }
    $dlg.FindName('txtIntro').Text = "Replaces the deployment type command lines the next time $AppLabel is packaged. An empty field keeps the packager's shipped command. Not applied to variant-split apps, whose variants carry their own commands."
    $txtInstall = $dlg.FindName('txtInstall')
    $txtUninstall = $dlg.FindName('txtUninstall')
    $txtInstall.Text = $Install
    $txtUninstall.Text = $Uninstall
    $result = $null
    $dlg.FindName('btnRevert').Add_Click({ $txtInstall.Text = ''; $txtUninstall.Text = '' }.GetNewClosure())
    $dlg.FindName('btnSave').Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $dlg.FindName('btnCancel').Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    if ([bool]$dlg.ShowDialog()) {
        $result = @{ Install = $txtInstall.Text.Trim(); Uninstall = $txtUninstall.Text.Trim() }
    }
    return $result
}

function New-DeploymentConditionsPanel {
    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" FontSize="12"
               Foreground="{DynamicResource MahApps.Brushes.Gray3}" Margin="0,0,0,10"
               Text="Per-app requirement rules attached to the deployment type at Package time. The client evaluates them at install time, so no collections are involved. Site conditions are created in MECM on first use; change a name below to attach to a condition your site already has."/>
    <Border DockPanel.Dock="Top" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="1" Padding="10,8,10,8" Margin="0,0,0,10">
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Grid.Column="0" Text="CPU architecture condition:" VerticalAlignment="Center" FontSize="12" Margin="0,0,10,6"/>
            <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtArchGc" FontSize="12" Margin="0,0,0,6"
                       ToolTip="Site global condition name for the WQL query on Win32_Processor.Architecture (9 = x64, 12 = ARM64)"/>
            <TextBlock Grid.Row="1" Grid.Column="0" Text="OS language condition:" VerticalAlignment="Center" FontSize="12" Margin="0,0,10,6"/>
            <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtLangGc" FontSize="12" Margin="0,0,0,6"
                       ToolTip="Name of the built-in Operating System Language site condition. Must already exist on the site."/>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="VPN condition:" VerticalAlignment="Center" FontSize="12" Margin="0,0,10,6"/>
            <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtVpnGc" FontSize="12" Margin="0,0,0,6"
                       ToolTip="Site global condition name for the script that reports whether a VPN adapter is active"/>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="VPN adapter patterns:" VerticalAlignment="Center" FontSize="12" Margin="0,0,10,0"/>
            <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtVpnPatterns" FontSize="12"
                       ToolTip="Comma-separated adapter description substrings that identify a VPN client. The script also matches interface aliases containing 'vpn'. Changes apply when the condition is next created; an existing site condition keeps its script."/>
        </Grid>
    </Border>
    <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" FontSize="11"
               Foreground="{DynamicResource MahApps.Brushes.Gray3}" Margin="0,0,0,8"
               Text="OS languages take comma-separated culture codes (e.g. de-DE, en-US); empty means no language requirement. Network 'VPN only' installs only when a VPN adapter is active, 'On-site only' installs only when it is not."/>
    <DataGrid x:Name="dgCondApps" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
              GridLinesVisibility="Horizontal" HeadersVisibility="Column" RowHeaderWidth="0" BorderThickness="0"
              IsTextSearchEnabled="True" TextSearch.TextPath="Application">
        <DataGrid.Columns>
            <DataGridTextColumn Header="Application" Width="*" Binding="{Binding Application}" IsReadOnly="True"/>
            <DataGridTextColumn Header="Vendor" Width="140" Binding="{Binding Vendor}" IsReadOnly="True"/>
            <DataGridComboBoxColumn Header="Architecture" Width="110" SelectedItemBinding="{Binding ArchitectureDisplay, UpdateSourceTrigger=PropertyChanged}"/>
            <DataGridTextColumn Header="OS languages" Width="150" Binding="{Binding LanguagesDisplay, UpdateSourceTrigger=LostFocus, Mode=TwoWay}"/>
            <DataGridComboBoxColumn Header="Network" Width="110" SelectedItemBinding="{Binding NetworkDisplay, UpdateSourceTrigger=PropertyChanged}"/>
            <DataGridTemplateColumn Header="Commands" Width="100">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <Button Content="{Binding CommandLabel}" FontSize="11" Padding="6,1,6,1" Margin="2"
                                Tag="{Binding}" x:Name="btnRowCommands"
                                ToolTip="Override the install/uninstall command lines for this app. Empty fields use the packager's shipped commands."/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTemplateColumn Header="Variant split" Width="120">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <ComboBox ItemsSource="{Binding VariantOptions}" SelectedItem="{Binding SplitDisplay, UpdateSourceTrigger=PropertyChanged}"
                                  IsEnabled="{Binding VariantCapable}" FontSize="12" BorderThickness="0" Background="Transparent"
                                  ToolTip="Stage one application with multiple deployment types. Only offered where the packager declares SupportsVariants."/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
        </DataGrid.Columns>
    </DataGrid>
</DockPanel>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $txtArchGc      = $element.FindName('txtArchGc')
    $txtLangGc      = $element.FindName('txtLangGc')
    $txtVpnGc       = $element.FindName('txtVpnGc')
    $txtVpnPatterns = $element.FindName('txtVpnPatterns')
    $dgCondApps     = $element.FindName('dgCondApps')

    $dgCondApps.Columns[2].ItemsSource = [string[]]@('Any', 'x64 only', 'ARM64 only')
    $dgCondApps.Columns[4].ItemsSource = [string[]]@('Any', 'VPN only', 'On-site only')

    $condDoc = Get-ConditionTemplates
    $archTemplate = @($condDoc.Conditions | Where-Object { [string]$_.Id -eq 'cpu-arch' })
    $langTemplate = @($condDoc.Conditions | Where-Object { [string]$_.Id -eq 'os-language' })
    $vpnTemplate  = @($condDoc.Conditions | Where-Object { [string]$_.Id -eq 'vpn-connected' })
    $archTemplate = if ($archTemplate.Count -gt 0) { $archTemplate[0] } else { $null }
    $langTemplate = if ($langTemplate.Count -gt 0) { $langTemplate[0] } else { $null }
    $vpnTemplate  = if ($vpnTemplate.Count  -gt 0) { $vpnTemplate[0]  } else { $null }

    if ($archTemplate) { $txtArchGc.Text = [string]$archTemplate.GlobalConditionName } else { $txtArchGc.IsEnabled = $false }
    if ($langTemplate) { $txtLangGc.Text = [string]$langTemplate.GlobalConditionName } else { $txtLangGc.IsEnabled = $false }
    if ($vpnTemplate) {
        $txtVpnGc.Text       = [string]$vpnTemplate.GlobalConditionName
        $txtVpnPatterns.Text = (@($vpnTemplate.AdapterPatterns) -join ', ')
    }
    else {
        $txtVpnGc.IsEnabled       = $false
        $txtVpnPatterns.IsEnabled = $false
    }

    $archToDisplay = @{ 'Any' = 'Any'; 'x64' = 'x64 only'; 'ARM64' = 'ARM64 only' }
    $networkToDisplay = @{ 'Any' = 'Any'; 'VpnOnly' = 'VPN only'; 'OnSiteOnly' = 'On-site only' }

    $currentApps = $script:Prefs.DeploymentConditions.Apps
    $currentCommands = $script:Prefs.CommandOverrides.Apps
    $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $packagers = Get-Packagers -Root $PackagersRoot | Sort-Object Vendor, Application
    foreach ($p in $packagers) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($p.Script)
        $arch = 'Any'
        $network = 'Any'
        $langsText = ''
        $entryProp = $null
        if ($currentApps) { $entryProp = $currentApps.PSObject.Properties[$base] }
        $split = 'None'
        if ($entryProp) {
            $entry = $entryProp.Value
            if ([string]$entry.Architecture -in @('x64', 'ARM64')) { $arch = [string]$entry.Architecture }
            if ([string]$entry.Network -in @('VpnOnly', 'OnSiteOnly')) { $network = [string]$entry.Network }
            if ($entry.Languages) { $langsText = (@($entry.Languages) -join ', ') }
            if ($entry.PSObject.Properties['Split'] -and [string]$entry.Split -in @($p.SupportsVariants)) { $split = [string]$entry.Split }
        }
        $cmdInstall = ''
        $cmdUninstall = ''
        if ($currentCommands) {
            $cmdProp = $currentCommands.PSObject.Properties[$base]
            if ($cmdProp) {
                $cmdInstall = ([string]$cmdProp.Value.Install).Trim()
                $cmdUninstall = ([string]$cmdProp.Value.Uninstall).Trim()
            }
        }
        $variantOptions = @('None') + @($p.SupportsVariants)
        $rows.Add([pscustomobject]@{
            Packager            = $base
            Application         = $p.Application
            Vendor              = $p.Vendor
            ArchitectureDisplay = $archToDisplay[$arch]
            LanguagesDisplay    = $langsText
            NetworkDisplay      = $networkToDisplay[$network]
            VariantOptions      = [string[]]$variantOptions
            VariantCapable      = (@($p.SupportsVariants).Count -gt 0)
            SplitDisplay        = $split
            CmdInstall          = $cmdInstall
            CmdUninstall        = $cmdUninstall
            CommandLabel        = $(if ($cmdInstall -or $cmdUninstall) { 'Modified' } else { 'Default' })
        })
    }
    $dgCondApps.ItemsSource = $rows

    # Template-column buttons share one routed handler; the row rides in
    # on the button's Tag. Combo boxes in the split column do not raise
    # Button.Click, so no filtering is needed.
    $dgCondApps.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($s, $e)
        $btn = $e.OriginalSource
        if ($btn -isnot [System.Windows.Controls.Button]) { return }
        $row = $btn.Tag
        if (-not $row -or -not $row.PSObject.Properties['CmdInstall']) { return }
        $edited = Show-CommandOverrideDialog -AppLabel ([string]$row.Application) -Install ([string]$row.CmdInstall) -Uninstall ([string]$row.CmdUninstall) -Owner $window
        if ($null -ne $edited) {
            $row.CmdInstall = [string]$edited.Install
            $row.CmdUninstall = [string]$edited.Uninstall
            $row.CommandLabel = $(if ($row.CmdInstall -or $row.CmdUninstall) { 'Modified' } else { 'Default' })
            $dgCondApps.Items.Refresh()
        }
    })

    $condState = @{
        Doc   = $condDoc
        Dirty = $false
        InitialArchGc   = $txtArchGc.Text
        InitialLangGc   = $txtLangGc.Text
        InitialVpnGc    = $txtVpnGc.Text
        InitialPatterns = $txtVpnPatterns.Text
    }

    $prefsRef = $script:Prefs
    $commit = {
        [void]$dgCondApps.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$dgCondApps.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,  $true)

        $displayToArch = @{ 'Any' = 'Any'; 'x64 only' = 'x64'; 'ARM64 only' = 'ARM64' }
        $displayToNetwork = @{ 'Any' = 'Any'; 'VPN only' = 'VpnOnly'; 'On-site only' = 'OnSiteOnly' }

        $condProps = [ordered]@{}
        foreach ($row in $rows) {
            $arch = 'Any'
            if ($displayToArch.ContainsKey([string]$row.ArchitectureDisplay)) { $arch = $displayToArch[[string]$row.ArchitectureDisplay] }
            $network = 'Any'
            if ($displayToNetwork.ContainsKey([string]$row.NetworkDisplay)) { $network = $displayToNetwork[[string]$row.NetworkDisplay] }
            $langs = @([string]$row.LanguagesDisplay -split '[,;]' |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ -match '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,2}$' })
            $split = 'None'
            if ([string]$row.SplitDisplay -in @('Architecture', 'Language', 'Network') -and $row.VariantCapable) { $split = [string]$row.SplitDisplay }
            if ($arch -eq 'Any' -and $network -eq 'Any' -and $langs.Count -eq 0 -and $split -eq 'None') { continue }
            $condProps[$row.Packager] = [pscustomobject]@{
                Architecture = $arch
                Languages    = $langs
                Network      = $network
                Split        = $split
            }
        }
        $prefsRef.DeploymentConditions.Apps = [pscustomobject]$condProps

        $cmdProps = [ordered]@{}
        foreach ($row in $rows) {
            $inst = ([string]$row.CmdInstall).Trim()
            $uninst = ([string]$row.CmdUninstall).Trim()
            if (-not $inst -and -not $uninst) { continue }
            $cmdProps[$row.Packager] = [pscustomobject]@{
                Install   = $inst
                Uninstall = $uninst
            }
        }
        $prefsRef.CommandOverrides.Apps = [pscustomobject]$cmdProps

        if ($archTemplate -and -not [string]::IsNullOrWhiteSpace($txtArchGc.Text)) {
            $archTemplate.GlobalConditionName = $txtArchGc.Text.Trim()
        }
        if ($langTemplate -and -not [string]::IsNullOrWhiteSpace($txtLangGc.Text)) {
            $langTemplate.GlobalConditionName = $txtLangGc.Text.Trim()
        }
        if ($vpnTemplate) {
            if (-not [string]::IsNullOrWhiteSpace($txtVpnGc.Text)) {
                $vpnTemplate.GlobalConditionName = $txtVpnGc.Text.Trim()
            }
            $patterns = @([string]$txtVpnPatterns.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($patterns.Count -gt 0) { $vpnTemplate.AdapterPatterns = $patterns }
        }
        if ($txtArchGc.Text -ne $condState.InitialArchGc -or
            $txtLangGc.Text -ne $condState.InitialLangGc -or
            $txtVpnGc.Text -ne $condState.InitialVpnGc -or
            $txtVpnPatterns.Text -ne $condState.InitialPatterns) {
            $condState.Dirty = $true
        }
    }.GetNewClosure()

    return @{
        Name           = 'Deployment Conditions'
        Element        = $element
        Commit         = $commit
        ConditionState = $condState
    }
}

function Show-OptionsDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$InitialSection = 'MECM Preferences'
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="940" Height="640"
    MinWidth="820" MinHeight="520"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ListBox Grid.Column="0" Grid.Row="0" x:Name="lstNav" BorderThickness="0" Padding="0,8,0,0">
            <ListBox.ItemContainerStyle>
                <Style TargetType="ListBoxItem">
                    <Setter Property="Padding" Value="16,10,16,10"/>
                    <Setter Property="FontSize" Value="13"/>
                </Style>
            </ListBox.ItemContainerStyle>
        </ListBox>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <ContentControl Grid.Column="2" Grid.Row="0" x:Name="contentArea" Margin="20,18,20,18"/>

        <Border Grid.Column="0" Grid.ColumnSpan="3" Grid.Row="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="20,12,20,12">
                <Button x:Name="btnOK"     Content="OK"     MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True" Style="{DynamicResource MahApps.Styles.Button.Square.Accent}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnCancel" Content="Cancel" MinWidth="90" Height="32" IsCancel="True" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $lstNav      = $dlg.FindName('lstNav')
    $contentArea = $dlg.FindName('contentArea')
    $btnOK       = $dlg.FindName('btnOK')
    $btnCancel   = $dlg.FindName('btnCancel')

    # All panels live in the unified Options window.
    $panels = @(
        (New-MecmPreferencesPanel),
        (New-PackagerPreferencesPanel),
        (New-AppFlowPanel),
        (New-ProductFilterPanel),
        (New-DeploymentConditionsPanel)
    )

    foreach ($p in $panels) { [void]$lstNav.Items.Add($p.Name) }

    $lstNav.Add_SelectionChanged({
        $idx = $lstNav.SelectedIndex
        if ($idx -ge 0 -and $idx -lt $panels.Count) {
            $contentArea.Content = $panels[$idx].Element
        }
    })

    $initialIdx = 0
    for ($i = 0; $i -lt $panels.Count; $i++) {
        if ($panels[$i].Name -eq $InitialSection) { $initialIdx = $i; break }
    }
    $lstNav.SelectedIndex = $initialIdx

    $script:OptionsDlgResult = $false
    $btnOK.Add_Click({
        try {
            foreach ($p in $panels) { if ($p.Commit) { & $p.Commit } }
            Save-Preferences -Prefs $script:Prefs
            # Panels that mutate sibling JSON configs expose the refs on
            # the panel hash; master persists them here so panel commits
            # stay free of function calls (GetNewClosure-safe).
            foreach ($p in $panels) {
                if ($p.CwaSwitches) { Save-CwaSwitches -Switches $p.CwaSwitches }
                if ($p.TvConfig)    { Save-TvHostConfig -Config $p.TvConfig }
                if ($p.ConditionState -and $p.ConditionState.Dirty) {
                    [void](Save-ConditionTemplates -Templates $p.ConditionState.Doc)
                }
            }
            Invoke-RefreshGrid
            $script:OptionsDlgResult = $true
            $dlg.Close()
        } catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save Failed' -Message $_.Exception.Message -Buttons OK -Icon Error)
        }
    })

    $btnCancel.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog()

    if ($script:OptionsDlgResult) {
        Add-LogLine -Message "Options saved."
    }
}

# =============================================================================
# Grid refresh helper
# =============================================================================
function Invoke-RefreshGrid {
    $script:PackagerData.Clear()

    $items = Get-Packagers -Root $PackagersRoot
    $hidden = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($script:Prefs.HiddenApplications),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    # Pre-load persistent history so first render shows Latest + LastChecked
    # from prior sessions (Full Run + manual Check Latest both write here).
    $history = @{}
    try { $history = Read-PackagerHistory } catch { }

    $hiddenCount = 0
    foreach ($m in $items) {
        if ($hidden.Contains($m.Script)) { $hiddenCount++; continue }

        $baseName     = [System.IO.Path]::GetFileNameWithoutExtension($m.Script)
        $latestStored = ""
        $lastChecked  = ""
        if ($history.ContainsKey($baseName)) {
            $entry = $history[$baseName]
            if ($entry -is [hashtable]) {
                if ($entry['LastKnownVersion']) { $latestStored = [string]$entry['LastKnownVersion'] }
                if ($entry['LastChecked'])      { $lastChecked  = [string]$entry['LastChecked'] }
            } else {
                if ($entry.LastKnownVersion) { $latestStored = [string]$entry.LastKnownVersion }
                if ($entry.LastChecked)      { $lastChecked  = [string]$entry.LastChecked }
            }
        }

        $script:PackagerData.Add([pscustomobject]@{
            Selected       = $false
            Vendor         = $m.Vendor
            Application    = $m.Application
            CurrentVersion = ""
            LatestVersion  = $latestStored
            Status         = $m.Status
            CMName         = $m.CMName
            Script         = $m.Script
            FullPath       = $m.FullPath
            VendorURL      = $m.VendorUrl
            Description    = $m.Description
            LastChecked    = $lastChecked
        })
    }

    if ($hiddenCount -gt 0) {
        $txtStatus.Text = ("{0} packager(s) loaded, {1} hidden. Ready." -f $script:PackagerData.Count, $hiddenCount)
    }
    else {
        $txtStatus.Text = ("Loaded {0} packager(s). Ready." -f $script:PackagerData.Count)
    }

    # A rebuild replaces every row object; a filtered grid would otherwise
    # keep showing the orphaned old rows.
    Update-GridFilter
}

# =============================================================================
# Async pipeline: background runspace + DispatcherTimer overlay.
# Brand-standard "beautiful spinner" pattern per
# reference_wpf_async_progress_overlay.md. Moves the per-app loops for
# Check Latest / Stage / Package / Full Run off the UI thread so the
# ProgressRing animates continuously, log drawer drains a queue instead
# of being mutated from bg, and row.Status flips render via a Refresh
# tick. Single-row button clicks remain synchronous.
# =============================================================================
$script:BgRunspace = $null
$script:BgPS       = $null
$script:BgHandle   = $null
$script:BgState    = $null
$script:BgTimer    = $null
$script:BgGraveyard = @()

function Initialize-BackgroundWorker {
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }

    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()

    # Pre-import AppPackagerCommon into the bg runspace so Update-PackagerHistory /
    # Read-PackagerHistory / New-MECMApplicationFromManifest / Write-StageManifest
    # resolve inside the bg scriptblock.
    $modulePath = Join-Path $PSScriptRoot 'Packagers\AppPackagerCommon.psm1'
    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
    }).AddArgument($modulePath)
    [void]$initPS.Invoke()
    $initPS.Dispose()

    # Reflect the inline helpers from this script into the bg runspace so
    # the per-app loop can call them directly. Single source of truth: the
    # helpers live in this file; the bg runspace gets definition snapshots.
    #
    # AST-enumerate every top-level function defined in this script rather
    # than maintain a hand-curated whitelist. The whitelist approach
    # previously broke silently whenever a new helper (or a new transitive
    # callee) was added to the bg-called path: callers like
    # Invoke-PackagerStage would throw "The term 'X' is not recognized"
    # because X wasn't in the list. Auto-enumeration is self-healing.
    # UI-only functions (those referencing $window / $dataGrid / etc.)
    # come along for the ride; they are harmless as long as they are
    # never *called* from the bg scriptblock.
    $selfTokens = $null
    $selfErrors = $null
    $selfAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $PSCommandPath, [ref]$selfTokens, [ref]$selfErrors
    )
    $fnDefs = @($selfAst.FindAll({
        param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true))
    $sb = [System.Text.StringBuilder]::new()
    foreach ($fn in $fnDefs) {
        [void]$sb.AppendLine($fn.Extent.Text)
    }
    $injectPS = [powershell]::Create()
    $injectPS.Runspace = $script:BgRunspace
    [void]$injectPS.AddScript($sb.ToString())
    [void]$injectPS.Invoke()
    $injectPS.Dispose()
}

function Invoke-MultiAppPipeline {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('CheckLatest','Stage','Package','FullRun')]
        [string]$Operation,
        [Parameter(Mandatory)]
        [array]$Rows,
        [Parameter(Mandatory)]
        [hashtable]$Context
    )

    Initialize-BackgroundWorker

    # Cancel any in-flight pipeline. BeginStop is best-effort and
    # non-blocking; the stopping pipeline parks in the graveyard until it
    # actually stops instead of freezing the UI thread on a stuck CM call.
    $script:BgGraveyard = @(Stop-SuiteBgWork -PowerShell $script:BgPS -Timer $script:BgTimer -Graveyard $script:BgGraveyard)
    $script:BgTimer  = $null
    $script:BgPS     = $null
    $script:BgHandle = $null
    $script:BgState  = $null

    # Synchronized state bridges bg -> UI. LogQueue is a ConcurrentQueue
    # so bg can enqueue without locking; the DispatcherTimer drains it
    # into Add-LogLine on the UI thread each tick.
    $script:BgState = [hashtable]::Synchronized(@{
        Step            = 'Starting...'
        Done            = $false
        ErrorMsg        = $null
        Paused          = $false
        CancelRequested = $false
        Canceled        = $false
        LogQueue        = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
        Counts          = $null
    })

    Set-ActionButtonsEnabled -Enabled $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait

    $titleMap = @{
        CheckLatest = 'Checking latest versions'
        Stage       = 'Staging packages'
        Package     = 'Packaging applications'
        FullRun     = 'One Click flow'
    }
    $txtProgressTitle.Text = $titleMap[$Operation]
    $txtProgressStep.Text  = 'Starting...'
    $btnPausePipeline.Content = 'Pause'
    $btnPausePipeline.IsEnabled = $true
    $btnCancelPipeline.IsEnabled = $true
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible

    $rowsArray = @($Rows)

    $script:BgPS = [powershell]::Create()
    $script:BgPS.Runspace = $script:BgRunspace
    [void]$script:BgPS.AddScript({
        param($Op, $RowsIn, $Ctx, $State)

        $counts = [ordered]@{
            Checked = 0; Updated = 0; Reported = 0; Staged = 0; Packaged = 0; StageAndPackage = 0
            NoChange = 0; Skipped = 0; Failed = 0; CheckFailed = 0
        }

        try {
            $rows = @($RowsIn)
            $n = $rows.Count
            $i = 0
            foreach ($row in $rows) {
                while ([bool]$State.Paused -and -not [bool]$State.CancelRequested) {
                    $State.Step = 'Paused before next app'
                    Start-Sleep -Milliseconds 250
                }
                if ([bool]$State.CancelRequested) {
                    $State.Canceled = $true
                    [void]$State.LogQueue.Enqueue('Canceled. Stopped before starting the next app.')
                    break
                }

                $i++
                $app      = [string]$row.Application
                $scrName  = [string]$row.Script
                $path     = [string]$row.FullPath
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($scrName)

                switch ($Op) {
                    'CheckLatest' {
                        $State.Step = ('Check {0}/{1}: {2}' -f $i, $n, $app)
                        $row.Status = 'Checking latest...'
                        [void]$State.LogQueue.Enqueue(('Latest: {0} ({1})' -f $app, $scrName))
                        try {
                            $priorKnown = [string]$row.LatestVersion
                            $latest = Invoke-PackagerGetLatestVersion `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -FileServerPath $Ctx.FileShareRoot `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode
                            $row.LatestVersion = $latest

                            $current = [string]$row.CurrentVersion
                            if (-not [string]::IsNullOrWhiteSpace($current)) {
                                $cmp = Compare-SemVer -A $current -B $latest
                                if ($cmp -lt 0)     { $row.Status = 'Update available' }
                                elseif ($cmp -eq 0) { $row.Status = 'Up to date' }
                                else                { $row.Status = 'Current newer' }
                            } else {
                                $row.Status = 'Latest retrieved'
                            }

                            $suffix = ''
                            if ($scrName -match 'm365') {
                                $chMap = @{ 'MonthlyEnterprise' = 'MEC'; 'Current' = 'CC' }
                                $suffix = ' [' + $chMap[$Ctx.M365Channel] + ']'
                            }
                            [void]$State.LogQueue.Enqueue(('Latest version: {0}{1}' -f $latest, $suffix))

                            $histResult = if ($priorKnown -and $priorKnown -eq $latest) { 'NoChange' } else { 'Updated' }
                            try {
                                Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result $histResult
                                $row.LastChecked = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                            } catch {
                                [void]$State.LogQueue.Enqueue(('History write failed for {0}: {1}' -f $baseName, $_.Exception.Message))
                            }
                            $counts['Checked']++
                            if ($histResult -eq 'NoChange') { $counts['NoChange']++ }
                            else { $counts['Updated']++ }
                        } catch {
                            $row.Status = 'Error'
                            [void]$State.LogQueue.Enqueue(('Error: ' + $_.Exception.Message))
                            try { Update-PackagerHistory -PackagerName $baseName -Event Checked -Result Failed } catch { }
                            $counts['CheckFailed']++
                        }
                    }

                    'Stage' {
                        $State.Step = ('Stage {0}/{1}: {2}' -f $i, $n, $app)
                        $row.Status = 'Staging...'
                        [void]$State.LogQueue.Enqueue(('Stage: {0} ({1})' -f $app, $scrName))
                        try {
                            $res = Invoke-PackagerStage `
                                -PackagerPath $path `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -SevenZipPath $Ctx.SevenZipPath

                            if ($res.ExitCode -eq 0) {
                                $row.Status = 'Staged'
                                [void]$State.LogQueue.Enqueue(('Staged. Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                $ver = [string]$row.LatestVersion
                                try {
                                    if ($ver) { Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $ver -Result Updated }
                                    else      { Update-PackagerHistory -PackagerName $baseName -Event Staged -Result Updated }
                                } catch { }
                                $counts['Staged']++
                            } else {
                                $row.Status = 'Stage error'
                                $stderrLines = @($res.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                if ($stderrLines.Count -gt 0) {
                                    $linesToShow = [Math]::Min($stderrLines.Count, 10)
                                    for ($k = 0; $k -lt $linesToShow; $k++) {
                                        [void]$State.LogQueue.Enqueue(('  stderr: ' + $stderrLines[$k]))
                                    }
                                } else {
                                    [void]$State.LogQueue.Enqueue(('Stage exit code {0}, no stderr.' -f $res.ExitCode))
                                }
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                $counts['Failed']++
                            }
                        } catch {
                            $row.Status = 'Stage error'
                            [void]$State.LogQueue.Enqueue(('Stage exception: ' + $_.Exception.Message))
                            $counts['Failed']++
                        }
                    }

                    'Package' {
                        $State.Step = ('Package {0}/{1}: {2}' -f $i, $n, $app)
                        $row.Status = 'Packaging...'
                        [void]$State.LogQueue.Enqueue(('Package: {0} ({1})' -f $app, $scrName))
                        try {
                            $reqJson = ''
                            if ($Ctx.RequirementsByApp) { $reqJson = [string]$Ctx.RequirementsByApp[$baseName] }
                            $varJson = ''
                            if ($Ctx.VariantsByApp) { $varJson = [string]$Ctx.VariantsByApp[$baseName] }
                            $cmdJson = ''
                            if ($Ctx.CommandsByApp) { $cmdJson = [string]$Ctx.CommandsByApp[$baseName] }
                            $res = Invoke-PackagerPackage `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -ProviderMachineName $Ctx.ProviderMachineName `
                                -Comment $Ctx.Comment `
                                -FileServerPath $Ctx.FileShareRoot `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -EstimatedRuntimeMins $Ctx.EstimatedRuntimeMins `
                                -MaximumRuntimeMins $Ctx.MaximumRuntimeMins `
                                -SevenZipPath $Ctx.SevenZipPath `
                                -CreateIntuneWin:([bool]$Ctx.IntuneWinCreate) `
                                -IntuneWinToolPath ([string]$Ctx.IntuneWinToolPath) `
                                -ContentLayout ([string]$Ctx.ContentLayout) `
                                -RequirementsJson $reqJson `
                                -VariantsJson $varJson `
                                -CommandsJson $cmdJson

                            if ($res.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                if ($res.PSObject.Properties['IntuneWin'] -and $res.IntuneWin) {
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $res.IntuneWin.Message))
                                    if ($res.IntuneWin.NetworkPath) {
                                        [void]$State.LogQueue.Enqueue(('Intunewin on network: ' + $res.IntuneWin.NetworkPath))
                                    }
                                }
                                $ver = [string]$row.LatestVersion
                                try {
                                    if ($ver) { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $ver -Result Updated }
                                    else      { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Result Updated }
                                } catch { }
                                $counts['Packaged']++
                            } else {
                                $row.Status = 'Package error'
                                $stderrLines = @($res.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                if ($stderrLines.Count -gt 0) {
                                    $linesToShow = [Math]::Min($stderrLines.Count, 10)
                                    for ($k = 0; $k -lt $linesToShow; $k++) {
                                        [void]$State.LogQueue.Enqueue(('  stderr: ' + $stderrLines[$k]))
                                    }
                                } else {
                                    [void]$State.LogQueue.Enqueue(('Package exit code {0}, no stderr.' -f $res.ExitCode))
                                }
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                $counts['Failed']++
                            }
                        } catch {
                            $row.Status = 'Package error'
                            [void]$State.LogQueue.Enqueue(('Package exception: ' + $_.Exception.Message))
                            $counts['Failed']++
                        }
                    }

                    'FullRun' {
                        # Cadence gate (Report only), MECM pre-flight, then Stage + optional Package.
                        # Mirrors the UI-thread handler behavior 1:1 so a Full Run here lands the
                        # same history entries and row.Status flips as the old path did.
                        $State.Step = ('One Click {0}/{1}: {2}' -f $i, $n, $app)

                        $lastChecked = $null; $lastKnown = $null; $lastStaged = $null; $lastPackaged = $null
                        try {
                            $hist = Read-PackagerHistory
                            if ($hist.ContainsKey($baseName)) {
                                $h = $hist[$baseName]
                                if ($h -is [hashtable]) {
                                    $lastChecked  = $h['LastChecked']
                                    $lastKnown    = $h['LastKnownVersion']
                                    $lastStaged   = $h['LastStaged']
                                    $lastPackaged = $h['LastPackaged']
                                } else {
                                    $lastChecked  = $h.LastChecked
                                    $lastKnown    = $h.LastKnownVersion
                                    $lastStaged   = $h.LastStaged
                                    $lastPackaged = $h.LastPackaged
                                }
                            }
                        } catch { }

                        if ($Ctx.Action -eq 'Report' -and -not $Ctx.ForceFlag -and $lastChecked) {
                            $cadenceDays = 7
                            $fromOverride = $false
                            if ($Ctx.Overrides) {
                                $op = $Ctx.Overrides.PSObject.Properties[$baseName]
                                if ($op) {
                                    $parsed = 0
                                    if ([int]::TryParse([string]$op.Value, [ref]$parsed) -and $parsed -ge 1) {
                                        $cadenceDays  = $parsed
                                        $fromOverride = $true
                                    }
                                }
                            }
                            if (-not $fromOverride) {
                                try {
                                    $meta = Get-PackagerMetadata -Path $path
                                    if ($meta.UpdateCadenceDays -and [int]$meta.UpdateCadenceDays -ge 1) {
                                        $cadenceDays = [int]$meta.UpdateCadenceDays
                                    }
                                } catch { }
                            }
                            try {
                                $nextDue = ([datetime]$lastChecked).ToUniversalTime().AddDays($cadenceDays)
                                if ($nextDue -gt (Get-Date).ToUniversalTime()) {
                                    $row.Status = 'Skipped (cadence)'
                                    [void]$State.LogQueue.Enqueue(('Skipped {0} (within {1}d cadence)' -f $app, $cadenceDays))
                                    $counts['Skipped']++
                                    continue
                                }
                            } catch { }
                        }

                        # 1. Check latest
                        $row.Status = 'Checking latest...'
                        [void]$State.LogQueue.Enqueue(('One Click: {0} ({1})' -f $app, $scrName))

                        $latest = $null
                        try {
                            $latest = Invoke-PackagerGetLatestVersion `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -FileServerPath $Ctx.FileShareRoot `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode
                            $row.LatestVersion = $latest
                            [void]$State.LogQueue.Enqueue(('Latest: ' + $latest))
                        } catch {
                            $row.Status = 'Check error'
                            [void]$State.LogQueue.Enqueue(('Latest check failed: ' + $_.Exception.Message))
                            $counts['CheckFailed']++
                            continue
                        }

                        # 1a. MECM pre-flight for Stage/StageAndPackage
                        if ($Ctx.Action -in @('Stage','StageAndPackage')) {
                            $cmName = [string]$row.CMName
                            if (-not [string]::IsNullOrWhiteSpace($cmName) -and $Ctx.AdminUiFound) {
                                try {
                                    $mecmRes = Get-MecmCurrentVersionByCMName -SiteCode $Ctx.SiteCode -ProviderMachineName $Ctx.ProviderMachineName -CMName $cmName
                                    if ($mecmRes.Found -and -not [string]::IsNullOrWhiteSpace([string]$mecmRes.SoftwareVersion)) {
                                        $row.CurrentVersion = [string]$mecmRes.SoftwareVersion
                                        $cmp = Compare-SemVer -A ([string]$mecmRes.SoftwareVersion) -B $latest
                                        if ($cmp -eq 0) {
                                            $row.Status = 'Up to date (MECM)'
                                            [void]$State.LogQueue.Enqueue(('MECM already has {0} at {1} - skipping' -f $app, $latest))
                                            try { Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result NoChange } catch { }
                                            $counts['NoChange']++
                                            continue
                                        }
                                    }
                                } catch {
                                    [void]$State.LogQueue.Enqueue(('MECM pre-flight for {0} failed: {1}' -f $app, $_.Exception.Message))
                                }
                            }
                        }

                        $versionChanged = (-not $lastKnown) -or ($lastKnown -ne $latest)
                        $neverStaged    = ($Ctx.Action -eq 'Stage'           -and -not $lastStaged)
                        $neverPackaged  = ($Ctx.Action -eq 'StageAndPackage' -and -not $lastPackaged)
                        $shouldAct      = $versionChanged -or $Ctx.ForceFlag -or $neverStaged -or $neverPackaged

                        $histResult = if ($versionChanged) { 'Updated' } else { 'NoChange' }
                        try {
                            Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result $histResult
                            $row.LastChecked = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                        } catch { }

                        if ($Ctx.Action -eq 'Report') {
                            $row.Status = if ($versionChanged) { 'Update available' } else { 'Up to date' }
                            if ($versionChanged) { $counts['Reported']++ } else { $counts['NoChange']++ }
                            continue
                        }

                        if (-not $shouldAct) {
                            $row.Status = 'Up to date'
                            [void]$State.LogQueue.Enqueue(('No change - skipping: ' + $app))
                            $counts['NoChange']++
                            continue
                        }

                        # 2. Stage
                        $State.Step = ('One Click {0}/{1}: staging {2}' -f $i, $n, $app)
                        $row.Status = 'Staging...'
                        [void]$State.LogQueue.Enqueue(('Stage: ' + $app))

                        $stageOk = $false
                        try {
                            $stg = Invoke-PackagerStage `
                                -PackagerPath $path `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -SevenZipPath $Ctx.SevenZipPath

                            if ($stg.ExitCode -eq 0) {
                                $stageOk = $true
                                $row.Status = 'Staged'
                                [void]$State.LogQueue.Enqueue(('Staged. Logs: ' + (Split-Path -Leaf $stg.OutLog)))
                                try {
                                    if ($latest) { Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $latest -Result Updated }
                                    else         { Update-PackagerHistory -PackagerName $baseName -Event Staged -Result Updated }
                                } catch { }
                            } else {
                                $row.Status = 'Stage error'
                                $stderrLines = @($stg.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                if ($stderrLines.Count -gt 0) {
                                    $linesToShow = [Math]::Min($stderrLines.Count, 10)
                                    for ($k = 0; $k -lt $linesToShow; $k++) {
                                        [void]$State.LogQueue.Enqueue(('  stderr: ' + $stderrLines[$k]))
                                    }
                                } else {
                                    [void]$State.LogQueue.Enqueue(('Stage exit code {0}, no stderr.' -f $stg.ExitCode))
                                }
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $stg.OutLog)))
                            }
                        } catch {
                            $row.Status = 'Stage error'
                            [void]$State.LogQueue.Enqueue(('Stage exception: ' + $_.Exception.Message))
                        }

                        if (-not $stageOk) {
                            $counts['Failed']++
                            continue
                        }

                        if ($Ctx.Action -eq 'Stage') {
                            $counts['Staged']++
                            continue
                        }

                        # 3. Package (StageAndPackage only)
                        $State.Step = ('One Click {0}/{1}: packaging {2}' -f $i, $n, $app)
                        $row.Status = 'Packaging...'
                        [void]$State.LogQueue.Enqueue(('Package: ' + $app))

                        try {
                            $reqJson = ''
                            if ($Ctx.RequirementsByApp) { $reqJson = [string]$Ctx.RequirementsByApp[$baseName] }
                            $varJson = ''
                            if ($Ctx.VariantsByApp) { $varJson = [string]$Ctx.VariantsByApp[$baseName] }
                            $cmdJson = ''
                            if ($Ctx.CommandsByApp) { $cmdJson = [string]$Ctx.CommandsByApp[$baseName] }
                            $pkg = Invoke-PackagerPackage `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -ProviderMachineName $Ctx.ProviderMachineName `
                                -Comment $Ctx.Comment `
                                -FileServerPath $Ctx.FileShareRoot `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -EstimatedRuntimeMins $Ctx.EstimatedRuntimeMins `
                                -MaximumRuntimeMins $Ctx.MaximumRuntimeMins `
                                -SevenZipPath $Ctx.SevenZipPath `
                                -CreateIntuneWin:([bool]$Ctx.IntuneWinCreate) `
                                -IntuneWinToolPath ([string]$Ctx.IntuneWinToolPath) `
                                -ContentLayout ([string]$Ctx.ContentLayout) `
                                -RequirementsJson $reqJson `
                                -VariantsJson $varJson `
                                -CommandsJson $cmdJson

                            if ($pkg.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $pkg.OutLog)))
                                if ($pkg.PSObject.Properties['IntuneWin'] -and $pkg.IntuneWin) {
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $pkg.IntuneWin.Message))
                                    if ($pkg.IntuneWin.NetworkPath) {
                                        [void]$State.LogQueue.Enqueue(('Intunewin on network: ' + $pkg.IntuneWin.NetworkPath))
                                    }
                                }
                                try {
                                    if ($latest) { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $latest -Result Updated }
                                    else         { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Result Updated }
                                } catch { }
                                $counts['StageAndPackage']++
                            } else {
                                $row.Status = 'Package error'
                                $stderrLines = @($pkg.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                if ($stderrLines.Count -gt 0) {
                                    $linesToShow = [Math]::Min($stderrLines.Count, 10)
                                    for ($k = 0; $k -lt $linesToShow; $k++) {
                                        [void]$State.LogQueue.Enqueue(('  stderr: ' + $stderrLines[$k]))
                                    }
                                } else {
                                    [void]$State.LogQueue.Enqueue(('Package exit code {0}, no stderr.' -f $pkg.ExitCode))
                                }
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $pkg.OutLog)))
                                $counts['Failed']++
                            }
                        } catch {
                            $row.Status = 'Package error'
                            [void]$State.LogQueue.Enqueue(('Package exception: ' + $_.Exception.Message))
                            $counts['Failed']++
                        }
                    }
                }
            }
            $State.Counts = $counts
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($Operation).AddArgument($rowsArray).AddArgument($Context).AddArgument($script:BgState)

    $script:BgHandle = $script:BgPS.BeginInvoke()

    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:BgTimer.Add_Tick({
        # Drain log queue onto the UI thread.
        if ($script:BgState -and $script:BgState.LogQueue) {
            $line = $null
            while ($script:BgState.LogQueue.TryDequeue([ref]$line)) {
                Add-LogLine -Message $line
            }
        }
        if ($script:BgState) {
            $cur = [string]$script:BgState.Step
            if ($txtProgressStep.Text -ne $cur) { $txtProgressStep.Text = $cur }
            if ([bool]$script:BgState.CancelRequested) {
                $btnPausePipeline.IsEnabled = $false
                $btnCancelPipeline.IsEnabled = $false
            }
            elseif ([bool]$script:BgState.Paused) {
                $btnPausePipeline.Content = 'Resume'
                $btnPausePipeline.IsEnabled = $true
                $btnCancelPipeline.IsEnabled = $true
            }
            else {
                $btnPausePipeline.Content = 'Pause'
                $btnPausePipeline.IsEnabled = $true
                $btnCancelPipeline.IsEnabled = $true
            }
        }
        # Re-render the grid so row.Status flips done in the bg are visible.
        try { $dataGrid.Items.Refresh() } catch { }

        if ($script:BgState -and $script:BgState.Done) {
            $doneState = $script:BgState
            $script:BgTimer.Stop()
            try { [void]$script:BgPS.EndInvoke($script:BgHandle) } catch { $null = $_ }
            try { $script:BgPS.Dispose() } catch { $null = $_ }
            $script:BgPS     = $null
            $script:BgHandle = $null

            # Final drain (anything enqueued after the last tick before Done).
            $line = $null
            if ($doneState.LogQueue) {
                while ($doneState.LogQueue.TryDequeue([ref]$line)) {
                    Add-LogLine -Message $line
                }
            }

            if ($doneState.ErrorMsg) {
                Add-LogLine -Message ('Pipeline failed: ' + $doneState.ErrorMsg)
                $txtStatus.Text = 'Failed.'
            } else {
                if ($doneState.Counts) {
                    $summaryEntries = @($doneState.Counts.GetEnumerator() | Where-Object { $_.Value -gt 0 })
                    if ($summaryEntries.Count -gt 0) {
                        $summaryLabel = switch ($Operation) {
                            'CheckLatest' { 'Check Latest summary:'; break }
                            'Stage'       { 'Stage summary:'; break }
                            'Package'     { 'Package summary:'; break }
                            'FullRun'     { 'One Click summary:'; break }
                            default       { 'Operation summary:'; break }
                        }
                        Add-LogSeparator
                        Add-LogLine -Message $summaryLabel
                        foreach ($entry in $summaryEntries) {
                            Add-LogLine -Message ('  {0,-18} {1}' -f $entry.Key, $entry.Value)
                        }
                    }
                }
                if ([bool]$doneState.Canceled) {
                    $txtStatus.Text = 'Canceled.'
                }
                else {
                    $txtStatus.Text = 'Complete.'
                }
            }

            try { $dataGrid.Items.Refresh() } catch { }
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $window.Cursor = $null
            Set-ActionButtonsEnabled -Enabled $true
            $btnPausePipeline.Content = 'Pause'
            $btnPausePipeline.IsEnabled = $true
            $btnCancelPipeline.IsEnabled = $true
            $script:BgTimer = $null
            $script:BgState = $null
        }
    })
    $script:BgTimer.Start()
}

# =============================================================================
# Drop-to-package intake: drag an installer onto the window, confirm the
# analyzed manifest, stage/package through the shared ad-hoc pipeline.
# =============================================================================

function Show-DropIntakeDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)]$Analysis,
        [bool]$PackageAvailable,
        [string]$PackageUnavailableReason
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Installer Drop"
    Width="620" Height="560"
    MinWidth="560" MinHeight="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,16,20,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" x:Name="txtDropFile" FontSize="14" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
        <TextBlock Grid.Row="1" x:Name="txtDropDetected" Margin="0,4,0,12" Opacity="0.75"/>

        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" MinWidth="130"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Application" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtDropAppName" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Publisher" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtDropPublisher" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="2" Grid.Column="0" Text="Version" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtDropVersion" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="3" Grid.Column="0" Text="Install args" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtDropInstallArgs" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="4" Grid.Column="0" x:Name="lblDropUninstall" Text="Uninstall command" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="4" Grid.Column="1" x:Name="txtDropUninstallCmd" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="5" Grid.Column="0" Text="Detection" VerticalAlignment="Top" Margin="0,4,12,0"/>
            <TextBlock Grid.Row="5" Grid.Column="1" x:Name="txtDropDetection" TextWrapping="Wrap" Opacity="0.75" Margin="0,4,0,0"/>
        </Grid>

        <CheckBox Grid.Row="3" x:Name="chkDropConfirm" Margin="0,12,0,0"
                  Content="I verified the predicted silent switches and detection for this installer"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnDropSavePackager" Content="Save as Packager" MinWidth="130" Height="32" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource MahApps.Styles.Button.Square}"/>
            <Button x:Name="btnDropStage" Content="Stage" MinWidth="90" Height="32" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource MahApps.Styles.Button.Square}"/>
            <Button x:Name="btnDropStagePackage" Content="Stage + Package" MinWidth="130" Height="32" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"/>
            <Button x:Name="btnDropCancel" Content="Cancel" MinWidth="90" Height="32" IsCancel="True" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource MahApps.Styles.Button.Square}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $txtFile      = $dlg.FindName('txtDropFile')
    $txtDetected  = $dlg.FindName('txtDropDetected')
    $txtAppName   = $dlg.FindName('txtDropAppName')
    $txtPublisher = $dlg.FindName('txtDropPublisher')
    $txtVersion   = $dlg.FindName('txtDropVersion')
    $txtArgs      = $dlg.FindName('txtDropInstallArgs')
    $lblUninst    = $dlg.FindName('lblDropUninstall')
    $txtUninst    = $dlg.FindName('txtDropUninstallCmd')
    $txtDetect    = $dlg.FindName('txtDropDetection')
    $chkConfirm   = $dlg.FindName('chkDropConfirm')
    $btnSave      = $dlg.FindName('btnDropSavePackager')
    $btnStageOnly = $dlg.FindName('btnDropStage')
    $btnStagePkg  = $dlg.FindName('btnDropStagePackage')
    $btnCancelDlg = $dlg.FindName('btnDropCancel')

    $isMsi = ([string]$Analysis.InstallerType -eq 'MSI')

    $txtFile.Text     = [string]$Analysis.FileName
    $txtDetected.Text = ('Detected: {0}   Confidence: {1}   Architecture: {2}' -f `
        $Analysis.InstallerType, $Analysis.Confidence, $Analysis.Architecture)
    $txtAppName.Text   = [string]$Analysis.AppName
    $txtPublisher.Text = [string]$Analysis.Publisher
    $txtVersion.Text   = [string]$Analysis.SoftwareVersion
    $txtArgs.Text      = [string]$Analysis.InstallArgs
    $txtUninst.Text    = [string]$Analysis.UninstallCommand

    if ($isMsi) {
        # MSI wrappers always run msiexec /qn /norestart against the file;
        # detection is the ProductCode ARP key. Nothing to edit or confirm.
        $txtArgs.IsReadOnly = $true
        $txtUninst.Visibility = [System.Windows.Visibility]::Collapsed
        $lblUninst.Visibility = [System.Windows.Visibility]::Collapsed
        $chkConfirm.Visibility = [System.Windows.Visibility]::Collapsed
        $txtDetect.Text = 'Registry: ARP key for ProductCode ' + $Analysis.ProductCode + ' (DisplayVersion match)'
    }
    else {
        $predicted = [string]$Analysis.UninstallRegistryKey
        if ([string]::IsNullOrWhiteSpace($predicted)) { $predicted = 'ARP key derived from the application name (verify after first install)' }
        $txtDetect.Text = 'Registry (predicted): ' + $predicted + ' (DisplayVersion match)'
    }

    $updateGate = {
        $confirmed = $isMsi -or ($chkConfirm.IsChecked -eq $true)
        $btnStagePkg.IsEnabled = $confirmed -and $PackageAvailable
        if (-not $PackageAvailable) {
            $btnStagePkg.ToolTip = $PackageUnavailableReason
        }
        elseif (-not $confirmed) {
            $btnStagePkg.ToolTip = 'Confirm the predicted values first. Packaging a wrong silent switch deploys a broken app.'
        }
        else {
            $btnStagePkg.ToolTip = $null
        }
    }
    & $updateGate
    $chkConfirm.Add_Checked({ & $updateGate })
    $chkConfirm.Add_Unchecked({ & $updateGate })

    $script:DropIntakeResult = $null
    $readValues = {
        @{
            AppName          = $txtAppName.Text.Trim()
            Publisher        = $txtPublisher.Text.Trim()
            SoftwareVersion  = $txtVersion.Text.Trim()
            InstallArgs      = $txtArgs.Text.Trim()
            UninstallCommand = $txtUninst.Text.Trim()
        }
    }
    $validate = {
        $v = & $readValues
        if ([string]::IsNullOrWhiteSpace($v.AppName)) { return 'Application name is required.' }
        if ([string]::IsNullOrWhiteSpace($v.SoftwareVersion)) { return 'Version is required.' }
        if (-not $isMsi -and [string]::IsNullOrWhiteSpace($v.InstallArgs)) { return 'Install args are required for a non-MSI installer.' }
        return $null
    }
    $chooseAction = {
        param($action)
        $problem = & $validate
        if ($problem) {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Missing Value' -Message $problem -Buttons OK -Icon Warning)
            return
        }
        $script:DropIntakeResult = @{ Action = $action; Values = (& $readValues) }
        $dlg.Close()
    }

    $btnStageOnly.Add_Click({ & $chooseAction 'Stage' })
    $btnStagePkg.Add_Click({ & $chooseAction 'StageAndPackage' })
    $btnSave.Add_Click({ & $chooseAction 'SavePackager' })
    $btnCancelDlg.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:DropIntakeResult
}

function Invoke-AdHocPipeline {
    param(
        [Parameter(Mandatory)][array]$Jobs,
        [Parameter(Mandatory)][hashtable]$Context
    )

    Initialize-BackgroundWorker

    $script:BgGraveyard = @(Stop-SuiteBgWork -PowerShell $script:BgPS -Timer $script:BgTimer -Graveyard $script:BgGraveyard)
    $script:BgTimer  = $null
    $script:BgPS     = $null
    $script:BgHandle = $null
    $script:BgState  = $null

    $script:BgState = [hashtable]::Synchronized(@{
        Step            = 'Starting...'
        Done            = $false
        ErrorMsg        = $null
        Paused          = $false
        CancelRequested = $false
        Canceled        = $false
        LogQueue        = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
        Counts          = $null
    })

    Set-ActionButtonsEnabled -Enabled $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $txtProgressTitle.Text = 'Processing dropped installers'
    $txtProgressStep.Text  = 'Starting...'
    $btnPausePipeline.Content = 'Pause'
    $btnPausePipeline.IsEnabled = $true
    $btnCancelPipeline.IsEnabled = $true
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible

    $jobsArray = @($Jobs)

    $script:BgPS = [powershell]::Create()
    $script:BgPS.Runspace = $script:BgRunspace
    [void]$script:BgPS.AddScript({
        param($JobsIn, $Ctx, $State)

        $counts = [ordered]@{ Staged = 0; Packaged = 0; Failed = 0 }
        $cmConnected = $null

        try {
            $jobs = @($JobsIn)
            $n = $jobs.Count
            $i = 0
            foreach ($job in $jobs) {
                while ([bool]$State.Paused -and -not [bool]$State.CancelRequested) {
                    $State.Step = 'Paused before next installer'
                    Start-Sleep -Milliseconds 250
                }
                if ([bool]$State.CancelRequested) {
                    $State.Canceled = $true
                    [void]$State.LogQueue.Enqueue('Canceled. Stopped before the next installer.')
                    break
                }

                $i++
                $v = $job.Values
                $State.Step = ('Stage {0}/{1}: {2}' -f $i, $n, $v.AppName)
                try {
                    $stage = New-AdHocStage -Analysis $job.Analysis `
                        -DownloadRoot $Ctx.DownloadRoot `
                        -AppName $v.AppName -Publisher $v.Publisher `
                        -SoftwareVersion $v.SoftwareVersion `
                        -InstallArgs $v.InstallArgs `
                        -UninstallCommand $v.UninstallCommand
                    $counts['Staged']++
                    [void]$State.LogQueue.Enqueue(('Staged: {0} {1} -> {2}' -f $v.AppName, $v.SoftwareVersion, $stage.StagedPath))
                }
                catch {
                    $counts['Failed']++
                    [void]$State.LogQueue.Enqueue(('Stage failed: {0}: {1}' -f $v.AppName, $_.Exception.Message))
                    continue
                }

                if ($job.Action -ne 'StageAndPackage') { continue }

                if ($null -eq $cmConnected) {
                    $State.Step = 'Connecting to site ' + $Ctx.SiteCode
                    $cmConnected = [bool](Connect-CMSite -SiteCode $Ctx.SiteCode -ProviderMachineName $Ctx.ProviderMachineName)
                    if (-not $cmConnected) {
                        [void]$State.LogQueue.Enqueue('Site connection failed. Staged content is intact; packaging skipped.')
                    }
                }
                if (-not $cmConnected) {
                    $counts['Failed']++
                    continue
                }

                $State.Step = ('Package {0}/{1}: {2}' -f $i, $n, $v.AppName)
                try {
                    $app = Invoke-AdHocPackage -StagedPath $stage.StagedPath `
                        -VendorFolder $stage.VendorFolder -AppFolder $stage.AppFolder `
                        -FileServerPath $Ctx.FileShareRoot -SiteCode $Ctx.SiteCode `
                        -Comment $Ctx.Comment -ContentLayout $Ctx.ContentLayout `
                        -EstimatedRuntimeMins $Ctx.EstimatedRuntimeMins `
                        -MaximumRuntimeMins $Ctx.MaximumRuntimeMins
                    $counts['Packaged']++
                    [void]$State.LogQueue.Enqueue(('Packaged: {0} {1}' -f $v.AppName, $v.SoftwareVersion))
                    $null = $app
                }
                catch {
                    $counts['Failed']++
                    [void]$State.LogQueue.Enqueue(('Package failed: {0}: {1}' -f $v.AppName, $_.Exception.Message))
                }
            }
            $State.Counts = $counts
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($jobsArray).AddArgument($Context).AddArgument($script:BgState)

    $script:BgHandle = $script:BgPS.BeginInvoke()

    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:BgTimer.Add_Tick({
        if ($script:BgState -and $script:BgState.LogQueue) {
            $line = $null
            while ($script:BgState.LogQueue.TryDequeue([ref]$line)) {
                Add-LogLine -Message $line
            }
        }
        if ($script:BgState) {
            $cur = [string]$script:BgState.Step
            if ($txtProgressStep.Text -ne $cur) { $txtProgressStep.Text = $cur }
        }

        if ($script:BgState -and $script:BgState.Done) {
            $doneState = $script:BgState
            $script:BgTimer.Stop()
            try { [void]$script:BgPS.EndInvoke($script:BgHandle) } catch { $null = $_ }
            try { $script:BgPS.Dispose() } catch { $null = $_ }
            $script:BgPS     = $null
            $script:BgHandle = $null

            $line = $null
            if ($doneState.LogQueue) {
                while ($doneState.LogQueue.TryDequeue([ref]$line)) {
                    Add-LogLine -Message $line
                }
            }

            if ($doneState.ErrorMsg) {
                Add-LogLine -Message ('Drop intake failed: ' + $doneState.ErrorMsg)
                $txtStatus.Text = 'Failed.'
            }
            else {
                if ($doneState.Counts) {
                    $summaryEntries = @($doneState.Counts.GetEnumerator() | Where-Object { $_.Value -gt 0 })
                    if ($summaryEntries.Count -gt 0) {
                        Add-LogSeparator
                        Add-LogLine -Message 'Drop intake summary:'
                        foreach ($entry in $summaryEntries) {
                            Add-LogLine -Message ('  {0,-18} {1}' -f $entry.Key, $entry.Value)
                        }
                    }
                }
                $txtStatus.Text = if ([bool]$doneState.Canceled) { 'Canceled.' } else { 'Complete.' }
            }

            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $window.Cursor = $null
            Set-ActionButtonsEnabled -Enabled $true
            $btnPausePipeline.Content = 'Pause'
            $btnPausePipeline.IsEnabled = $true
            $btnCancelPipeline.IsEnabled = $true
            $script:BgTimer = $null
            $script:BgState = $null
        }
    })
    $script:BgTimer.Start()
}

function Invoke-DropIntake {
    param([Parameter(Mandatory)][string[]]$Paths)

    # The action buttons are disabled while a pipeline runs, but the window
    # itself still receives drops; starting a second pipeline here would
    # hard-stop the running one mid-flight.
    if ($script:BgState -and -not [bool]$script:BgState.Done) {
        [void](Show-ThemedMessage -Owner $window -Title 'Pipeline Running' `
            -Message 'A pipeline is already running. Wait for it to finish (or cancel it), then drop the installer again.' `
            -Buttons OK -Icon Info)
        return
    }

    $installers = @($Paths | Where-Object { $_ -match '\.(msi|exe)$' -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    $ignored = @($Paths).Count - $installers.Count
    if ($ignored -gt 0) {
        Add-LogLine -Message ("Ignored {0} dropped item(s): only .msi and .exe files are supported." -f $ignored)
    }
    if ($installers.Count -eq 0) { return }

    if ([string]::IsNullOrWhiteSpace($script:Prefs.DownloadRoot)) {
        [void](Show-ThemedMessage -Owner $window -Title 'Download Root Required' `
            -Message 'Staging a dropped installer needs a Download Root. Open OPTIONS -> MECM Preferences to configure it.' `
            -Buttons OK -Icon Warning)
        return
    }

    # Packaging prerequisites decide whether Stage + Package is offered at all.
    $packageAvailable = $true
    $packageReason = ''
    if (-not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        $packageAvailable = $false; $packageReason = 'The Configuration Manager Console is not detected on this workstation.'
    }
    elseif ([string]::IsNullOrWhiteSpace($script:Prefs.SiteCode)) {
        $packageAvailable = $false; $packageReason = 'SiteCode is not configured (OPTIONS -> MECM Preferences).'
    }
    elseif ([string]::IsNullOrWhiteSpace($script:Prefs.FileShareRoot)) {
        $packageAvailable = $false; $packageReason = 'File Share Root is not configured (OPTIONS -> MECM Preferences).'
    }

    $jobs = @()
    foreach ($installer in $installers) {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $analysis = Get-InstallerAnalysis -Path $installer
        }
        catch {
            $window.Cursor = $null
            Add-LogLine -Message ('Analysis failed for {0}: {1}' -f (Split-Path -Leaf $installer), $_.Exception.Message)
            continue
        }
        $window.Cursor = $null
        Add-LogLine -Message ('Analyzed drop: {0} ({1}, {2})' -f $analysis.FileName, $analysis.InstallerType, $analysis.Confidence)

        $choice = Show-DropIntakeDialog -Owner $window -Analysis $analysis `
            -PackageAvailable $packageAvailable -PackageUnavailableReason $packageReason
        if (-not $choice) {
            Add-LogLine -Message ('Skipped: ' + $analysis.FileName)
            continue
        }

        if ($choice.Action -eq 'SavePackager') {
            try {
                $generated = New-PackagerFromDrop -Analysis $analysis `
                    -PackagersRoot (Join-Path $PSScriptRoot 'Packagers') `
                    -AppName $choice.Values.AppName -Publisher $choice.Values.Publisher `
                    -SoftwareVersion $choice.Values.SoftwareVersion
                Add-LogLine -Message ('Packager written: {0}. Fill in the download source before first use.' -f (Split-Path -Leaf $generated))
                Invoke-RefreshGrid
            }
            catch {
                [void](Show-ThemedMessage -Owner $window -Title 'Save Failed' -Message $_.Exception.Message -Buttons OK -Icon Error)
            }
            continue
        }

        $jobs += ,@{ Analysis = $analysis; Values = $choice.Values; Action = $choice.Action }
    }

    if ($jobs.Count -eq 0) { return }

    $txtStatus.Text = 'Processing dropped installers...'
    Invoke-AdHocPipeline -Jobs $jobs -Context @{
        DownloadRoot         = $script:Prefs.DownloadRoot
        SiteCode             = $script:Prefs.SiteCode
        ProviderMachineName  = $script:Prefs.ProviderMachineName
        FileShareRoot        = $script:Prefs.FileShareRoot
        ContentLayout        = $script:Prefs.ContentLayout
        Comment              = $txtComment.Text.Trim()
        EstimatedRuntimeMins = $script:Prefs.EstimatedRuntimeMins
        MaximumRuntimeMins   = $script:Prefs.MaximumRuntimeMins
    }
}

$script:PendingDropQueue = New-Object System.Collections.Queue
$window.AllowDrop = $true
$window.Add_PreviewDragOver({
    param($senderObj, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
        $e.Handled = $true
    }
})
$window.Add_PreviewDrop({
    param($senderObj, $e)
    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
    $e.Handled = $true
    # Queue per drop: a shared last-writer-wins variable would lose the
    # first drop if two land before the dispatcher drains.
    $script:PendingDropQueue.Enqueue(@($e.Data.GetData([System.Windows.DataFormats]::FileDrop)))
    # Defer past the drag-drop callback so the intake dialog does not block
    # the OLE drop source (Explorer hangs until DoDragDrop returns).
    [void]$window.Dispatcher.BeginInvoke([action]{
        if ($script:PendingDropQueue.Count -gt 0) {
            Invoke-DropIntake -Paths ([string[]]$script:PendingDropQueue.Dequeue())
        }
    })
})

# =============================================================================
# Action button handlers
# =============================================================================

# --- 1. Check Latest ---
$btnCheckLatest.Add_Click({
    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    $txtStatus.Text = "Checking latest versions for selected packagers..."
    Invoke-MultiAppPipeline -Operation CheckLatest -Rows $selectedRows -Context @{
        SiteCode       = $siteCodeValue
        FileShareRoot  = $script:Prefs.FileShareRoot
        DownloadRoot   = $script:Prefs.DownloadRoot
        M365Channel    = $script:Prefs.M365Channel
        M365DeployMode = $script:Prefs.M365DeployMode
        SevenZipPath   = Get-SevenZipPathForContext
    }
})

# --- 2. Check MECM ---
$btnCheckMECM.Add_Click({
    if (-not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Check MECM requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and reboot if you just installed) before running Check MECM." `
            -Buttons OK -Icon Warning)
        return
    }

    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    Set-ActionButtonsEnabled -Enabled $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait

    try {
        $txtStatus.Text = "Querying MECM for selected products..."

        foreach ($row in $selectedRows) {
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )

            $app    = [string]$row.Application
            $cmName = [string]$row.CMName

            Add-LogLine -Message ("MECM: {0}" -f $app)
            $row.Status = "Querying MECM..."
            $dataGrid.Items.Refresh()

            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -ProviderMachineName $script:Prefs.ProviderMachineName -CMName $cmName

                if (-not $res.Found) {
                    $row.CurrentVersion = ""
                    $row.Status = "Not found in MECM"
                    Add-LogLine -Message "Not found."
                    continue
                }

                $row.CurrentVersion = [string]$res.SoftwareVersion

                $latest = [string]$row.LatestVersion
                if (-not [string]::IsNullOrWhiteSpace($latest) -and -not [string]::IsNullOrWhiteSpace($res.SoftwareVersion)) {
                    $cmp = Compare-SemVer -A ([string]$res.SoftwareVersion) -B $latest
                    if ($cmp -lt 0)      { $row.Status = "Update available" }
                    elseif ($cmp -eq 0)  { $row.Status = "Up to date" }
                    else                 { $row.Status = "Current newer" }
                }
                else {
                    $row.Status = "MECM version retrieved"
                }

                if ($res.MatchCount -gt 1) {
                    Add-LogLine -Message ("Found {0} matches; using: {1} ({2})" -f $res.MatchCount, $res.DisplayName, $res.SoftwareVersion)
                }
                else {
                    Add-LogLine -Message ("Current version: {0}" -f $res.SoftwareVersion)
                }
            }
            catch {
                $row.Status = "Error"
                Add-LogLine -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        Select-OnlyUpdateAvailable
        $dataGrid.Items.Refresh()

        # Auto-discovery: offer to hide apps not found in MECM
        if (@($script:Prefs.HiddenApplications).Count -eq 0) {
            $notFound = @()
            foreach ($item in $script:PackagerData) {
                if ([string]$item.Status -eq "Not found in MECM") {
                    $notFound += [string]$item.Script
                }
            }
            if ($notFound.Count -gt 0 -and $notFound.Count -lt $script:PackagerData.Count) {
                $answer = Show-ThemedMessage -Owner $window -Title "Hide Unused Applications" `
                    -Message ("{0} application(s) were not found in MECM.`n`nHide them from the grid? You can change this later via Product Filter." -f $notFound.Count) `
                    -Buttons YesNo -Icon Question
                if ($answer -eq 'Yes') {
                    $script:Prefs.HiddenApplications = $notFound
                    Save-Preferences -Prefs $script:Prefs
                    Invoke-RefreshGrid
                    Add-LogLine -Message ("{0} application(s) hidden. Manage via Product Filter." -f $notFound.Count)
                }
            }
        }

        $txtStatus.Text = "MECM query complete."
    }
    finally {
        $window.Cursor = $null
        Set-ActionButtonsEnabled -Enabled $true
    }
})

# --- 3. Stage Packages ---
$btnStage.Add_Click({
    $dlRootValue = $script:Prefs.DownloadRoot
    if ([string]::IsNullOrWhiteSpace($dlRootValue)) {
        Add-LogLine -Message "Download Root is required for staging. Open Preferences to configure."
        $txtStatus.Text = "Download Root is required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    $txtStatus.Text = "Staging selected packages..."
    Invoke-MultiAppPipeline -Operation Stage -Rows $selectedRows -Context @{
        DownloadRoot   = $dlRootValue
        M365Channel    = $script:Prefs.M365Channel
        M365DeployMode = $script:Prefs.M365DeployMode
        LogFolder      = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath   = Get-SevenZipPathForContext
    }
})

# --- 4. Package Apps ---
$btnPackage.Add_Click({
    if (-not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Package requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and reboot if you just installed) before packaging." `
            -Buttons OK -Icon Warning)
        return
    }

    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $fsPathValue = $script:Prefs.FileShareRoot
    if ([string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -Message "File Share Root is required. Open Preferences to configure."
        $txtStatus.Text = "File Share Root is required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    $txtStatus.Text = "Packaging selected applications..."
    Invoke-MultiAppPipeline -Operation Package -Rows $selectedRows -Context @{
        SiteCode             = $siteCodeValue
        ProviderMachineName  = $script:Prefs.ProviderMachineName
        Comment              = $txtComment.Text.Trim()
        FileShareRoot        = $fsPathValue
        ContentLayout        = $script:Prefs.ContentLayout
        DownloadRoot         = $script:Prefs.DownloadRoot
        M365Channel          = $script:Prefs.M365Channel
        M365DeployMode       = $script:Prefs.M365DeployMode
        EstimatedRuntimeMins = $script:Prefs.EstimatedRuntimeMins
        MaximumRuntimeMins   = $script:Prefs.MaximumRuntimeMins
        LogFolder            = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath         = Get-SevenZipPathForContext
        IntuneWinCreate      = ([bool]$script:Prefs.Intune.CreateIntuneWin -and -not [string]::IsNullOrWhiteSpace((Get-IntuneWinToolPathForContext)))
        IntuneWinToolPath    = Get-IntuneWinToolPathForContext
        RequirementsByApp    = Get-RequirementsMapForContext
        VariantsByApp        = Get-VariantsMapForContext
        CommandsByApp        = Get-CommandsMapForContext
    }
})

# --- 5. Full Run (one-click tracked-apps flow) ---
# Thin dispatch: validates prefs + ConfigMgr availability + tracked set, then
# routes to Invoke-MultiAppPipeline -Operation FullRun. The bg scriptblock
# there mirrors the original per-row cadence / MECM pre-flight / Stage /
# Package logic so history entries and row.Status flips stay identical.
$btnFullRun.Add_Click({
    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open MECM Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $actionPlanned = $script:Prefs.AppFlow.Action
    if ($actionPlanned -eq 'StageAndPackage' -and -not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "One Click with Stage and Package requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and reboot if you just installed) before running Stage and Package, or switch One Click Settings action to Report or Stage." `
            -Buttons OK -Icon Warning)
        return
    }

    $trackedBases = @($script:Prefs.AppFlow.Tracked)
    if ($trackedBases.Count -eq 0) {
        Add-LogLine -Message "No apps are tracked for One Click. Open OPTIONS -> One Click Settings to configure."
        $txtStatus.Text = "No apps tracked."
        [void](Show-ThemedMessage -Owner $window -Title 'One Click Not Configured' `
            -Message "No apps are tracked yet.`n`nOpen OPTIONS (sidebar) and select One Click Settings, then choose which packagers to include, pick an action (Report / Stage / Stage and Package), and click OK." `
            -Buttons OK -Icon Info)
        return
    }

    $action       = $script:Prefs.AppFlow.Action
    $forceFlag    = [bool]$script:Prefs.AppFlow.ForceOnLaunch
    $fsPathValue  = $script:Prefs.FileShareRoot
    $dlRootValue  = $script:Prefs.DownloadRoot

    if ($action -eq 'StageAndPackage' -and [string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -Message ("File Share Root is required for action '{0}'. Open MECM Preferences." -f $action)
        $txtStatus.Text = "File Share Root is required."
        return
    }
    if ($action -in @('Stage','StageAndPackage') -and [string]::IsNullOrWhiteSpace($dlRootValue)) {
        Add-LogLine -Message ("Download Root is required for action '{0}'. Open MECM Preferences." -f $action)
        $txtStatus.Text = "Download Root is required."
        return
    }

    # Match tracked base names to currently-visible grid rows
    $trackedSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]$trackedBases,
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $rows = @($script:PackagerData | Where-Object {
        $trackedSet.Contains([System.IO.Path]::GetFileNameWithoutExtension([string]$_.Script))
    })
    if ($rows.Count -eq 0) {
        Add-LogLine -Message ("No tracked apps are visible in the grid. Check Product Filter.")
        $txtStatus.Text = "No visible tracked apps."
        return
    }

    Add-LogSeparator
    Add-LogLine -Message ("One Click: {0} app(s), action={1}{2}" -f $rows.Count, $action, $(if ($forceFlag) { ', force=on' } else { '' }))
    $txtStatus.Text = ("One Click: {0} app(s)..." -f $rows.Count)

    Invoke-MultiAppPipeline -Operation FullRun -Rows $rows -Context @{
        SiteCode             = $siteCodeValue
        ProviderMachineName  = $script:Prefs.ProviderMachineName
        Action               = $action
        ForceFlag            = $forceFlag
        Overrides            = $script:Prefs.AppFlow.CadenceOverrides
        Comment              = $txtComment.Text.Trim()
        FileShareRoot        = $fsPathValue
        ContentLayout        = $script:Prefs.ContentLayout
        DownloadRoot         = $dlRootValue
        M365Channel          = $script:Prefs.M365Channel
        M365DeployMode       = $script:Prefs.M365DeployMode
        EstimatedRuntimeMins = $script:Prefs.EstimatedRuntimeMins
        MaximumRuntimeMins   = $script:Prefs.MaximumRuntimeMins
        AdminUiFound         = $script:Prefs.DetectedTools.ConfigMgrConsole.Found
        LogFolder            = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath         = Get-SevenZipPathForContext
        IntuneWinCreate      = ([bool]$script:Prefs.Intune.CreateIntuneWin -and -not [string]::IsNullOrWhiteSpace((Get-IntuneWinToolPathForContext)))
        IntuneWinToolPath    = Get-IntuneWinToolPathForContext
        RequirementsByApp    = Get-RequirementsMapForContext
        VariantsByApp        = Get-VariantsMapForContext
        CommandsByApp        = Get-CommandsMapForContext
    }
})

# =============================================================================
# Window lifecycle
# =============================================================================
$window.Add_Loaded({
    Add-LogLine -Message ("Loading packagers from: {0}" -f $PackagersRoot)
    Invoke-RefreshGrid
    Add-LogLine -Message ("{0} packager(s) loaded. Ready." -f $script:PackagerData.Count)
})

$window.Add_Closing({
    Save-WindowState -Window $window -Path (Get-WindowStatePath) -ExtraState @{
        DarkTheme    = ($toggleTheme.IsOn -eq $true)
        DebugColumns = ($toggleDebugCols.IsOn -eq $true)
    }

    # Tear down the async pipeline without blocking shutdown: a stuck
    # pipeline stops asynchronously and the runspace closes async so the
    # bg thread cannot keep the process alive or freeze the close.
    $script:BgGraveyard = @(Stop-SuiteBgWork -PowerShell $script:BgPS -Timer $script:BgTimer -Graveyard $script:BgGraveyard)
    $script:BgTimer = $null
    $script:BgPS    = $null
    Close-SuiteBgRunspace -Runspace $script:BgRunspace
    $script:BgRunspace = $null
    $script:BgHandle = $null
    $script:BgState  = $null
})

# Defaults (overridden by Restore-WindowState if saved state exists)
$script:SavedDarkTheme = $true
$script:SavedDebugCols = $false

# Restore previous window position + saved preferences
Restore-WindowState -Window $window -Path (Get-WindowStatePath) -OnStateLoaded {
    param($s)
    $script:SavedDarkTheme = if ($null -ne $s.DarkTheme) { [bool]$s.DarkTheme } else { $true }
    $script:SavedDebugCols = if ($null -ne $s.DebugColumns) { [bool]$s.DebugColumns } else { $false }
}

# Apply saved theme and debug column state
if (-not $script:SavedDarkTheme) {
    $toggleTheme.IsOn = $false
    # Toggled event fires automatically and applies Light.Blue + button colors
}
if ($script:SavedDebugCols) {
    $toggleDebugCols.IsOn = $true
    # Toggled event fires automatically and shows debug columns
}

# =============================================================================
# Show window (blocks until closed)
# =============================================================================
[void]$window.ShowDialog()
