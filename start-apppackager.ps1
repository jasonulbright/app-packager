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
    Version    : 1.0.0
    Updated    : 2026-05-02
#>

param(
    [string]$SiteCode = "MCM",
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
        FileShareRoot        = "\\fileserver\sccm$"
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
        }
        ContentDistribution  = [pscustomobject]@{
            AutoDistribute = $false
            DPGroupName    = ''
        }
    }

    $path = Get-PreferencesPath
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($null -ne $data.SiteCode)             { $defaults.SiteCode             = [string]$data.SiteCode }
        if ($null -ne $data.FileShareRoot)         { $defaults.FileShareRoot        = [string]$data.FileShareRoot }
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
    catch { }
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

$script:Prefs = Read-Preferences

# Refresh tool detection once per launch. Persists into the same
# preferences JSON so the status is available immediately on next start.
try {
    $script:Prefs.DetectedTools.ConfigMgrConsole = Invoke-DetectConfigMgrConsole
    $script:Prefs.DetectedTools.SevenZipCli      = Invoke-DetectSevenZipCli
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
        [Parameter(Mandatory)][string[]]$Arguments
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
        $lines = Get-Content -LiteralPath $ScriptPath -TotalCount 120 -ErrorAction Stop
        foreach ($line in $lines) {
            if (-not $info.DownloadSubfolder -and $line -match '\$BaseDownloadRoot\s*=\s*Join-Path\s+\$DownloadRoot\s+"([^"]+)"') {
                $info.DownloadSubfolder = $matches[1]
            }
            if (-not $info.VendorFolder -and $line -match '^\s*\$VendorFolder\s*=\s*"([^"]+)"') {
                $info.VendorFolder = $matches[1]
            }
            if (-not $info.AppFolder -and $line -match '^\s*\$AppFolder\s*=\s*"([^"]+)"') {
                $info.AppFolder = $matches[1]
            }
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
        [string]$DownloadRoot = $null
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
        $networkContentPath = Join-Path (Join-Path (Join-Path $FileServerPath 'Applications') $info.VendorFolder) $info.AppFolder
        $networkContentPath = Join-Path $networkContentPath $manifest.SoftwareVersion
    }

    $comparison = Compare-StageFileHashes -Root $networkContentPath -Expected $manifest.FileHashes
    if (-not $comparison.Pass) {
        throw ("Package integrity verification failed: {0}" -f (Get-StageFileHashComparisonMessage -Comparison $comparison))
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

    try { Set-Location "${SiteCode}:" -ErrorAction Stop }
    catch { throw ("Failed to connect to CM site PSDrive '{0}:'" -f $SiteCode) }

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

function Invoke-ProcessWithStreaming {
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [Parameter(Mandatory)][string]$OutLog,
        [Parameter(Mandatory)][string]$ErrLog,
        [string]$StructuredLog = '',
        [System.Windows.Controls.TextBox]$LogTextBox = $null
    )

    $p = New-Object System.Diagnostics.Process
    try {
        $p.StartInfo = $StartInfo
        $null = $p.Start()

        $outLines = New-Object System.Collections.Generic.List[string]
        $errTask = $p.StandardError.ReadToEndAsync()
        $reader   = $p.StandardOutput
        $lineTask = $reader.ReadLineAsync()

        while ($true) {
            if ($lineTask.IsCompleted) {
                $line = $lineTask.Result
                if ($null -eq $line) { break }
                $outLines.Add($line)

                if ($LogTextBox) {
                    $displayLine = $line -replace '^\[[\d: -]+\] \[\w+\s*\] ', ''
                    if ($displayLine.Trim()) {
                        Add-LogLine -Message ("  {0}" -f $displayLine)
                    }
                }

                $lineTask = $reader.ReadLineAsync()
            }

            # WPF dispatcher pump
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )
            Start-Sleep -Milliseconds 50
        }

        if (-not $p.WaitForExit(15000)) {
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
        [AllowEmptyString()][string]$Comment = '',
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [string]$M365Channel = $null,
        [string]$M365DeployMode = $null,
        [int]$EstimatedRuntimeMins = 0,
        [int]$MaximumRuntimeMins = 0,
        [string]$SevenZipPath = '',
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
    $argsBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PackagerPath, '-PackageOnly', '-SiteCode', $SiteCode, '-Comment', $Comment, '-LogPath', $structuredLog)
    if (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath) {
        $argsBase += @('-FileServerPath', $FileServerPath)
    }
    if ($DownloadRoot) { $argsBase += @('-DownloadRoot', $DownloadRoot) }
    if ($M365Channel) { $argsBase += @('-M365Channel', $M365Channel) }
    if ($M365DeployMode) { $argsBase += @('-M365DeployMode', $M365DeployMode) }
    if ($EstimatedRuntimeMins -gt 0) { $argsBase += @('-EstimatedRuntimeMins', [string]$EstimatedRuntimeMins) }
    if ($MaximumRuntimeMins -gt 0) { $argsBase += @('-MaximumRuntimeMins', [string]$MaximumRuntimeMins) }
    Set-ProcessStartInfoArgumentList -StartInfo $psi -Arguments $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    Set-PackagerEnvironment -StartInfo $psi -SevenZipPath $SevenZipPath

    $result = Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
    Assert-PackagerPackageIntegrity -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot
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
        [string]$SevenZipPath
    )
    if (-not [string]::IsNullOrWhiteSpace($SevenZipPath)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_SEVENZIP'] = [string]$SevenZipPath
    }
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

function Select-OnlyUpdateAvailable {
    foreach ($item in $script:PackagerData) {
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

function Save-WindowState {
    param([Parameter(Mandatory)]$Window)

    $state = @{}
    if ($Window.WindowState -eq [System.Windows.WindowState]::Normal) {
        $state.Left   = [int]$Window.Left
        $state.Top    = [int]$Window.Top
        $state.Width  = [int]$Window.Width
        $state.Height = [int]$Window.Height
    }
    else {
        $state.Left   = [int]$Window.RestoreBounds.Left
        $state.Top    = [int]$Window.RestoreBounds.Top
        $state.Width  = [int]$Window.RestoreBounds.Width
        $state.Height = [int]$Window.RestoreBounds.Height
    }
    $state.Maximized = ($Window.WindowState -eq [System.Windows.WindowState]::Maximized)
    $state.DarkTheme = ($toggleTheme.IsOn -eq $true)
    $state.DebugColumns = ($toggleDebugCols.IsOn -eq $true)

    try {
        $json = $state | ConvertTo-Json
        Set-Content -LiteralPath (Get-WindowStatePath) -Value $json -Encoding UTF8
    }
    catch { }
}

function Restore-WindowState {
    param([Parameter(Mandatory)]$Window)

    $path = Get-WindowStatePath
    if (-not (Test-Path -LiteralPath $path)) { return }

    try {
        $state = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        $w = [int]$state.Width
        $h = [int]$state.Height
        if ($w -lt $Window.MinWidth)  { $w = [int]$Window.MinWidth }
        if ($h -lt $Window.MinHeight) { $h = [int]$Window.MinHeight }

        # Check if the saved position is visible on any monitor
        $screens = [System.Windows.Forms.Screen]::AllScreens
        $visible = $false
        foreach ($screen in $screens) {
            $titleBarRect = New-Object System.Drawing.Rectangle ([int]$state.Left), ([int]$state.Top), $w, 40
            if ($screen.WorkingArea.IntersectsWith($titleBarRect)) {
                $visible = $true
                break
            }
        }

        if ($visible) {
            $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $Window.Left   = [double]$state.Left
            $Window.Top    = [double]$state.Top
            $Window.Width  = [double]$w
            $Window.Height = [double]$h
        }

        if ($state.Maximized -eq $true) {
            $Window.WindowState = [System.Windows.WindowState]::Maximized
        }

        # Restore theme and debug column state (applied after controls are wired)
        $script:SavedDarkTheme = if ($null -ne $state.DarkTheme) { [bool]$state.DarkTheme } else { $true }
        $script:SavedDebugCols = if ($null -ne $state.DebugColumns) { [bool]$state.DebugColumns } else { $false }
    }
    catch { }
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
        [string]$FileServerPath,
        [string]$DownloadRoot,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30,
        [string]$Comment = '',
        [string]$SevenZipPath = '',
        [pscustomobject]$CadenceOverrides,
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

        # 1. Discover vendor's current version. On failure, don't poison
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
        if ($DownloadRoot)         { $pkgArgs += @('-DownloadRoot',         $DownloadRoot)         }
        if ($EstimatedRuntimeMins) { $pkgArgs += @('-EstimatedRuntimeMins', $EstimatedRuntimeMins) }
        if ($MaximumRuntimeMins)   { $pkgArgs += @('-MaximumRuntimeMins',   $MaximumRuntimeMins)   }
        if ($Comment)              { $pkgArgs += @('-Comment',              $Comment)              }
        if ($OnUpdateFound -eq 'Stage') { $pkgArgs += '-StageOnly' }

        try {
            $restoreSevenZipEnv = $false
            $previousSevenZipEnv = $null
            try {
                if (-not [string]::IsNullOrWhiteSpace($SevenZipPath)) {
                    $restoreSevenZipEnv = $true
                    $previousSevenZipEnv = $env:APP_PACKAGER_SEVENZIP
                    $env:APP_PACKAGER_SEVENZIP = $SevenZipPath
                }

                & powershell.exe @pkgArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    if ($line) { Write-Log ("[batch:{0}] {1}" -f $baseName, $line) -Level INFO }
                }
            }
            finally {
                if ($restoreSevenZipEnv) {
                    if ($null -ne $previousSevenZipEnv) {
                        $env:APP_PACKAGER_SEVENZIP = $previousSevenZipEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_SEVENZIP -ErrorAction SilentlyContinue
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
    $downloadRoot     = if ($prefs -and $prefs.DownloadRoot)              { $prefs.DownloadRoot }              else { $null }
    $cadenceOverrides = if ($prefs -and $prefs.AppFlow.CadenceOverrides)  { $prefs.AppFlow.CadenceOverrides }  else { $null }
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
        -FileServerPath    $fileServerPath `
        -DownloadRoot      $downloadRoot `
        -SevenZipPath      $sevenZipPath `
        -CadenceOverrides  $cadenceOverrides `
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
# Title-bar drag fallback. PS51-WPF-033.
# Some VS Code PowerShell launch contexts can leave MahApps' custom title
# thumb unable to initiate native window move. Install a WM_NCHITTEST hook
# returning HTCAPTION for the title band, plus a managed DragMove fallback
# for hosts where HwndSource cannot be hooked. Wire on every MetroWindow
# (main window and every modal popup).
# =============================================================================
$script:TitleBarHitTestWindows = @{}
$script:TitleBarHitTestHooks   = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $h = [double]$Window.TitleBarHeight
        if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h }
    } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch { $null = $_ }
}

function Remove-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
        }
        if ($script:TitleBarHitTestWindows.ContainsKey($key)) {
            $script:TitleBarHitTestWindows.Remove($key)
        }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

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

# Button color palettes
$script:DarkButtonBg      = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:LightWfBg         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')  # Windows blue
$script:LightWfBorder     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:LightOptBg        = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')  # Same Windows blue
$script:LightOptBorder    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')

$script:WorkflowButtons = @($btnFullRun, $btnCheckLatest, $btnCheckMECM, $btnStage, $btnPackage)
$script:OptionsButtons  = @($btnOptions)

# LOG OUTPUT label Foreground: #B0B0B0 passes AA on dark (#252525) at 7.07:1
# but fails AA on light (#FFFFFF) at 2.17:1. Apply per theme instead of
# hardcoding in XAML. See reference_srl_wpf_brand.md.
$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

function Set-ButtonTheme {
    param([bool]$IsDark)
    if ($IsDark) {
        foreach ($b in $script:WorkflowButtons) { $b.Background = $script:DarkButtonBg; $b.BorderBrush = $script:DarkButtonBorder }
        foreach ($b in $script:OptionsButtons)  { $b.Background = $script:DarkButtonBg; $b.BorderBrush = $script:DarkButtonBorder }
        if ($lblLogOutput) { $lblLogOutput.Foreground = $script:LogLabelDark }
    }
    else {
        foreach ($b in $script:WorkflowButtons) { $b.Background = $script:LightWfBg;  $b.BorderBrush = $script:LightWfBorder }
        foreach ($b in $script:OptionsButtons)  { $b.Background = $script:LightOptBg; $b.BorderBrush = $script:LightOptBorder }
        if ($lblLogOutput) { $lblLogOutput.Foreground = $script:LogLabelLight }
    }
}

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

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$toggleTheme.Add_Toggled({
    if ($toggleTheme.IsOn) {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, "Dark.Steel")
        $txtThemeLabel.Text = "Dark Theme"
        Set-ButtonTheme -IsDark $true
        # Reset title bar to theme default (Steel grey)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    }
    else {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, "Light.Blue")
        $txtThemeLabel.Text = "Light Theme"
        Set-ButtonTheme -IsDark $false
        # Override title bar to exact Windows blue, active and inactive
        $window.WindowTitleBrush = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
})

# =============================================================================
# DataGrid binding
# =============================================================================
$dataGrid.ItemsSource = $script:PackagerData

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
                foreach ($item in $script:PackagerData) { $item.Selected = $true }
                Add-LogLine -Message "Selected all rows."
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
function Show-ThemedMessage {
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('OK','YesNo')][string]$Buttons = 'OK',
        [ValidateSet('Info','Warning','Error','Question')][string]$Icon = 'Info'
    )

    $glyph = switch ($Icon) {
        'Info'     { 'i' }
        'Warning'  { '!' }
        'Error'    { 'X' }
        'Question' { '?' }
    }

    $estLines = [math]::Max(1, [math]::Ceiling($Message.Length / 60.0))
    $height   = [math]::Min(420, 180 + ($estLines * 22))

    $xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="$Title"
    Width="460" Height="$height"
    MinWidth="360" MinHeight="160"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="NoResize"
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
    <DockPanel Margin="16">
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0" x:Name="panelButtons"/>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="48"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Width="36" Height="36" CornerRadius="18"
                    VerticalAlignment="Top" HorizontalAlignment="Left"
                    Background="{DynamicResource MahApps.Brushes.Accent}">
                <TextBlock Text="$glyph" FontSize="20" FontWeight="Bold"
                           Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <TextBlock Grid.Column="1" x:Name="txtMessage"
                       FontSize="13" TextWrapping="Wrap" VerticalAlignment="Top"
                       Margin="4,2,0,0"/>
        </Grid>
    </DockPanel>
</Controls:MetroWindow>
"@

    [xml]$xmlDoc = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $win = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $win

    Set-DialogChromeFromOwner -Dialog $win -Owner $Owner

    $win.FindName('txtMessage').Text = $Message
    $panelButtons = $win.FindName('panelButtons')

    $script:ThemedMessageResult = 'Cancel'

    function New-ThemedMsgButton {
        # $ParentWindow is passed explicitly so the Add_Click closure captures
        # it as a function-local (rather than relying on dynamic scope lookup
        # of $win from the outer function, which GetNewClosure does not do).
        param(
            [string]$Content,
            [string]$ResultValue,
            [bool]$IsDefault,
            [bool]$IsCancel,
            [bool]$IsAccent,
            [Parameter(Mandatory)]$ParentWindow
        )
        $b = New-Object System.Windows.Controls.Button
        $b.Content  = $Content
        $b.MinWidth = 90
        $b.Height   = 32
        $b.Margin   = New-Object System.Windows.Thickness(0, 0, 8, 0)
        $styleKey = if ($IsAccent) { 'MahApps.Styles.Button.Square.Accent' } else { 'MahApps.Styles.Button.Square' }
        $b.SetResourceReference([System.Windows.Controls.Button]::StyleProperty, $styleKey)
        [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($b, [System.Windows.Controls.CharacterCasing]::Normal)
        if ($IsDefault) { $b.IsDefault = $true }
        if ($IsCancel)  { $b.IsCancel  = $true }
        $b.Add_Click({
            $script:ThemedMessageResult = $ResultValue
            $ParentWindow.Close()
        }.GetNewClosure())
        return $b
    }

    if ($Buttons -eq 'OK') {
        [void]$panelButtons.Children.Add((New-ThemedMsgButton -Content 'OK'  -ResultValue 'OK'  -IsDefault $true  -IsCancel $true  -IsAccent $true  -ParentWindow $win))
    } else {
        [void]$panelButtons.Children.Add((New-ThemedMsgButton -Content 'Yes' -ResultValue 'Yes' -IsDefault $true  -IsCancel $false -IsAccent $true  -ParentWindow $win))
        [void]$panelButtons.Children.Add((New-ThemedMsgButton -Content 'No'  -ResultValue 'No'  -IsDefault $false -IsCancel $true  -IsAccent $false -ParentWindow $win))
    }

    [void]$win.ShowDialog()
    return $script:ThemedMessageResult
}

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
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="140"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Text="Site Code:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtSC" Width="80" FontSize="13" HorizontalAlignment="Left" MaxLength="5" Margin="0,0,0,8" ToolTip="ConfigMgr site code PSDrive name (e.g., MCM)"/>

    <TextBlock Grid.Row="1" Grid.Column="0" Text="File Share Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtFS" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="UNC path to the SCCM content file share"/>

    <TextBlock Grid.Row="2" Grid.Column="0" Text="Download Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtDL" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="Local folder where installers are downloaded during staging"/>

    <TextBlock Grid.Row="3" Grid.Column="0" Text="Est. Runtime:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBox x:Name="txtEst" Width="60" FontSize="13" MaxLength="4" ToolTip="Estimated install runtime in minutes"/>
        <TextBlock Text=" mins" FontSize="13" VerticalAlignment="Center" Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
    </StackPanel>

    <TextBlock Grid.Row="4" Grid.Column="0" Text="Max Runtime:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="4" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBox x:Name="txtMax" Width="60" FontSize="13" MaxLength="4" ToolTip="Maximum allowed install runtime in minutes"/>
        <TextBlock Text=" mins" FontSize="13" VerticalAlignment="Center" Foreground="{DynamicResource MahApps.Brushes.Gray5}"/>
    </StackPanel>

    <TextBlock Grid.Row="5" Grid.Column="0" Text="Auto-distribute:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="When enabled, the Package phase calls Start-CMContentDistribution after creating each MECM Application."/>
    <CheckBox  Grid.Row="5" Grid.Column="1" x:Name="chkAutoDist" Content="Start-CMContentDistribution after Package" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="6" Grid.Column="0" Text="DP Group:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Exact name of the Distribution Point Group to target."/>
    <TextBox   Grid.Row="6" Grid.Column="1" x:Name="txtDPGroup" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="Distribution Point Group display name (e.g. 'All DPs')"/>

    <TextBlock Grid.Row="7" Grid.Column="0" Text="Console:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="Configuration Manager Console (AdminUI) detection status. Checked once per launch."/>
    <TextBlock Grid.Row="7" Grid.Column="1" x:Name="txtConsoleStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,6,0,0"/>

    <TextBlock Grid.Row="8" Grid.Column="0" Text="7-Zip CLI:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,6,0,0" ToolTip="7-Zip command-line (7z.exe) detection status. Required by Adobe Reader + TeamViewer Host packagers."/>
    <TextBlock Grid.Row="8" Grid.Column="1" x:Name="txtSevenZipStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,6,0,0"/>
</Grid>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $txtSC  = $element.FindName('txtSC')
    $txtFS  = $element.FindName('txtFS')
    $txtDL  = $element.FindName('txtDL')
    $txtEst = $element.FindName('txtEst')
    $txtMax = $element.FindName('txtMax')
    $chkAutoDist = $element.FindName('chkAutoDist')
    $txtDPGroup  = $element.FindName('txtDPGroup')
    $txtConsoleStatus  = $element.FindName('txtConsoleStatus')
    $txtSevenZipStatus = $element.FindName('txtSevenZipStatus')

    $txtSC.Text  = [string]$script:Prefs.SiteCode
    $txtFS.Text  = [string]$script:Prefs.FileShareRoot
    $txtDL.Text  = [string]$script:Prefs.DownloadRoot
    $txtEst.Text = [string]$script:Prefs.EstimatedRuntimeMins
    $txtMax.Text = [string]$script:Prefs.MaximumRuntimeMins
    $chkAutoDist.IsChecked = [bool]$script:Prefs.ContentDistribution.AutoDistribute
    $txtDPGroup.Text       = [string]$script:Prefs.ContentDistribution.DPGroupName

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

    # Closure captures panel-local controls by value. Prefs ref is captured too
    # so the commit writes to the live $script:Prefs without needing $script:
    # scope resolution from inside the closure (which can be unreliable).
    $prefsRef = $script:Prefs
    $commit = {
        $estVal = 15; $maxVal = 30
        if (-not [int]::TryParse($txtEst.Text.Trim(), [ref]$estVal)) { $estVal = 15 }
        if (-not [int]::TryParse($txtMax.Text.Trim(), [ref]$maxVal)) { $maxVal = 30 }

        $prefsRef.SiteCode             = $txtSC.Text.Trim()
        $prefsRef.FileShareRoot        = $txtFS.Text.Trim()
        $prefsRef.DownloadRoot         = $txtDL.Text.Trim()
        $prefsRef.EstimatedRuntimeMins = $estVal
        $prefsRef.MaximumRuntimeMins   = $maxVal
        $prefsRef.ContentDistribution.AutoDistribute = [bool]$chkAutoDist.IsChecked
        $prefsRef.ContentDistribution.DPGroupName    = $txtDPGroup.Text.Trim()
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

    # All four panels live in the unified Options window now.
    $panels = @(
        (New-MecmPreferencesPanel),
        (New-PackagerPreferencesPanel),
        (New-AppFlowPanel),
        (New-ProductFilterPanel)
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

    # Cancel any in-flight pipeline. Stop is best-effort; the current
    # packager may still finish its current step before yielding.
    if ($script:BgTimer) {
        try { $script:BgTimer.Stop() } catch { $null = $_ }
        $script:BgTimer = $null
    }
    if ($script:BgPS)    {
        try { [void]$script:BgPS.Stop() } catch { $null = $_ }
        try { $script:BgPS.Dispose() }   catch { $null = $_ }
        $script:BgPS = $null
    }
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
                            $res = Invoke-PackagerPackage `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -Comment $Ctx.Comment `
                                -FileServerPath $Ctx.FileShareRoot `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -EstimatedRuntimeMins $Ctx.EstimatedRuntimeMins `
                                -MaximumRuntimeMins $Ctx.MaximumRuntimeMins `
                                -SevenZipPath $Ctx.SevenZipPath

                            if ($res.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $res.OutLog)))
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
                                    $mecmRes = Get-MecmCurrentVersionByCMName -SiteCode $Ctx.SiteCode -CMName $cmName
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
                            $pkg = Invoke-PackagerPackage `
                                -PackagerPath $path `
                                -SiteCode $Ctx.SiteCode `
                                -Comment $Ctx.Comment `
                                -FileServerPath $Ctx.FileShareRoot `
                                -LogFolder $Ctx.LogFolder `
                                -DownloadRoot $Ctx.DownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -EstimatedRuntimeMins $Ctx.EstimatedRuntimeMins `
                                -MaximumRuntimeMins $Ctx.MaximumRuntimeMins `
                                -SevenZipPath $Ctx.SevenZipPath

                            if ($pkg.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $pkg.OutLog)))
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
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -CMName $cmName

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
        Comment              = $txtComment.Text.Trim()
        FileShareRoot        = $fsPathValue
        DownloadRoot         = $script:Prefs.DownloadRoot
        M365Channel          = $script:Prefs.M365Channel
        M365DeployMode       = $script:Prefs.M365DeployMode
        EstimatedRuntimeMins = $script:Prefs.EstimatedRuntimeMins
        MaximumRuntimeMins   = $script:Prefs.MaximumRuntimeMins
        LogFolder            = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath         = Get-SevenZipPathForContext
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
        Action               = $action
        ForceFlag            = $forceFlag
        Overrides            = $script:Prefs.AppFlow.CadenceOverrides
        Comment              = $txtComment.Text.Trim()
        FileShareRoot        = $fsPathValue
        DownloadRoot         = $dlRootValue
        M365Channel          = $script:Prefs.M365Channel
        M365DeployMode       = $script:Prefs.M365DeployMode
        EstimatedRuntimeMins = $script:Prefs.EstimatedRuntimeMins
        MaximumRuntimeMins   = $script:Prefs.MaximumRuntimeMins
        AdminUiFound         = $script:Prefs.DetectedTools.ConfigMgrConsole.Found
        LogFolder            = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath         = Get-SevenZipPathForContext
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
    Save-WindowState -Window $window

    # Dispose the async pipeline runspace so the bg thread doesn't keep
    # the process alive after the window closes. Per
    # reference_wpf_async_progress_overlay.md.
    if ($script:BgTimer) {
        try { $script:BgTimer.Stop() } catch { $null = $_ }
        $script:BgTimer = $null
    }
    if ($script:BgPS)    {
        try { [void]$script:BgPS.Stop() } catch { $null = $_ }
        try { $script:BgPS.Dispose() }   catch { $null = $_ }
        $script:BgPS = $null
    }
    if ($script:BgRunspace) {
        try { $script:BgRunspace.Close() }   catch { $null = $_ }
        try { $script:BgRunspace.Dispose() } catch { $null = $_ }
        $script:BgRunspace = $null
    }
    $script:BgHandle = $null
    $script:BgState  = $null
})

# Defaults (overridden by Restore-WindowState if saved state exists)
$script:SavedDarkTheme = $true
$script:SavedDebugCols = $false

# Restore previous window position + saved preferences
Restore-WindowState -Window $window

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
