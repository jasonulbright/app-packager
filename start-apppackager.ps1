<#
.SYNOPSIS
    Main window of AppPackager, which packages applications for ConfigMgr and Intune.

.DESCRIPTION
    Front-end for the application packager scripts in the Packagers folder.
    Sidebar layout with a dark and light theme toggle.

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
      - 7-Zip (required by Adobe Reader)
      - Local administrator (required by some packagers)

    ScriptName : start-apppackager.ps1
    Purpose    : Main window of AppPackager
    Owner      : CM Engineering
    Version    : 2026.09.25.0093
    Updated    : 2026-09-09
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
# Its SuiteCommon import repairs the process PSModulePath (PowerShell 7 module
# roots inherited from the launching shell) before any runspace or child
# powershell.exe starts.
Import-Module (Join-Path $PSScriptRoot 'Packagers\AppPackagerCommon.psm1') -Force -DisableNameChecking -ErrorAction SilentlyContinue

# The workbench definition model and the signing service. Common imports
# them for the packager children; the GUI imports them here so the editor
# and the Options signing panel work before any child process starts. A
# second import of an already-loaded module is a no-op.
foreach ($workbenchModule in @('AppPackagerWorkbench.psm1', 'AppPackagerSigning.psm1')) {
    $modulePath = Join-Path $PSScriptRoot ('Packagers\' + $workbenchModule)
    if (Test-Path -LiteralPath $modulePath) {
        Import-Module $modulePath -Force -Global -DisableNameChecking -ErrorAction SilentlyContinue
    }
}

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

function Resolve-FirstRunCompleted {
    # A preferences file written before the flag existed belongs to a
    # configured user: absence of the flag in an existing file counts as
    # completed, so only a genuinely missing file triggers the wizard.
    param(
        $StoredValue,
        [bool]$PreferencesFileExisted
    )

    if (-not $PreferencesFileExisted) { return $false }
    if ($null -eq $StoredValue) { return $true }
    try { return [bool]$StoredValue } catch { return $true }
}

function Test-FirstRunWizardNeeded {
    param([Parameter(Mandatory)][pscustomobject]$Prefs)

    $completed = $false
    try { $completed = [bool]$Prefs.FirstRunCompleted } catch { $completed = $false }
    return (-not $completed)
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
        DBeaverInstallOptions = [pscustomobject]@{
            InstallScope = "System"
            DisableAI    = $false
        }
        BeyondCompareKeyFile = ""
        LocalSourceFolders   = [pscustomobject]@{}
        HiddenApplications   = @()
        FirstRunCompleted    = $false
        IncludeVersionInTitle = $false
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
            CreateIntuneWin       = $false
            # ConfigMgr = today's flow; MECMAndIntune = ConfigMgr app + Graph publish;
            # IntuneOnly = stage + .intunewin + Graph publish, no site touch.
            DeploymentTarget      = 'MECM'
            PublishToIntune       = $false
            TenantId              = ''
            ClientId              = ''
            # DPAPI-protected (ConvertFrom-SecureString); never plaintext.
            ClientSecretProtected = ''
        }
        DeploymentConditions = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
        CommandOverrides = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
        WorkbenchDataRoot = ''
        # Signing creation and signature enforcement are separate concerns;
        # every switch defaults off so an existing installation keeps its
        # current behavior until an operator turns one on.
        ScriptSigning = [pscustomobject]@{
            SignDetection         = $false
            SignRequirements      = $false
            SignDeployment        = $false
            RequireDetection      = $false
            RequireRequirements   = $false
            RequireDeployment     = $false
            CertificateThumbprint = ''
            StoreLocation         = 'CurrentUser'
            TimestampServer       = ''
            TimestampRequired     = $false
            HashAlgorithm         = 'SHA256'
        }
    }

    $path = Get-PreferencesPath
    if (-not (Test-Path -LiteralPath $path)) { return $defaults }
    $defaults.FirstRunCompleted = Resolve-FirstRunCompleted -StoredValue $null -PreferencesFileExisted $true

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

        if ($null -ne $data.DBeaverInstallOptions) {
            $dbv = $data.DBeaverInstallOptions
            if ([string]$dbv.InstallScope -in @('System','User')) { $defaults.DBeaverInstallOptions.InstallScope = [string]$dbv.InstallScope }
            if ($null -ne $dbv.DisableAI) {
                try { $defaults.DBeaverInstallOptions.DisableAI = [bool]$dbv.DisableAI } catch { }
            }
        }

        if ($null -ne $data.BeyondCompareKeyFile)  { $defaults.BeyondCompareKeyFile = [string]$data.BeyondCompareKeyFile }
        if ($null -ne $data.LocalSourceFolders) {
            $sourceProps = [ordered]@{}
            foreach ($prop in $data.LocalSourceFolders.PSObject.Properties) {
                if ($prop.Name -notmatch '^package-') { continue }
                if ([string]::IsNullOrWhiteSpace([string]$prop.Value)) { continue }
                $sourceProps[$prop.Name] = [string]$prop.Value
            }
            $defaults.LocalSourceFolders = [pscustomobject]$sourceProps
        }
        if ($null -ne $data.HiddenApplications)    { $defaults.HiddenApplications  = @($data.HiddenApplications) }
        $defaults.FirstRunCompleted = Resolve-FirstRunCompleted -StoredValue $data.FirstRunCompleted -PreferencesFileExisted $true
        if ($null -ne $data.IncludeVersionInTitle) {
            try { $defaults.IncludeVersionInTitle = [bool]$data.IncludeVersionInTitle } catch { }
        }

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

        # Intune: .intunewin production and Graph publishing during Package.
        if ($null -ne $data.Intune) {
            if ($null -ne $data.Intune.CreateIntuneWin) { try { $defaults.Intune.CreateIntuneWin = [bool]$data.Intune.CreateIntuneWin } catch { } }
            if ($null -ne $data.Intune.PublishToIntune) { try { $defaults.Intune.PublishToIntune = [bool]$data.Intune.PublishToIntune } catch { } }
            if ([string]$data.Intune.DeploymentTarget -in @('MECM', 'MECMAndIntune', 'IntuneOnly')) {
                $defaults.Intune.DeploymentTarget = [string]$data.Intune.DeploymentTarget
            }
            elseif ($defaults.Intune.PublishToIntune) {
                # Pre-1.4.0.14 prefs expressed publishing as a bare toggle.
                $defaults.Intune.DeploymentTarget = 'MECMAndIntune'
            }
            if ($null -ne $data.Intune.TenantId)              { $defaults.Intune.TenantId              = [string]$data.Intune.TenantId }
            if ($null -ne $data.Intune.ClientId)              { $defaults.Intune.ClientId              = [string]$data.Intune.ClientId }
            if ($null -ne $data.Intune.ClientSecretProtected) { $defaults.Intune.ClientSecretProtected = [string]$data.Intune.ClientSecretProtected }
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
                $titleMode = ''
                if ([string]$entry.TitleMode -in @('IncludeVersion', 'NoVersion')) { $titleMode = [string]$entry.TitleMode }
                $installMode = ''
                if ([string]$entry.InstallMode -in @('CurrentUser', 'AllUsers')) { $installMode = [string]$entry.InstallMode }
                if ($arch -eq 'Any' -and $network -eq 'Any' -and $langs.Count -eq 0 -and $split -eq 'None' -and -not $titleMode -and -not $installMode) { continue }
                $condProps[$prop.Name] = [pscustomobject]@{
                    Architecture = $arch
                    Languages    = $langs
                    Network      = $network
                    Split        = $split
                    TitleMode    = $titleMode
                    InstallMode  = $installMode
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

        if ($null -ne $data.WorkbenchDataRoot) { $defaults.WorkbenchDataRoot = [string]$data.WorkbenchDataRoot }

        # ScriptSigning: an unknown or missing key keeps the all-off default,
        # so a preferences file from an older build never enables signing.
        if ($null -ne $data.ScriptSigning) {
            $sign = $data.ScriptSigning
            foreach ($flag in @('SignDetection','SignRequirements','SignDeployment','RequireDetection','RequireRequirements','RequireDeployment','TimestampRequired')) {
                if ($null -ne $sign.$flag) { try { $defaults.ScriptSigning.$flag = [bool]$sign.$flag } catch { } }
            }
            if ($null -ne $sign.CertificateThumbprint) {
                $thumb = ([string]$sign.CertificateThumbprint).Trim()
                if ($thumb -match '^[0-9A-Fa-f]{40}$') { $defaults.ScriptSigning.CertificateThumbprint = $thumb.ToUpperInvariant() }
            }
            if ([string]$sign.StoreLocation -in @('CurrentUser','LocalMachine')) { $defaults.ScriptSigning.StoreLocation = [string]$sign.StoreLocation }
            if ($null -ne $sign.TimestampServer) { $defaults.ScriptSigning.TimestampServer = ([string]$sign.TimestampServer).Trim() }
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
        $pkgPrefs["DBeaverInstallOptions"] = $Prefs.DBeaverInstallOptions
        $pkgPrefs["BeyondCompareKeyFile"] = [string]$Prefs.BeyondCompareKeyFile
        $pkgPrefs["LocalSourceFolders"] = $Prefs.LocalSourceFolders
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
    # the Adobe enterprise installer. Supporting
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

function Get-GitHubApiAuthStatus {
    # Reports how the 90 GitHub-backed packagers will authenticate against
    # api.github.com, in the same order Get-GitHubApiCurlArgs resolves the
    # token: GITHUB_TOKEN, GH_TOKEN, then the GitHub CLI's stored login.
    # Anonymous calls are limited to 60 per hour per address, which a version
    # sweep over the catalog exceeds. The rate-limit probe is one request with
    # a short timeout; when it cannot complete the source is still reported.
    $source = ''
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)) { $source = 'GITHUB_TOKEN environment variable' }
    elseif (-not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)) { $source = 'GH_TOKEN environment variable' }
    else {
        $gh = Get-Command -Name 'gh.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($gh) {
            $args = @()
            try { $args = @(Get-GitHubApiCurlArgs) } catch { $args = @() }
            if ($args.Count -gt 0) {
                $source = 'GitHub CLI login (gh.exe)'
                try {
                    $status = (& $gh.Source auth status 2>&1 | Out-String)
                    if ($status -match 'account\s+(\S+)') { $source = 'GitHub CLI login (gh.exe), account ' + $Matches[1] }
                }
                catch { }
            }
        }
    }

    $limit = ''
    $remaining = ''
    try {
        $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
        if ($curl) {
            $headers = (& $curl.Source -sI --max-time 6 -A 'app-packager' @(Get-GitHubApiCurlArgs) 'https://api.github.com/rate_limit' 2>$null) -join "`n"
            if ($headers -match '(?im)^x-ratelimit-limit:\s*(\d+)')     { $limit = $Matches[1] }
            if ($headers -match '(?im)^x-ratelimit-remaining:\s*(\d+)') { $remaining = $Matches[1] }
        }
    }
    catch { }

    return [pscustomobject]@{
        Authenticated = -not [string]::IsNullOrWhiteSpace($source)
        Source        = $source
        Limit         = $limit
        Remaining     = $remaining
    }
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

function Set-LauncherConnectionDefault {
    # The suite launcher hands its site code and provider to each tool it
    # starts. A saved value or a parameter wins; the launcher value fills an
    # empty one, and it is saved only when the operator saves the settings.
    param([Parameter(Mandatory)]$Prefs)

    if ([string]::IsNullOrWhiteSpace([string]$Prefs.SiteCode) -and -not [string]::IsNullOrWhiteSpace($env:SUITE_CM_SITECODE)) {
        $Prefs.SiteCode = $env:SUITE_CM_SITECODE.Trim()
    }
    if ([string]::IsNullOrWhiteSpace([string]$Prefs.ProviderMachineName) -and -not [string]::IsNullOrWhiteSpace($env:SUITE_CM_PROVIDER)) {
        $Prefs.ProviderMachineName = $env:SUITE_CM_PROVIDER.Trim()
    }
}

Set-LauncherConnectionDefault -Prefs $script:Prefs

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
        SupportsInstallModes = @()
        LocalSource       = $false
        LocalSourceRequired = $false
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
        if ($meta.SupportsInstallModes.Count -eq 0 -and $l -match '^\s*(?:#\s*)?SupportsInstallModes\s*:\s*(.+?)\s*$') {
            $meta.SupportsInstallModes = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -in @('CurrentUser', 'AllUsers') })
            continue
        }
        if (-not $meta.LocalSource -and $l -match '^\s*(?:#\s*)?LocalSource\s*:\s*(Required|Optional)\s*$') {
            $meta.LocalSource = $true
            $meta.LocalSourceRequired = ($Matches[1] -eq 'Required')
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
        SupportsInstallModes = @($meta.SupportsInstallModes)
        LocalSource       = [bool]$meta.LocalSource
        LocalSourceRequired = [bool]$meta.LocalSourceRequired
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
                SupportsVariants  = @($m.SupportsVariants)
                SupportsInstallModes = @($m.SupportsInstallModes)
                LocalSource       = [bool]$m.LocalSource
                LocalSourceRequired = [bool]$m.LocalSourceRequired
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
                SupportsVariants  = @()
                SupportsInstallModes = @()
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
        # Drop-generated packagers name the subfolder through $AppFolder or
        # $VendorFolder; an unresolved subfolder widens manifest searches to
        # the whole download root.
        $subfolderVariable = $null
        foreach ($line in Get-Content -LiteralPath $ScriptPath -ErrorAction Stop) {
            if (-not $info.DownloadSubfolder -and -not $subfolderVariable -and
                $line -match '\$BaseDownloadRoot\s*=\s*Join-Path\s+\$DownloadRoot\s+(?:"([^"]+)"|''([^'']+)''|\$(AppFolder|VendorFolder)\b)') {
                if ($matches[1]) { $info.DownloadSubfolder = $matches[1] }
                elseif ($matches[2]) { $info.DownloadSubfolder = $matches[2] }
                else { $subfolderVariable = $matches[3] }
            }
            if (-not $info.VendorFolder -and $line -match '^\s*\$VendorFolder\s*=\s*(?:"([^"]+)"|''([^'']+)'')') {
                $info.VendorFolder = $(if ($matches[1]) { $matches[1] } else { $matches[2] })
            }
            if (-not $info.AppFolder -and $line -match '^\s*\$AppFolder\s*=\s*(?:"([^"]+)"|''([^'']+)'')') {
                $info.AppFolder = $(if ($matches[1]) { $matches[1] } else { $matches[2] })
            }
            if (($info.DownloadSubfolder -or $subfolderVariable) -and $info.VendorFolder -and $info.AppFolder) { break }
        }
        if (-not $info.DownloadSubfolder -and $subfolderVariable) {
            $info.DownloadSubfolder = $info[$subfolderVariable]
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

function Get-ExistingConflictFromOutput {
    # Parses packager stdout for the marker New-MECMApplicationFromManifest
    # writes when it skips an application that already exists at the same
    # version. Returns $null when the run reported no such conflict.
    param([AllowEmptyString()][string]$Output = '')

    if ([string]::IsNullOrWhiteSpace($Output)) { return $null }
    $match = [regex]::Match(
        $Output,
        "\[APP_PACKAGER_CONFLICT\]\s+existing-same-version\s+app='([^']*)'\s+version='([^']*)'"
    )
    if (-not $match.Success) { return $null }
    return [pscustomobject]@{
        AppName = $match.Groups[1].Value
        Version = $match.Groups[2].Value
    }
}

function Invoke-PackagerPackageWithConflictPrompt {
    # Probe the exact manifest selected by the packager before share/site
    # writes, then ask about any exact-name collision, regardless of version.
    #
    # The prompt itself belongs to the UI thread: this runs in the background
    # runspace, so it parks a request on the synchronized state and polls for
    # the answer, the same handshake the Pause button uses.
    param(
        [Parameter(Mandatory)][hashtable]$State,
        [Parameter(Mandatory)][string]$AppLabel,
        [Parameter(Mandatory)][hashtable]$PackageArgs
    )

    if ([string]$PackageArgs.DeploymentTarget -eq 'IntuneOnly') {
        return (Invoke-PackagerPackage @PackageArgs)
    }
    $probeArgs = @{} + $PackageArgs
    $probeArgs['Preflight'] = $true
    $res = Invoke-PackagerPackage @probeArgs
    if ($res.ExitCode -ne 0) { return $res }
    $identityMatch = [regex]::Match([string]$res.StdOut, '(?m)^\[APP_PACKAGER_PREFLIGHT\] (.+)$')
    if (-not $identityMatch.Success) { throw 'Package preflight did not return an application identity; refusing to package without a conflict check.' }
    $identity = $identityMatch.Groups[1].Value | ConvertFrom-Json -ErrorAction Stop
    $existing = Get-MecmCurrentVersionByCMName -SiteCode $PackageArgs.SiteCode -ProviderMachineName $PackageArgs.ProviderMachineName -CMName $identity.AppName -ExactMatch
    if ([bool]$State.CancelRequested) {
        $State.Canceled = $true
        $res | Add-Member -NotePropertyName PackageOutcome -NotePropertyValue 'Canceled' -Force
        return $res
    }
    if (-not $existing.Found) {
        $createArgs = @{} + $PackageArgs
        $createArgs['OnExisting'] = 'Fail'
        return (Invoke-PackagerPackage @createArgs)
    }
    $conflict = [pscustomobject]@{ AppName = $identity.AppName; Version = $existing.SoftwareVersion; IncomingVersion = $identity.Version }
    $res | Add-Member -NotePropertyName PackageOutcome -NotePropertyValue 'Skipped' -Force

    # An apply-to-all answer lives for this run only and is never persisted.
    $decision = [string]$State.ConflictDecisionForAll
    if ([string]::IsNullOrWhiteSpace($decision)) {
        $State.ConflictResponse = $null
        $State.ConflictRequest = [pscustomobject]@{
            AppLabel = $AppLabel
            AppName  = $conflict.AppName
            Version  = $conflict.Version
            IncomingVersion = $conflict.IncomingVersion
        }
        while ($null -eq $State.ConflictResponse -and -not [bool]$State.CancelRequested) {
            $State.Step = ('Waiting for overwrite decision: {0}' -f $AppLabel)
            Start-Sleep -Milliseconds 150
        }
        $answer = $State.ConflictResponse
        $State.ConflictRequest = $null
        $State.ConflictResponse = $null
        if ($null -eq $answer) { $res.PackageOutcome = 'Canceled'; $State.Canceled = $true; return $res }
        $decision = [string]$answer.Choice
        if ([bool]$answer.ApplyToAll -and $decision -ne 'Cancel') {
            $State.ConflictDecisionForAll = $decision
        }
    }

    switch ($decision) {
        'Overwrite' {
            [void]$State.LogQueue.Enqueue(('Existing {0} v{1}: replacing deployment types.' -f $conflict.AppName, $conflict.Version))
            $retryArgs = @{} + $PackageArgs
            $retryArgs['OnExisting'] = 'Overwrite'
            return (Invoke-PackagerPackage @retryArgs)
        }
        'Cancel' {
            $State.CancelRequested = $true
            $State.Canceled = $true
            $res.PackageOutcome = 'Canceled'
            [void]$State.LogQueue.Enqueue(('Existing {0} v{1}: run canceled at the operator''s request.' -f $conflict.AppName, $conflict.Version))
            return $res
        }
        default {
            [void]$State.LogQueue.Enqueue(('Existing {0} v{1}: left unchanged.' -f $conflict.AppName, $conflict.Version))
            return $res
        }
    }
}

function Invoke-PackagerIntuneWinPostStep {
    # Produces <AppFolder>-<Version>.intunewin from the staged content and
    # copies it beside the network content version folder. The artifact
    # lands in the parent of both version folders, never inside them:
    # stage hash verification fails on any file added to verified content.
    # Failures never fail the package run - the ConfigMgr application already
    # exists when this executes - so the returned note carries Ok/Message
    # for the caller to surface.
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$FileServerPath,
        [string]$DownloadRoot = $null,
        [string]$ToolPath = '',
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested',
        [switch]$SkipNetworkCopy
    )

    $note = [pscustomobject]@{
        Ok          = $false
        Message     = ''
        LocalPath   = ''
        NetworkPath = ''
    }

    try {
        if ([string]::IsNullOrWhiteSpace($ToolPath) -or -not (Test-Path -LiteralPath $ToolPath)) {
            $note.Message = 'IntuneWinAppUtil.exe not available; skipped. Configure it in ConfigMgr Preferences.'
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

        $networkContentPath = if ($SkipNetworkCopy) { '' } else { Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Network content path' }
        if (-not $SkipNetworkCopy -and [string]::IsNullOrWhiteSpace($networkContentPath) -and $info.VendorFolder -and $info.AppFolder) {
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

        $note | Add-Member -NotePropertyName ManifestPath -NotePropertyValue $manifestPath -Force
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
        # [version] rejects a single number such as NetBeans "31".
        $va = [version](($A -replace '[+-].*$', '') -replace '^(\d+)$', '$1.0')
        $vb = [version](($B -replace '[+-].*$', '') -replace '^(\d+)$', '$1.0')

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
        if ($version -notmatch '^\d+(\.\d+){0,3}([+-]\d+)?$') {
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
        [Parameter(Mandatory)][string]$CMName,
        [switch]$ExactMatch
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
            throw ("Failed to connect to CM site PSDrive '{0}:'. Open the ConfigMgr console once on this machine, or set Provider Machine in Options > ConfigMgr Preferences (the ProviderMachineName value from the AdminUI connect script)." -f $SiteCode)
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
        if ($ExactMatch) {
            $apps = @(Get-CMApplication -Name $CMName -DisableWildcardHandling -ErrorAction Stop |
                Where-Object { $_.LocalizedDisplayName -eq $CMName -or $_.Name -eq $CMName })
            if ($apps.Count -gt 1) { throw "Multiple applications have the exact title '$CMName'; resolve duplicates before packaging." }
            if ($apps.Count -eq 0) { return [pscustomobject]@{ Found = $false; SoftwareVersion = ''; MatchCount = 0 } }
            return [pscustomobject]@{ Found = $true; DisplayName = $CMName; SoftwareVersion = [string]$apps[0].SoftwareVersion; MatchCount = 1 }
        }
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
                try { $null = [version]([string]$a.SoftwareVersion -replace '^(\d+)$', '$1.0'); $parsable += $a }
                catch { $nonParsable += $a }
            }
            if ($parsable.Count -gt 0) {
                $chosen = $parsable | Sort-Object { [version]([string]$_.SoftwareVersion -replace '^(\d+)$', '$1.0') } -Descending | Select-Object -First 1
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
        [string]$VariantsJson = '',
        [string]$InstallMode = '',
        [string]$RunSnapshotPath = '',
        [string]$SigningJson = '',
        [string]$WorkbenchDataRoot = '',
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
    # Variant splits and the install mode shape the staged content, so
    # they travel with the Stage child too.
    Set-PackagerEnvironment -StartInfo $psi -SevenZipPath $SevenZipPath -VariantsJson $VariantsJson -InstallMode $InstallMode -RunSnapshotPath $RunSnapshotPath -SigningJson $SigningJson -WorkbenchDataRoot $WorkbenchDataRoot -DownloadRoot $DownloadRoot

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
        [hashtable]$IntunePublishConfig = $null,
        [ValidateSet('MECM', 'MECMAndIntune', 'IntuneOnly')][string]$DeploymentTarget = 'MECM',
        [ValidateSet('Nested','Flat')][string]$ContentLayout = 'Nested',
        [string]$RequirementsJson = '',
        [string]$VariantsJson = '',
        [string]$CommandsJson = '',
        [ValidateSet('', 'Skip', 'Overwrite', 'Fail')][string]$OnExisting = '',
        [ValidateSet('', 'Default', 'IncludeVersion', 'NoVersion')][string]$TitleMode = '',
        [switch]$Preflight,
        [string]$InstallMode = '',
        [string]$RunSnapshotPath = '',
        [string]$SigningJson = '',
        [string]$WorkbenchDataRoot = '',
        [string]$BuildId = '',
        [System.Windows.Controls.TextBox]$LogTextBox = $null
    )

    # Package consumes the build the operator chose. The child resolves its
    # own content from staged-version.txt, so a selection that is not the
    # newest stage in the tree the child will read is refused here rather
    # than packaged as something else.
    $selectedManifestPath = ''
    if (-not [string]::IsNullOrWhiteSpace($BuildId)) {
        $selected = Assert-WorkbenchBuildSelection -BuildId $BuildId -DownloadRoot $DownloadRoot -PackagerPath $PackagerPath
        if ($selected) { $selectedManifestPath = [string]$selected.Path }
    }

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
    # Intune-only runs stop at Stage: no site connection, no share copy,
    # no ConfigMgr application. The .intunewin build and Graph publish below
    # work entirely from the local staged content.
    $phaseSwitch = if ($DeploymentTarget -eq 'IntuneOnly') { '-StageOnly' } else { '-PackageOnly' }
    $argsBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PackagerPath, $phaseSwitch, '-SiteCode', $SiteCode, '-Comment', $Comment, '-LogPath', $structuredLog)
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
    Set-PackagerEnvironment -StartInfo $psi -SevenZipPath $SevenZipPath -ProviderMachineName $ProviderMachineName -RequirementsJson $RequirementsJson -VariantsJson $VariantsJson -CommandsJson $CommandsJson -OnExisting $OnExisting -InstallMode $InstallMode -RunSnapshotPath $RunSnapshotPath -SigningJson $SigningJson -WorkbenchDataRoot $WorkbenchDataRoot -DownloadRoot $DownloadRoot -BuildId $BuildId -StageManifestPath $selectedManifestPath
    $psi.EnvironmentVariables['APP_PACKAGER_TITLE_MODE'] = $TitleMode
    $psi.EnvironmentVariables['APP_PACKAGER_PACKAGE_PREFLIGHT'] = $(if ($Preflight) { '1' } else { '0' })

    $result = Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
    if ($Preflight) { return $result }
    if ($DeploymentTarget -ne 'IntuneOnly') {
        # Network-copy verification; an Intune-only run never copies to the
        # share (stage integrity is verified when the manifest is written).
        Assert-PackagerPackageIntegrity -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ContentLayout $ContentLayout
    }

    # Optional post-step: produce a .intunewin beside the network content.
    # Runs only after integrity passes; failures ride on the result for the
    # caller to surface, never thrown - the ConfigMgr application already exists
    # by this point.
    $intuneOnly = ($DeploymentTarget -eq 'IntuneOnly')
    if (($CreateIntuneWin -or $intuneOnly) -and $result.ExitCode -eq 0) {
        $intuneNote = Invoke-PackagerIntuneWinPostStep -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ToolPath $IntuneWinToolPath -ContentLayout $ContentLayout -SkipNetworkCopy:$intuneOnly
        if ($intuneOnly -and -not $IntunePublishConfig -and $intuneNote.Ok) {
            $result | Add-Member -NotePropertyName IntunePublish -NotePropertyValue ([pscustomobject]@{ Ok = $false; Message = 'Intune credentials not configured; set Tenant ID, Client ID, and Client Secret in ConfigMgr Preferences.' }) -Force
        }
        if ($IntunePublishConfig -and $intuneNote.Ok) {
            $pubNote = [pscustomobject]@{ Ok = $false; Message = '' }
            try {
                $manifest = Read-StageManifest -Path $intuneNote.ManifestPath
                if ($manifest.PSObject.Properties['DeploymentTypes'] -and $manifest.DeploymentTypes) {
                    $pubNote.Message = 'variant-split app; Intune publish skipped (single-payload apps only for now).'
                }
                else {
                    # The staged icon lives in the content version folder next
                    # to the manifest, not beside the .intunewin.
                    $iconArg = @{}
                    if ($manifest.PSObject.Properties['Icon'] -and $manifest.Icon) {
                        $iconFile = Join-Path (Split-Path -Parent $intuneNote.ManifestPath) ([string]$manifest.Icon)
                        if (Test-Path -LiteralPath $iconFile) { $iconArg.IconPath = $iconFile }
                    }
                    $pubId = Publish-IntuneWin32App -TenantId $IntunePublishConfig.TenantId -ClientId $IntunePublishConfig.ClientId -ClientSecret $IntunePublishConfig.ClientSecret -IntuneWinPath $intuneNote.LocalPath -Manifest $manifest -Description $Comment @iconArg
                    $pubNote.Ok = $true
                    $pubNote.Message = ('published (app id {0})' -f $pubId)
                }
            }
            catch {
                $pubNote.Message = ('publish failed: {0}' -f $_.Exception.Message)
            }
            $result | Add-Member -NotePropertyName IntunePublish -NotePropertyValue $pubNote -Force
        }
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
        [string]$CommandsJson,
        [string]$OnExisting,
        [string]$InstallMode,
        [string]$RunSnapshotPath,
        [string]$SigningJson,
        [string]$WorkbenchDataRoot,
        [string]$DownloadRoot,
        [string]$BuildId,
        [string]$StageManifestPath
    )
    # The selected build travels to the child so a packager that learns to
    # honor it resolves the same content this caller verified.
    if (-not [string]::IsNullOrWhiteSpace($BuildId)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_BUILD_ID'] = [string]$BuildId
    }
    if (-not [string]::IsNullOrWhiteSpace($StageManifestPath)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_STAGE_MANIFEST'] = [string]$StageManifestPath
    }
    # The shared download cache resolves from the child's download root, so a
    # per-profile subroot must still reach the common cache folder.
    if (-not [string]::IsNullOrWhiteSpace($DownloadRoot)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_DOWNLOAD_ROOT'] = [string]$DownloadRoot
    }
    # The child resolves the same data root as the caller, so a snapshot
    # written here is readable there.
    if (-not [string]::IsNullOrWhiteSpace($RunSnapshotPath)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_RUN_SNAPSHOT'] = [string]$RunSnapshotPath
    }
    if (-not [string]::IsNullOrWhiteSpace($WorkbenchDataRoot)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_WORKBENCH_ROOT'] = [string]$WorkbenchDataRoot
    }
    # Signing travels even without a snapshot, so the legacy path signs too.
    if (-not [string]::IsNullOrWhiteSpace($SigningJson)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_SIGNING'] = [string]$SigningJson
    }
    if (-not [string]::IsNullOrWhiteSpace($OnExisting)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_ON_EXISTING'] = [string]$OnExisting
    }
    if (-not [string]::IsNullOrWhiteSpace($InstallMode)) {
        $StartInfo.EnvironmentVariables['APP_PACKAGER_INSTALL_MODE'] = [string]$InstallMode
    }
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

function ConvertTo-InstallModeValue {
    # Maps one Deployment Conditions prefs entry onto the
    # APP_PACKAGER_INSTALL_MODE value a SupportsInstallModes packager
    # consumes via Get-RequestedInstallMode. Returns '' for the default.
    param($Entry)

    if (-not $Entry -or -not $Entry.PSObject.Properties['InstallMode']) { return '' }
    $mode = [string]$Entry.InstallMode
    if ($mode -notin @('CurrentUser', 'AllUsers')) { return '' }
    return $mode
}

function Get-InstallModesMapForContext {
    # Prebuilt on the UI thread, same reason as Get-RequirementsMapForContext.
    $map = @{}
    try {
        $apps = $script:Prefs.DeploymentConditions.Apps
        if ($apps) {
            foreach ($prop in $apps.PSObject.Properties) {
                $mode = ConvertTo-InstallModeValue -Entry $prop.Value
                if ($mode) { $map[$prop.Name] = $mode }
            }
        }
    } catch { }
    return $map
}

function Get-TitleModesMapForContext {
    $map = @{}
    $apps = $script:Prefs.DeploymentConditions.Apps
    if ($apps) {
        foreach ($prop in $apps.PSObject.Properties) {
            $mode = [string]$prop.Value.TitleMode
            if ($mode -in @('IncludeVersion', 'NoVersion')) { $map[$prop.Name] = $mode }
        }
    }
    return $map
}

function Get-DefaultTitleModeForContext {
    if ($script:Prefs -and [bool]$script:Prefs.IncludeVersionInTitle) { return 'IncludeVersion' }
    return ''
}

function Confirm-LocalSourceFolders {
    # Runs on the UI thread before a Stage: a hybrid packager whose installer
    # folder is unset or missing cannot prompt from the background runspace.
    param([array]$Rows)

    $kept = New-Object System.Collections.Generic.List[object]
    $changed = $false
    foreach ($row in @($Rows)) {
        $meta = $null
        try { $meta = Get-PackagerMetadata -Path ([string]$row.FullPath) } catch { }
        if (-not $meta -or -not $meta.LocalSourceRequired) { $kept.Add($row); continue }

        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script)
        $folder = ''
        if ($script:Prefs.LocalSourceFolders -and $script:Prefs.LocalSourceFolders.PSObject.Properties[$base]) {
            $folder = [string]$script:Prefs.LocalSourceFolders.$base
        }
        if ($folder -and (Test-Path -LiteralPath $folder -PathType Container)) { $kept.Add($row); continue }

        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = ('Choose the folder that holds the {0} installer. Its download requires a sign-in, so it stages from a local copy.' -f [string]$row.Application)
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $props = [ordered]@{}
            if ($script:Prefs.LocalSourceFolders) {
                foreach ($p in $script:Prefs.LocalSourceFolders.PSObject.Properties) { $props[$p.Name] = $p.Value }
            }
            $props[$base] = $dlg.SelectedPath
            $script:Prefs.LocalSourceFolders = [pscustomobject]$props
            $changed = $true
            Add-LogLine -Message ('Installer source for {0}: {1}' -f [string]$row.Application, $dlg.SelectedPath)
            $kept.Add($row)
        }
        else {
            Add-LogLine -Message ('Skipped {0}: no installer source folder chosen.' -f [string]$row.Application)
        }
    }
    if ($changed) { Save-Preferences -Prefs $script:Prefs }
    return $kept.ToArray()
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

function Get-IntunePublishConfigForContext {
    # Decrypts the stored client secret (DPAPI, current Windows user)
    # just-in-time. Returns $null when publishing is off or incomplete,
    # so callers can pass the result straight through.
    param($Prefs = $script:Prefs)
    try {
        $i = $Prefs.Intune
        if (-not $i) { return $null }
        $target = [string]$i.DeploymentTarget
        if ($target -notin @('MECMAndIntune', 'IntuneOnly') -and -not [bool]$i.PublishToIntune) { return $null }
        if ([string]::IsNullOrWhiteSpace([string]$i.TenantId) -or
            [string]::IsNullOrWhiteSpace([string]$i.ClientId) -or
            [string]::IsNullOrWhiteSpace([string]$i.ClientSecretProtected)) { return $null }
        $sec = ConvertTo-SecureString -String ([string]$i.ClientSecretProtected) -ErrorAction Stop
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($sec)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
        return @{ TenantId = [string]$i.TenantId; ClientId = [string]$i.ClientId; ClientSecret = $plain }
    } catch { return $null }
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
# Version and update check
# =============================================================================
$script:UpdateRepo        = 'jasonulbright/app-packager'
$script:UpdateCheckHours  = 24
$script:UpdateUserAgent   = 'AppPackager-UpdateCheck'

function Get-AppVersion {
    # The header comment block is the single source of truth for the version;
    # parsing it keeps the value from drifting between header and code.
    param([string]$ScriptPath = $PSCommandPath)

    if ([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -LiteralPath $ScriptPath)) { return $null }

    foreach ($line in (Get-Content -LiteralPath $ScriptPath -TotalCount 80 -ErrorAction SilentlyContinue)) {
        $m = [regex]::Match($line, '^\s*Version\s*:\s*([0-9]+(?:\.[0-9]+)+)\s*$')
        if ($m.Success) { return $m.Groups[1].Value }
    }
    return $null
}

function ConvertFrom-ReleaseTag {
    param([AllowNull()][string]$Tag)

    if ([string]::IsNullOrWhiteSpace($Tag)) { return $null }
    $m = [regex]::Match($Tag.Trim(), '^v?([0-9]+(?:\.[0-9]+)+)$')
    if (-not $m.Success) { return $null }
    return $m.Groups[1].Value
}

function Test-UpdateAvailable {
    param(
        [AllowNull()][string]$CurrentVersion,
        [AllowNull()][string]$LatestVersion
    )

    # An unparseable version on either side means no claim can be made; the
    # indicator stays hidden rather than nagging on bad data.
    $current = $null
    $latest  = $null
    if (-not [version]::TryParse(($CurrentVersion -as [string]), [ref]$current)) { return $false }
    if (-not [version]::TryParse(($LatestVersion  -as [string]), [ref]$latest))  { return $false }
    return ($latest -gt $current)
}

function Test-UpdateCheckDue {
    param(
        $LastCheckUtc,
        [datetime]$NowUtc = (Get-Date).ToUniversalTime(),
        [int]$IntervalHours = 24
    )

    if ($null -eq $LastCheckUtc -or ($LastCheckUtc -is [string] -and [string]::IsNullOrWhiteSpace($LastCheckUtc))) { return $true }

    $last = [datetime]::MinValue
    if ($LastCheckUtc -is [datetime]) {
        $last = $LastCheckUtc
    }
    elseif (-not [datetime]::TryParse(
        ($LastCheckUtc -as [string]), [cultureinfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$last)) {
        return $true
    }

    # A timestamp in the future means a clock change or a hand-edited cache;
    # treat it as stale so the check is never suppressed indefinitely.
    if ($last -gt $NowUtc) { return $true }
    return (($NowUtc - $last).TotalHours -ge $IntervalHours)
}

function Get-UpdateCheckCachePath {
    Join-Path (Join-Path $env:LOCALAPPDATA 'AppPackager') 'update-check.json'
}

function Read-UpdateCheckCache {
    $path = Get-UpdateCheckCachePath
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        return (Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        return $null
    }
}

function Save-UpdateCheckCache {
    param(
        [AllowNull()][string]$LatestVersion,
        [AllowNull()][string]$ReleaseUrl
    )

    try {
        $path = Get-UpdateCheckCachePath
        $dir  = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [pscustomobject]@{
            LastCheckUtc  = (Get-Date).ToUniversalTime().ToString('o')
            LatestVersion = $LatestVersion
            ReleaseUrl    = $ReleaseUrl
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $path -Encoding UTF8
    } catch {
        # A cache that cannot be written only costs an extra query next launch.
    }
}

function Get-UpdateReleaseUrl {
    param([AllowNull()][string]$Version)

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return ("https://github.com/{0}/releases/latest" -f $script:UpdateRepo)
    }
    return ("https://github.com/{0}/releases/tag/v{1}" -f $script:UpdateRepo, $Version)
}

function Show-UpdateIndicator {
    param(
        [AllowNull()][string]$LatestVersion,
        [AllowNull()][string]$ReleaseUrl
    )

    if (-not $pnlUpdate) { return }

    if ([string]::IsNullOrWhiteSpace($LatestVersion)) {
        $pnlUpdate.Visibility = 'Collapsed'
        return
    }

    $script:UpdateLatestVersion = $LatestVersion
    $script:UpdateReleaseUrl    = if ([string]::IsNullOrWhiteSpace($ReleaseUrl)) { Get-UpdateReleaseUrl -Version $LatestVersion } else { $ReleaseUrl }
    $runUpdateText.Text         = ("Update available: v{0}" -f $LatestVersion)
    $lnkUpdateAvailable.ToolTip = ("Open the v{0} release page on GitHub" -f $LatestVersion)
    $pnlUpdate.Visibility       = 'Visible'
}

function Start-UpdateCheck {
    # Never let a failed or slow check touch launch: everything below is
    # best-effort and the caller keeps going regardless.
    try {
        $current = Get-AppVersion
        if ([string]::IsNullOrWhiteSpace($current)) { return }

        $cache = Read-UpdateCheckCache
        $lastCheck = $null
        if ($cache) { $lastCheck = $cache.LastCheckUtc }

        if (-not (Test-UpdateCheckDue -LastCheckUtc $lastCheck -IntervalHours $script:UpdateCheckHours)) {
            if ($cache -and (Test-UpdateAvailable -CurrentVersion $current -LatestVersion $cache.LatestVersion)) {
                Add-LogLine -Message ("Update available: v{0} (cached check)." -f $cache.LatestVersion)
                Show-UpdateIndicator -LatestVersion $cache.LatestVersion -ReleaseUrl $cache.ReleaseUrl
            }
            return
        }

        $script:UpdateCheckState = [hashtable]::Synchronized(@{
            Done   = $false
            Latest = $null
            Url    = $null
        })

        $script:UpdateCheckRunspace = [runspacefactory]::CreateRunspace()
        $script:UpdateCheckRunspace.ApartmentState = 'MTA'
        $script:UpdateCheckRunspace.Open()

        $script:UpdateCheckPS = [powershell]::Create()
        $script:UpdateCheckPS.Runspace = $script:UpdateCheckRunspace
        [void]$script:UpdateCheckPS.AddScript({
            param($State, $Repo, $UserAgent)
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                # curl.exe first: SSL-inspecting proxies that break the .NET
                # chain commonly pass curl, which the packagers already use.
                $release = $null
                $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
                if ($curl) {
                    $json = (& $curl.Source -L --fail --silent --max-time 20 -A $UserAgent "https://api.github.com/repos/$Repo/releases/latest") -join "`n"
                    if ($LASTEXITCODE -eq 0 -and $json) { $release = $json | ConvertFrom-Json }
                }
                if (-not $release) {
                    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" `
                        -Headers @{ 'User-Agent' = $UserAgent } -UseBasicParsing -TimeoutSec 20
                }
                $State.Latest = $release.tag_name
                $State.Url    = $release.html_url
            } catch {
            } finally {
                $State.Done = $true
            }
        }).AddArgument($script:UpdateCheckState).AddArgument($script:UpdateRepo).AddArgument($script:UpdateUserAgent)

        $script:UpdateCheckHandle = $script:UpdateCheckPS.BeginInvoke()

        $script:UpdateCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:UpdateCheckTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:UpdateCheckTimer.Add_Tick({
            if (-not $script:UpdateCheckState.Done) { return }
            $script:UpdateCheckTimer.Stop()
            try {
                $latest = ConvertFrom-ReleaseTag -Tag $script:UpdateCheckState.Latest
                if ($latest) {
                    Save-UpdateCheckCache -LatestVersion $latest -ReleaseUrl $script:UpdateCheckState.Url
                    if (Test-UpdateAvailable -CurrentVersion (Get-AppVersion) -LatestVersion $latest) {
                        Add-LogLine -Message ("Update available: v{0}" -f $latest)
                        Show-UpdateIndicator -LatestVersion $latest -ReleaseUrl $script:UpdateCheckState.Url
                    }
                }
            } catch {
                Add-LogLine -Message ("Update check failed: {0}" -f $_.Exception.Message)
            } finally {
                try { $script:UpdateCheckPS.EndInvoke($script:UpdateCheckHandle) } catch { }
                try { $script:UpdateCheckPS.Dispose() } catch { }
                try { $script:UpdateCheckRunspace.Close() } catch { }
                $script:UpdateCheckPS = $null
                $script:UpdateCheckRunspace = $null
            }
        })
        $script:UpdateCheckTimer.Start()
    } catch {
        try { Add-LogLine -Message ("Update check skipped: {0}" -f $_.Exception.Message) } catch { }
    }
}

function Invoke-SelfUpdate {
    param([Parameter(Mandatory)]$Owner)

    $latest = $script:UpdateLatestVersion
    if ([string]::IsNullOrWhiteSpace($latest)) { return }

    $answer = Show-ThemedMessage -Owner $Owner -Title 'Update AppPackager' `
        -Message ("AppPackager v{0} will be downloaded and installed over {1}.`n`nThe application closes and relaunches itself when the update finishes. Preferences, window state, and logs are preserved.`n`nContinue?" -f $latest, $PSScriptRoot) `
        -Buttons YesNo -Icon Question
    if ($answer -ne 'Yes') { return }

    try {
        $installer = Join-Path $PSScriptRoot 'install.ps1'
        if (-not (Test-Path -LiteralPath $installer)) {
            $installer = Join-Path ([IO.Path]::GetTempPath()) ('apinstall-' + [Guid]::NewGuid().ToString('N') + '.ps1')
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            # Release-asset host, not raw.githubusercontent.com - proxies that
            # block the raw host allow the asset host the packagers use.
            $installerUrl = ("https://github.com/{0}/releases/latest/download/install.ps1" -f $script:UpdateRepo)
            $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
            $fetched = $false
            if ($curl) {
                & $curl.Source -L --fail --silent -A $script:UpdateUserAgent -o $installer $installerUrl
                $fetched = ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $installer))
            }
            if (-not $fetched) {
                Invoke-WebRequest -Uri $installerUrl -OutFile $installer -UseBasicParsing -Headers @{ 'User-Agent' = $script:UpdateUserAgent }
            }
        }

        # The updater replaces the folder this process is running from, so it
        # has to outlive the process: a detached child waits on the PID, then
        # installs and relaunches. Wait-Process is bounded so a hung shutdown
        # cannot leave the child waiting forever.
        $relaunch = Join-Path $PSScriptRoot 'start-apppackager.ps1'
        $script = @"
Wait-Process -Id $PID -Timeout 120 -ErrorAction SilentlyContinue
& '$($installer.Replace("'","''"))' -InstallPath '$($PSScriptRoot.Replace("'","''"))'
Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','$($relaunch.Replace("'","''"))'
"@
        $bootstrap = Join-Path ([IO.Path]::GetTempPath()) ('apupdate-' + [Guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $bootstrap -Value $script -Encoding UTF8

        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $bootstrap

        Add-LogLine -Message ("Updating to v{0}; closing." -f $latest)
        $Owner.Close()
    } catch {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Update Failed' -Message $_.Exception.Message -Buttons OK -Icon Error)
    }
}

function Open-UpdateReleasePage {
    try {
        $url = $script:UpdateReleaseUrl
        if ([string]::IsNullOrWhiteSpace($url)) { $url = Get-UpdateReleaseUrl -Version $script:UpdateLatestVersion }
        Start-Process $url
    } catch {
        Add-LogLine -Message ("Could not open the release page: {0}" -f $_.Exception.Message)
    }
}

# =============================================================================
# Packager icon pack
# =============================================================================
$script:IconPackRepo      = 'jasonulbright/app-packager-icons'
$script:IconPackUserAgent = 'AppPackager-IconPack'
$script:IconPackAssetName = 'icon-pack.zip'
$script:IconPackSumsName  = 'checksums.txt'

function Get-IconPackRoot {
    param([string]$AppRoot = $PSScriptRoot)
    Join-Path (Join-Path $AppRoot 'Packagers') 'Icons'
}

function Get-IconPackManifestPath {
    param([string]$AppRoot = $PSScriptRoot)
    Join-Path (Get-IconPackRoot -AppRoot $AppRoot) 'manifest.json'
}

function Get-WorkbenchInheritedIconPath {
    # Mirrors what Stage publishes: the icon pack entry named for the
    # packager first, then the newest build's staged app-icon.
    param(
        [string]$ScriptPath,
        [string]$DownloadRoot
    )

    if ([string]::IsNullOrWhiteSpace($ScriptPath)) { return '' }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath) -replace '^package-', ''
    $pack = @(Get-ChildItem -LiteralPath (Get-IconPackRoot) -Filter "$base.*" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -eq $base -and $_.Extension -match '^\.(png|ico)$' } |
        Sort-Object Extension | Select-Object -First 1)
    if ($pack.Count -gt 0) { return $pack[0].FullName }

    if ((Get-PackagerIconSource -ScriptPath $ScriptPath) -eq 'None') { return '' }

    # Runs on the UI thread at every profile load: without a resolved
    # subfolder the manifest search would walk the entire download root.
    $manifest = $null
    if ((Get-PackagerFolderInfo -ScriptPath $ScriptPath).DownloadSubfolder) {
        $manifest = Find-NewestStageManifestForPackager -PackagerPath $ScriptPath -DownloadRoot $DownloadRoot
    }
    if ($manifest) {
        $staged = @(Get-ChildItem -LiteralPath (Split-Path -Parent $manifest) -Filter 'app-icon.*' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^\.(ico|png)$' } | Select-Object -First 1)
        if ($staged.Count -gt 0) { return $staged[0].FullName }
    }
    return ''
}

function Read-IconPackManifest {
    <#
    .SYNOPSIS
        Reads an installed icon pack manifest.

    .DESCRIPTION
        Returns PackVersion, MinAppVersion, and IconCount. A missing,
        unreadable, or malformed manifest returns $null so the caller can show
        the not-installed state instead of failing.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

    try {
        $data = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return $null
    }
    if ($null -eq $data) { return $null }

    $icons = @()
    if ($data.PSObject.Properties.Name -contains 'Icons' -and $data.Icons) { $icons = @($data.Icons) }

    return [pscustomobject]@{
        PackVersion   = [string]$data.PackVersion
        MinAppVersion = [string]$data.MinAppVersion
        IconCount     = $icons.Count
    }
}

function Test-IconPackAppVersion {
    <#
    .SYNOPSIS
        Reports whether the running app satisfies a pack's MinAppVersion.

    .DESCRIPTION
        An absent or unparseable version on either side is treated as
        satisfied: the pack is decoration and an unreadable bound must not
        block an install, only the warning that goes with it.
    #>
    param(
        [AllowNull()][string]$MinAppVersion,
        [AllowNull()][string]$CurrentVersion
    )

    $min     = $null
    $current = $null
    if (-not [version]::TryParse(($MinAppVersion  -as [string]), [ref]$min))     { return $true }
    if (-not [version]::TryParse(($CurrentVersion -as [string]), [ref]$current)) { return $true }
    return ($current -ge $min)
}

function Get-IconPackChecksum {
    # checksums.txt is sha256sum output: "<hash> *<filename>" or "<hash>  <filename>".
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ChecksumText,
        [Parameter(Mandatory)][string]$FileName
    )

    foreach ($line in ($ChecksumText -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $m = [regex]::Match($line.Trim(), '^([0-9a-fA-F]{64})\s+\*?(.+)$')
        if (-not $m.Success) { continue }
        if ($m.Groups[2].Value.Trim() -eq $FileName) { return $m.Groups[1].Value.ToLowerInvariant() }
    }
    return $null
}

function Test-IconPackChecksum {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [AllowNull()][string]$ExpectedSha256
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedSha256)) { return $false }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { return $false }
    $actual = (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash
    return ($actual -eq $ExpectedSha256.Trim().ToUpperInvariant())
}

function Select-IconPackAsset {
    <#
    .SYNOPSIS
        Picks the pack zip and checksum download URLs out of a release object.

    .DESCRIPTION
        Returns PackUrl, SumsUrl, and Tag. A release missing either asset
        yields $null for that URL, which the caller reports as an empty
        release rather than treating as an error.
    #>
    param([AllowNull()]$Release)

    if ($null -eq $Release) { return $null }

    $assets  = @()
    if ($Release.PSObject.Properties.Name -contains 'assets' -and $Release.assets) { $assets = @($Release.assets) }
    $packUrl = ($assets | Where-Object { $_.name -eq $script:IconPackAssetName } | Select-Object -First 1).browser_download_url
    $sumsUrl = ($assets | Where-Object { $_.name -eq $script:IconPackSumsName  } | Select-Object -First 1).browser_download_url

    return [pscustomobject]@{
        Tag     = [string]$Release.tag_name
        PackUrl = $packUrl
        SumsUrl = $sumsUrl
    }
}

function Get-IconPackStatusText {
    param([AllowNull()]$Manifest)

    if ($null -eq $Manifest) {
        return ([char]0x2717 + ' Not installed  -  download, or install from a file')
    }
    $version = if ([string]::IsNullOrWhiteSpace($Manifest.PackVersion)) { 'unknown' } else { $Manifest.PackVersion }
    if ($Manifest.IconCount -eq 0) {
        return ([char]0x2713 + " Installed  -  pack v{0}, no icons yet" -f $version)
    }
    $noun = if ($Manifest.IconCount -eq 1) { 'icon' } else { 'icons' }
    return ([char]0x2713 + " Installed  -  pack v{0}, {1} {2}" -f $version, $Manifest.IconCount, $noun)
}

function Install-IconPack {
    <#
    .SYNOPSIS
        Downloads, verifies, and extracts the latest icon pack release.

    .DESCRIPTION
        Fetches the release through the GitHub API, verifies icon-pack.zip
        against checksums.txt, and extracts it into Packagers\Icons. Files
        are downloaded to a scratch folder outside the app so nothing carries
        a zone identifier into the extracted tree. A MinAppVersion newer than
        the running app warns and still installs.

    .OUTPUTS
        [pscustomobject] Installed, Message, Manifest.
    #>
    param(
        [string]$AppRoot = $PSScriptRoot,
        [string]$Repo    = $script:IconPackRepo
    )

    $destination = Get-IconPackRoot -AppRoot $AppRoot
    $scratch     = Join-Path ([IO.Path]::GetTempPath()) ('apicons-' + [Guid]::NewGuid().ToString('N'))
    $result      = [pscustomobject]@{ Installed = $false; Message = ''; Manifest = $null }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $headers = @{ 'User-Agent' = $script:IconPackUserAgent }

        try {
            $release = Invoke-RestMethod -Uri ("https://api.github.com/repos/{0}/releases/latest" -f $Repo) `
                -Headers $headers -UseBasicParsing -TimeoutSec 20
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 403 -or $status -eq 429) {
                $result.Message = 'GitHub rate limit reached; try again later.'
            } elseif ($status -eq 404) {
                $result.Message = 'No icon pack release is published yet.'
            } else {
                $result.Message = ("Icon pack lookup failed: {0}" -f $_.Exception.Message)
            }
            return $result
        }

        $selection = Select-IconPackAsset -Release $release
        if ($null -eq $selection -or -not $selection.PackUrl -or -not $selection.SumsUrl) {
            $result.Message = 'The latest icon pack release carries no pack assets.'
            return $result
        }

        New-Item -ItemType Directory -Path $scratch -Force | Out-Null
        $zipPath  = Join-Path $scratch $script:IconPackAssetName
        $sumsPath = Join-Path $scratch $script:IconPackSumsName

        Invoke-WebRequest -Uri $selection.PackUrl -OutFile $zipPath  -UseBasicParsing -Headers $headers -TimeoutSec 120
        Invoke-WebRequest -Uri $selection.SumsUrl -OutFile $sumsPath -UseBasicParsing -Headers $headers -TimeoutSec 60

        $expected = Get-IconPackChecksum -ChecksumText (Get-Content -LiteralPath $sumsPath -Raw) -FileName $script:IconPackAssetName
        if (-not (Test-IconPackChecksum -FilePath $zipPath -ExpectedSha256 $expected)) {
            $result.Message = 'Icon pack checksum did not match; nothing was extracted.'
            return $result
        }

        return Complete-IconPackInstall -ZipPath $zipPath -AppRoot $AppRoot -SourceLabel $selection.Tag -Scratch $scratch
    }
    catch {
        $result.Message = ("Icon pack install failed: {0}" -f $_.Exception.Message)
        return $result
    }
    finally {
        if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Complete-IconPackInstall {
    <#
    .SYNOPSIS
        Extracts a verified icon pack zip into Packagers\Icons.

    .OUTPUTS
        [pscustomobject] Installed, Message, Manifest.
    #>
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [string]$AppRoot = $PSScriptRoot,
        [string]$SourceLabel = 'pack',
        [string]$Scratch = ''
    )

    $result = [pscustomobject]@{ Installed = $false; Message = ''; Manifest = $null }
    $destination = Get-IconPackRoot -AppRoot $AppRoot
    if ([string]::IsNullOrWhiteSpace($Scratch)) {
        $Scratch = Join-Path ([IO.Path]::GetTempPath()) ('apicons-' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $Scratch -Force | Out-Null
    }

    $extracted = Join-Path $Scratch 'extracted'
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $extracted -Force

    if (-not (Test-Path -LiteralPath $destination)) { New-Item -ItemType Directory -Path $destination -Force | Out-Null }
    Get-ChildItem -LiteralPath $extracted -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $destination -Recurse -Force
    }

    $manifest = Read-IconPackManifest -Path (Get-IconPackManifestPath -AppRoot $AppRoot)
    $result.Manifest  = $manifest
    $result.Installed = $true

    $warning = ''
    if ($manifest -and -not (Test-IconPackAppVersion -MinAppVersion $manifest.MinAppVersion -CurrentVersion (Get-AppVersion))) {
        $warning = (" Pack targets AppPackager {0} or newer; installed anyway." -f $manifest.MinAppVersion)
    }
    $count = if ($manifest) { $manifest.IconCount } else { 0 }
    $result.Message = ("Icon pack {0} installed into {1} ({2} icons).{3}" -f $SourceLabel, $destination, $count, $warning)
    return $result
}

function Install-IconPackFromFile {
    <#
    .SYNOPSIS
        Installs an icon pack from a local or UNC zip, for hosts whose proxy
        blocks the release download.

    .DESCRIPTION
        A checksums.txt beside the zip is verified when present; without one
        the install proceeds and the message says the pack was unverified.
        The zip is copied to a scratch folder first so the extraction source
        never carries a zone identifier from the original location.

    .OUTPUTS
        [pscustomobject] Installed, Message, Manifest.
    #>
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [string]$AppRoot = $PSScriptRoot
    )

    $result  = [pscustomobject]@{ Installed = $false; Message = ''; Manifest = $null }
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ('apicons-' + [Guid]::NewGuid().ToString('N'))

    try {
        if (-not (Test-Path -LiteralPath $ZipPath -PathType Leaf)) {
            $result.Message = ("Icon pack file not found: {0}" -f $ZipPath)
            return $result
        }

        New-Item -ItemType Directory -Path $scratch -Force | Out-Null
        $localZip = Join-Path $scratch ([IO.Path]::GetFileName($ZipPath))
        Copy-Item -LiteralPath $ZipPath -Destination $localZip -Force
        Unblock-File -LiteralPath $localZip -ErrorAction SilentlyContinue

        $sumsBeside = Join-Path (Split-Path -Parent $ZipPath) $script:IconPackSumsName
        $verifyNote = ' Checksum not verified (no checksums.txt beside the zip).'
        if (Test-Path -LiteralPath $sumsBeside -PathType Leaf) {
            $expected = Get-IconPackChecksum -ChecksumText (Get-Content -LiteralPath $sumsBeside -Raw) -FileName ([IO.Path]::GetFileName($ZipPath))
            if (-not (Test-IconPackChecksum -FilePath $localZip -ExpectedSha256 $expected)) {
                $result.Message = 'Icon pack checksum did not match; nothing was extracted.'
                return $result
            }
            $verifyNote = ''
        }

        $installed = Complete-IconPackInstall -ZipPath $localZip -AppRoot $AppRoot -SourceLabel 'from file' -Scratch $scratch
        $installed.Message += $verifyNote
        return $installed
    }
    catch {
        $result.Message = ("Icon pack install failed: {0}" -f $_.Exception.Message)
        return $result
    }
    finally {
        if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue }
    }
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
# The signing policy and per-profile staging root are read by the CLI
# batch driver below, which runs before the workbench region loads.
function Get-WorkbenchProfileDownloadRoot {
    # Two profiles of one application must never share mutable staging
    # output; the default profile keeps today's paths byte for byte.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$DownloadRoot, [AllowEmptyString()][string]$ProfileId)
    if ([string]::IsNullOrWhiteSpace($ProfileId) -or $ProfileId -eq 'default') { return $DownloadRoot }
    if ([string]::IsNullOrWhiteSpace($DownloadRoot)) { return $DownloadRoot }
    return (Join-Path (Join-Path $DownloadRoot 'profiles') $ProfileId)
}

function Get-WorkbenchSigningPolicy {
    param($Prefs = $script:Prefs)
    try { return $Prefs.ScriptSigning } catch { return $null }
}

function Get-WorkbenchSigningPolicyJson {
    param($Prefs = $script:Prefs)
    $block = Get-WorkbenchSigningPolicy -Prefs $Prefs
    if (-not $block) { return '' }
    return ($block | ConvertTo-Json -Depth 4 -Compress)
}

function Get-WorkbenchSigningPolicyDigest {
    param($Prefs = $script:Prefs)
    $json = Get-WorkbenchSigningPolicyJson -Prefs $Prefs
    if (-not $json) { return '' }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($json))
        return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally { $sha.Dispose() }
}

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
        [string]$DefaultTitleMode = '',
        [string]$SigningJson = '',
        [switch]$Force
    )

    # Unattended builds honor the same policy an interactive one does; a
    # configured Require flag must not be silently absent here.
    $batchWorkbenchRoot = ''
    if (Get-Command -Name 'Get-WorkbenchDataRoot' -ErrorAction SilentlyContinue) {
        try { $batchWorkbenchRoot = [string](Get-WorkbenchDataRoot) } catch { }
    }

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
        $installMode = ''
        $commandsJson = ''
        $titleMode = ''
        if ($ConditionApps) {
            $condProp = $ConditionApps.PSObject.Properties[$baseName]
            if ($condProp) {
                $requirementsJson = ConvertTo-RequirementsJson -Entry $condProp.Value
                $variantsJson = ConvertTo-VariantsJson -Entry $condProp.Value
                $installMode = ConvertTo-InstallModeValue -Entry $condProp.Value
                if ([string]$condProp.Value.TitleMode -in @('IncludeVersion', 'NoVersion')) { $titleMode = [string]$condProp.Value.TitleMode }
            }
        }
        if (-not $titleMode) { $titleMode = $DefaultTitleMode }
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
            $restoreModeEnv = $false
            $previousModeEnv = $null
            $previousTitleModeEnv = $env:APP_PACKAGER_TITLE_MODE

            # The workbench bridge, the same set Invoke-PackagerStage and
            # Invoke-PackagerPackage put on their child's environment block.
            # This path launches through the process environment, so each
            # value is saved here and restored in finally.
            $batchRunSnapshot = ''
            $batchDownloadRoot = $DownloadRoot
            if ($batchWorkbenchRoot -and (Get-Command -Name 'New-RunSnapshot' -ErrorAction SilentlyContinue)) {
                # The snapshot carries the signing policy so the certificate
                # gate runs here, before the child writes any content.
                $batchSigningPolicy = $null
                if (-not [string]::IsNullOrWhiteSpace($SigningJson)) {
                    try { $batchSigningPolicy = $SigningJson | ConvertFrom-Json } catch { $batchSigningPolicy = $null }
                }
                try {
                    $batchAppId = [string](New-ApplicationId -Kind Catalog -ScriptPath $scriptPath)
                    $batchDefinition = Get-ApplicationDefinition -ApplicationId $batchAppId
                    $batchProfileId = [string]$batchDefinition.ActiveProfileId
                    if (-not $batchProfileId) { $batchProfileId = 'default' }
                    $batchDownloadRoot = Get-WorkbenchProfileDownloadRoot -DownloadRoot $DownloadRoot -ProfileId $batchProfileId
                    $batchSnapshot = New-RunSnapshot -ApplicationId $batchAppId -ProfileId $batchProfileId `
                        -Target 'MECM' -SigningPolicy $batchSigningPolicy -PackagerScriptPath $scriptPath -DownloadRoot $batchDownloadRoot
                    $batchRunSnapshot = [string]$batchSnapshot.Path
                }
                catch {
                    if ($_.Exception.Message -match 'SigningCertificateUnavailable') {
                        Write-Log ("[batch] {0}: skipped, the signing policy cannot be met: {1}" -f $baseName, $_.Exception.Message) -Level ERROR
                        $results += [pscustomobject]@{ Name = $baseName; Action = 'Failed'; OldVersion = $lastKnown; NewVersion = $null; Reason = $_.Exception.Message }
                        continue
                    }
                    Write-Log ("[batch] {0}: run snapshot not created: {1}" -f $baseName, $_.Exception.Message) -Level WARN
                }
            }
            $batchBridge = [ordered]@{
                APP_PACKAGER_SIGNING        = $SigningJson
                APP_PACKAGER_RUN_SNAPSHOT   = $batchRunSnapshot
                APP_PACKAGER_WORKBENCH_ROOT = $batchWorkbenchRoot
                APP_PACKAGER_DOWNLOAD_ROOT  = $batchDownloadRoot
            }
            $batchBridgeSaved = @{}
            foreach ($bridgeName in @($batchBridge.Keys)) {
                $batchBridgeSaved[$bridgeName] = [Environment]::GetEnvironmentVariable($bridgeName, 'Process')
            }

            $packagerWorkingDirectory = Split-Path -Parent $scriptPath
            $pushedPackagerLocation = $false
            try {
                $env:APP_PACKAGER_TITLE_MODE = $titleMode
                foreach ($bridgeName in @($batchBridge.Keys)) {
                    $bridgeValue = [string]$batchBridge[$bridgeName]
                    if (-not [string]::IsNullOrWhiteSpace($bridgeValue)) {
                        [Environment]::SetEnvironmentVariable($bridgeName, $bridgeValue, 'Process')
                    }
                }
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
                if (-not [string]::IsNullOrWhiteSpace($installMode)) {
                    $restoreModeEnv = $true
                    $previousModeEnv = $env:APP_PACKAGER_INSTALL_MODE
                    $env:APP_PACKAGER_INSTALL_MODE = $installMode
                }

                Push-Location -LiteralPath $packagerWorkingDirectory
                $pushedPackagerLocation = $true
                & powershell.exe @pkgArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    if ($line) { Write-Log ("[batch:{0}] {1}" -f $baseName, $line) -Level INFO }
                }
            }
            finally {
                $env:APP_PACKAGER_TITLE_MODE = $previousTitleModeEnv
                foreach ($bridgeName in @($batchBridgeSaved.Keys)) {
                    [Environment]::SetEnvironmentVariable($bridgeName, $batchBridgeSaved[$bridgeName], 'Process')
                }
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
                if ($restoreModeEnv) {
                    if ($null -ne $previousModeEnv) {
                        $env:APP_PACKAGER_INSTALL_MODE = $previousModeEnv
                    }
                    else {
                        Remove-Item Env:\APP_PACKAGER_INSTALL_MODE -ErrorAction SilentlyContinue
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
        -DefaultTitleMode  $(if ($prefs -and [bool]$prefs.IncludeVersionInTitle) { 'IncludeVersion' } else { '' }) `
        -SigningJson       (Get-WorkbenchSigningPolicyJson) `
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
$txtAppVersion   = $window.FindName('txtAppVersion')
$pnlUpdate       = $window.FindName('pnlUpdate')
$runUpdateText   = $window.FindName('runUpdateText')
$lnkUpdateAvailable = $window.FindName('lnkUpdateAvailable')
$btnUpdateNow    = $window.FindName('btnUpdateNow')
$toggleTheme     = $window.FindName('toggleTheme')
$txtThemeLabel   = $window.FindName('txtThemeLabel')
$btnCheckLatest  = $window.FindName('btnCheckLatest')
$btnCheckMECM    = $window.FindName('btnCheckMECM')
$btnStage        = $window.FindName('btnStage')
$btnPackage      = $window.FindName('btnPackage')
$btnAddInstaller = $window.FindName('btnAddInstaller')
$btnWorkbench    = $window.FindName('btnWorkbench')
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

$menuEditApplication = New-Object System.Windows.Controls.MenuItem
$menuEditApplication.Header = "Edit application..."
$menuSep0 = New-Object System.Windows.Controls.Separator

$contextMenu.Items.Add($menuEditApplication) | Out-Null
$contextMenu.Items.Add($menuSep0) | Out-Null
$contextMenu.Items.Add($menuOpenLogFolder) | Out-Null
$contextMenu.Items.Add($menuOpenStagedFolder) | Out-Null
$contextMenu.Items.Add($menuOpenNetworkShare) | Out-Null
$contextMenu.Items.Add($menuSep1) | Out-Null
$contextMenu.Items.Add($menuCopyLatestVersion) | Out-Null

$dataGrid.ContextMenu = $contextMenu

# Row entry points into the workbench. Both preselect the clicked row so
# the editor opens on the application the operator pointed at.
$menuEditApplication.Add_Click({
    $row = $dataGrid.SelectedItem
    $base = ''
    if ($row) { $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script) }
    Show-ApplicationWorkbench -Owner $window -PreselectPackagerBase $base
})

$dataGrid.Add_MouseDoubleClick({
    param($s, $e)
    # A double-click inside an editable cell (the selection checkbox, a
    # combo) belongs to that cell, not to the row.
    $source = $e.OriginalSource
    if ($source -is [System.Windows.Controls.Primitives.ToggleButton] -or
        $source -is [System.Windows.Controls.TextBox] -or
        $source -is [System.Windows.Controls.ComboBox]) { return }
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    Show-ApplicationWorkbench -Owner $window -PreselectPackagerBase ([System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script))
})

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
    Update-SidebarForDeploymentTarget
}

# =============================================================================
# Sidebar state for the Deployment Target preference
# -----------------------------------------------------------------------------
# Get-SidebarTargetState holds the decision and touches no WPF types, so it is
# exercised headlessly. Update-SidebarForDeploymentTarget applies it and is the
# only writer of these two buttons' label and tooltips; it runs at launch and
# after every path that can change Prefs.Intune.DeploymentTarget.
# =============================================================================
function Get-SidebarTargetState {
    param([string]$DeploymentTarget)

    if ($DeploymentTarget -eq 'IntuneOnly') {
        return @{
            CheckMecmEnabled  = $false
            CheckMecmToolTip  = "Check ConfigMgr needs a ConfigMgr site. The Deployment Target is Intune only - change it in Options, ConfigMgr Preferences, to use this."
            PackageContent    = 'Publish Apps'
            PackageToolTip    = 'Stage each checked app, build the .intunewin, and publish it to Intune'
            SkipMecmPreflight = $true
        }
    }

    return @{
        CheckMecmEnabled  = $true
        CheckMecmToolTip  = 'Query ConfigMgr for the currently deployed version of each checked app'
        PackageContent    = 'Package Apps'
        PackageToolTip    = 'Build ConfigMgr Application + Deployment Type for every checked app'
        SkipMecmPreflight = $false
    }
}

function Update-SidebarForDeploymentTarget {
    $state = Get-SidebarTargetState -DeploymentTarget ([string]$script:Prefs.Intune.DeploymentTarget)

    # Disabled rather than hidden: the button stays discoverable and its
    # tooltip states why it cannot run.
    # Only the disable is forced here; enablement otherwise stays with
    # Set-ActionButtonsEnabled, which calls back into this function.
    if (-not $state.CheckMecmEnabled) { $btnCheckMECM.IsEnabled = $false }
    $btnCheckMECM.ToolTip = $state.CheckMecmToolTip
    $btnPackage.Content     = $state.PackageContent
    $btnPackage.ToolTip     = $state.PackageToolTip
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
$btnWorkbench.Add_Click({
    $row = $dataGrid.SelectedItem
    $base = ''
    if ($row) { $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script) }
    Show-ApplicationWorkbench -Owner $window -PreselectPackagerBase $base
})

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
<ScrollViewer xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
      VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
<Grid Margin="0,0,4,0">
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

    <TextBlock Grid.Row="7" Grid.Column="0" Text="Auto-distribute:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="When enabled, the Package phase calls Start-CMContentDistribution after creating each ConfigMgr Application."/>
    <CheckBox  Grid.Row="7" Grid.Column="1" x:Name="chkAutoDist" Content="Start-CMContentDistribution after Package" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="8" Grid.Column="0" Text="DP Group:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Exact name of the Distribution Point Group to target."/>
    <TextBox   Grid.Row="8" Grid.Column="1" x:Name="txtDPGroup" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="Distribution Point Group display name (e.g. 'All DPs')"/>

    <TextBlock Grid.Row="9" Grid.Column="0" Text="Test deployment:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Requires Auto-distribute enabled and a DP Group name. After content distribution, deploys the application (Available, immediately, default options) to the test collection."/>
    <CheckBox  Grid.Row="9" Grid.Column="1" x:Name="chkTestDeploy" Content="Deploy to test collection after distribution" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="10" Grid.Column="0" Text="Test collection:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Exact device collection name that receives the Available test deployment."/>
    <TextBox   Grid.Row="10" Grid.Column="1" x:Name="txtTestCollection" FontSize="13" MaxLength="255" Margin="0,0,0,8" ToolTip="Device collection display name (e.g. 'App Test Devices')"/>

    <TextBlock Grid.Row="11" Grid.Column="0" Text="" Margin="0,0,0,8"/>
    <CheckBox  Grid.Row="11" Grid.Column="1" x:Name="chkCreateTestColl" Content="Create collection if it does not exist" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal" ToolTip="Creates an empty direct-membership device collection limited to All Systems when the named collection is missing."/>

    <TextBlock Grid.Row="12" Grid.Column="0" Text="Application title:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Default application naming for every Package run."/>
    <CheckBox  Grid.Row="12" Grid.Column="1" x:Name="chkTitleVersion" Content="Include version in application name" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal" ToolTip="Adds the version to every application name, creating one ConfigMgr application per release. A per-application choice in the Application Workbench overrides this. Existing applications are not renamed."/>

    <TextBlock Grid.Row="13" Grid.Column="0" Text="Console:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Configuration Manager Console (AdminUI) detection status. Checked once per launch."/>
    <Grid Grid.Row="13" Grid.Column="1" MinHeight="26" Margin="0,0,0,8"><TextBlock x:Name="txtConsoleStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/></Grid>

    <TextBlock Grid.Row="14" Grid.Column="0" Text="7-Zip CLI:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="7-Zip command-line (7z.exe) detection status. Required by the Adobe Reader packager."/>
    <Grid Grid.Row="14" Grid.Column="1" MinHeight="26" Margin="0,0,0,8"><TextBlock x:Name="txtSevenZipStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/></Grid>
    <TextBlock Grid.Row="15" Grid.Column="0" Text="GitHub API:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="How the 90 packagers that read GitHub releases authenticate. Anonymous calls are limited to 60 per hour per address; a token raises that to 5000. Resolved from GITHUB_TOKEN, then GH_TOKEN, then the GitHub CLI login (gh auth login)."/>
    <Grid Grid.Row="15" Grid.Column="1" MinHeight="26" Margin="0,0,0,8"><TextBlock x:Name="txtGitHubStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/></Grid>

    <TextBlock Grid.Row="16" Grid.Column="0" Text="Content Prep:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) detection status. Downloaded on first use, or place the exe on PATH."/>
    <!-- Status text in a star column so a long message wraps instead of
         pushing the buttons past the panel edge, where they clip out of view. -->
    <Grid Grid.Row="16" Grid.Column="1" MinHeight="26" Margin="0,0,0,8">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" x:Name="txtIntuneWinStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/>
        <Button Grid.Column="1" x:Name="btnIntuneWinDownload" Content="Download" FontSize="11" Margin="10,0,0,0" Padding="10,2" VerticalAlignment="Center" Visibility="Collapsed"/>
    </Grid>

    <TextBlock Grid.Row="17" Grid.Column="0" Text="Icon Pack:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Packager icon pack for IconSource External packagers. Installs into Packagers\Icons and is read at stage time."/>
    <Grid Grid.Row="17" Grid.Column="1" MinHeight="26" Margin="0,0,0,8">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" x:Name="txtIconPackStatus" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/>
        <Button Grid.Column="1" x:Name="btnIconPackDownload" Content="Download packager icon pack" FontSize="11" Margin="10,0,0,0" Padding="10,2" VerticalAlignment="Center" ToolTip="Fetches the latest pack release, verifies its sha256 against checksums.txt, and extracts it into Packagers\Icons. Existing icons with the same name are replaced."/>
        <Button Grid.Column="2" x:Name="btnIconPackFromFile" Content="Install from file..." FontSize="11" Margin="6,0,0,0" Padding="10,2" VerticalAlignment="Center" ToolTip="Installs an icon pack from a local or UNC icon-pack.zip when the release download is blocked (proxy/SSL inspection). A checksums.txt beside the zip is verified when present."/>
    </Grid>

    <TextBlock Grid.Row="18" Grid.Column="0" Text="Intunewin:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="When enabled, a successful Package also produces an .intunewin from the staged content and stores it beside the network content version folder."/>
    <CheckBox  Grid.Row="18" Grid.Column="1" x:Name="chkIntuneWin" Content="Create .intunewin during Package" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
    <TextBlock Grid.Row="19" Grid.Column="0" Text="Intune Tenant ID:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Entra tenant ID (GUID or domain) for Graph publishing."/>
    <TextBox   Grid.Row="19" Grid.Column="1" x:Name="txtIntuneTenant" FontSize="13" Margin="0,0,0,8"/>
    <TextBlock Grid.Row="20" Grid.Column="0" Text="Intune Client ID:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="App registration (client) ID with application permission DeviceManagementApps.ReadWrite.All, admin-consented."/>
    <TextBox   Grid.Row="20" Grid.Column="1" x:Name="txtIntuneClient" FontSize="13" Margin="0,0,0,8"/>
    <TextBlock Grid.Row="21" Grid.Column="0" Text="Intune Client Secret:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Stored DPAPI-protected for the current Windows user; leave empty to keep the saved secret."/>
    <PasswordBox Grid.Row="21" Grid.Column="1" x:Name="pwdIntuneSecret" FontSize="13" Margin="0,0,0,8"/>
    <TextBlock Grid.Row="22" Grid.Column="0" Text="Deployment Target:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Where Package creates applications. ConfigMgr only: today's flow. ConfigMgr + Intune: ConfigMgr app plus a Graph publish of the .intunewin. Intune only: stage, build the .intunewin, and publish via Graph - no ConfigMgr console, site, or file share needed. Repeat publishes update the existing Intune app."/>
    <ComboBox  Grid.Row="22" Grid.Column="1" x:Name="cboDeployTarget" FontSize="13" Margin="0,0,0,8" Width="260" HorizontalAlignment="Left">
        <ComboBoxItem Content="ConfigMgr only" Tag="MECM"/>
        <ComboBoxItem Content="ConfigMgr + Intune" Tag="MECMAndIntune"/>
        <ComboBoxItem Content="Intune only" Tag="IntuneOnly"/>
    </ComboBox>
</Grid>
</ScrollViewer>
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
    $chkTitleVersion   = $element.FindName('chkTitleVersion')
    $txtConsoleStatus  = $element.FindName('txtConsoleStatus')
    $txtSevenZipStatus = $element.FindName('txtSevenZipStatus')
    $txtGitHubStatus   = $element.FindName('txtGitHubStatus')
    $txtIntuneWinStatus   = $element.FindName('txtIntuneWinStatus')
    $btnIntuneWinDownload = $element.FindName('btnIntuneWinDownload')
    $txtIconPackStatus    = $element.FindName('txtIconPackStatus')
    $btnIconPackDownload  = $element.FindName('btnIconPackDownload')
    $btnIconPackFromFile  = $element.FindName('btnIconPackFromFile')
    $chkIntuneWin         = $element.FindName('chkIntuneWin')
    $txtIntuneTenant      = $element.FindName('txtIntuneTenant')
    $txtIntuneClient      = $element.FindName('txtIntuneClient')
    $pwdIntuneSecret      = $element.FindName('pwdIntuneSecret')
    $cboDeployTarget      = $element.FindName('cboDeployTarget')

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
    $chkTitleVersion.IsChecked   = [bool]$script:Prefs.IncludeVersionInTitle

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
        $txtSevenZipStatus.Text = ([char]0x2717 + " Not detected  -  Adobe Reader requires 7-Zip CLI")
        $txtSevenZipStatus.ToolTip = "Detected once per launch via registry ARP + Program Files\7-Zip"
    }

    $gh = Get-GitHubApiAuthStatus
    $ghQuota = ''
    if ($gh.Limit) {
        $ghQuota = ' (' + $gh.Limit + ' requests/hour'
        if ($gh.Remaining) { $ghQuota += ', ' + $gh.Remaining + ' remaining' }
        $ghQuota += ')'
    }
    if ($gh.Authenticated) {
        $txtGitHubStatus.Text = ([char]0x2713 + ' Authenticated  -  ' + $gh.Source + $ghQuota)
        $txtGitHubStatus.ToolTip = 'The token is sent as a bearer token by every GitHub-backed packager and by the version monitor. Resolved when this window opened.'
    } else {
        $anonRemaining = if ($gh.Remaining) { ', ' + $gh.Remaining + ' remaining' } else { '' }
        $txtGitHubStatus.Text = ([char]0x2717 + ' Anonymous  -  60 requests/hour per address' + $anonRemaining + '; set GITHUB_TOKEN or run gh auth login')
        $txtGitHubStatus.ToolTip = 'A version check across the 90 GitHub-backed packagers exceeds the anonymous limit. A personal access token with no scopes is enough.'
    }
    $chkIntuneWin.IsChecked = [bool]$script:Prefs.Intune.CreateIntuneWin
    $txtIntuneTenant.Text = [string]$script:Prefs.Intune.TenantId
    $txtIntuneClient.Text = [string]$script:Prefs.Intune.ClientId
    $currentTarget = [string]$script:Prefs.Intune.DeploymentTarget
    foreach ($item in $cboDeployTarget.Items) {
        if ([string]$item.Tag -eq $currentTarget) { $cboDeployTarget.SelectedItem = $item; break }
    }
    if (-not $cboDeployTarget.SelectedItem) { $cboDeployTarget.SelectedIndex = 0 }
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

    $updateIconPackState = {
        $txtIconPackStatus.Text = Get-IconPackStatusText -Manifest (Read-IconPackManifest -Path (Get-IconPackManifestPath))
        $txtIconPackStatus.ToolTip = ("Manifest: {0}" -f (Get-IconPackManifestPath))
    }.GetNewClosure()
    & $updateIconPackState

    $btnIconPackDownload.Add_Click({
        $btnIconPackDownload.IsEnabled = $false
        $txtIconPackStatus.Text = 'Downloading packager icon pack...'
        try {
            $outcome = Install-IconPack
            Add-LogLine -Message $outcome.Message
            $failureText = $outcome.Message
        } finally {
            $btnIconPackDownload.IsEnabled = $true
            & $updateIconPackState
            # A failed fetch leaves no manifest change, so the refreshed status
            # line would silently repeat itself; the reason replaces it instead.
            if ($outcome -and -not $outcome.Installed) {
                $txtIconPackStatus.Text = ([char]0x2717 + ' ' + $failureText)
            }
        }
    }.GetNewClosure())

    $btnIconPackFromFile.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title  = 'Select icon pack zip'
        $dlg.Filter = 'Icon pack (*.zip)|*.zip'
        if (-not $dlg.ShowDialog()) { return }
        $btnIconPackFromFile.IsEnabled = $false
        $txtIconPackStatus.Text = 'Installing icon pack from file...'
        try {
            $outcome = Install-IconPackFromFile -ZipPath $dlg.FileName
            Add-LogLine -Message $outcome.Message
            $failureText = $outcome.Message
        } finally {
            $btnIconPackFromFile.IsEnabled = $true
            & $updateIconPackState
            if ($outcome -and -not $outcome.Installed) {
                $txtIconPackStatus.Text = ([char]0x2717 + ' ' + $failureText)
            }
        }
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
        $prefsRef.IncludeVersionInTitle = [bool]$chkTitleVersion.IsChecked
        $prefsRef.Intune.CreateIntuneWin = [bool]$chkIntuneWin.IsChecked
        $prefsRef.Intune.DeploymentTarget = [string]$cboDeployTarget.SelectedItem.Tag
        $prefsRef.Intune.PublishToIntune = ($prefsRef.Intune.DeploymentTarget -ne 'MECM')
        $prefsRef.Intune.TenantId = $txtIntuneTenant.Text.Trim()
        $prefsRef.Intune.ClientId = $txtIntuneClient.Text.Trim()
        # An empty box keeps the stored secret; a typed value replaces it,
        # protected with DPAPI for the current Windows user.
        if ($pwdIntuneSecret.SecurePassword.Length -gt 0) {
            $prefsRef.Intune.ClientSecretProtected = ($pwdIntuneSecret.SecurePassword | ConvertFrom-SecureString)
        }
    }.GetNewClosure()

    return @{ Name = 'ConfigMgr Preferences'; Element = $element; Commit = $commit }
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
    $dbv  = $script:Prefs.DBeaverInstallOptions

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
    & $addLabelRow "UI Mode:" $cmbSsmsUiMode "Quiet adds --quiet for a fully hidden install. Passive adds --passive for progress-only UI and is less suitable for required ConfigMgr deployments."

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
    # DBEAVER COMMUNITY
    # =============================================
    & $addDivider
    & $addHeader "DBeaver Community"

    $cmbDbvScope = New-Object System.Windows.Controls.ComboBox
    $cmbDbvScope.FontSize = 13
    $cmbDbvScope.Width = 120
    $cmbDbvScope.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    foreach ($val in @("System", "User")) { [void]$cmbDbvScope.Items.Add($val) }
    $cmbDbvScope.SelectedItem = if ([string]$dbv.InstallScope -in @('System','User')) { [string]$dbv.InstallScope } else { 'System' }
    & $addLabelRow "Install Scope:" $cmbDbvScope "System installs machine-wide to C:\Program Files\DBeaver with /allusers. User installs to %LOCALAPPDATA%\DBeaver with /currentuser and needs no elevation; deploy that one in user context so detection resolves the right profile."

    $chkDbvDisableAI = & $addCheckBox "Disable AI features (-Dai.disabled=true)" ([bool]$dbv.DisableAI) "Appends -Dai.disabled=true to the installed dbeaver.ini after a successful install, which is DBeaver's documented way to turn off AI assistant features. The line is added once and re-applied after every reinstall."

    # =============================================
    # BEYOND COMPARE 5
    # =============================================
    & $addDivider
    & $addHeader "Beyond Compare 5"

    $txtBc5KeyFile = New-Object System.Windows.Controls.TextBox
    $txtBc5KeyFile.Text = [string]$script:Prefs.BeyondCompareKeyFile
    $txtBc5KeyFile.FontSize = 13
    $txtBc5KeyFile.Width = 350
    $btnBc5KeyBrowse = New-Object System.Windows.Controls.Button
    $btnBc5KeyBrowse.Content = 'Browse...'
    $btnBc5KeyBrowse.MinWidth = 90
    $btnBc5KeyBrowse.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
    $btnBc5KeyBrowse.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, 'MahApps.Styles.Button.Square')
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnBc5KeyBrowse, [System.Windows.Controls.CharacterCasing]::Normal)
    $spBc5Key = New-Object System.Windows.Controls.StackPanel
    $spBc5Key.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    [void]$spBc5Key.Children.Add($txtBc5KeyFile)
    [void]$spBc5Key.Children.Add($btnBc5KeyBrowse)
    & $addLabelRow "License Key File:" $spBc5Key "BC5Key.txt from Scooter Software. Stage places it beside the installer, where setup reads it to register Beyond Compare 5. Required to stage and package Beyond Compare 5."
    $btnBc5KeyBrowse.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Choose the Beyond Compare 5 license key file'
        $dlg.Filter = 'License key file (BC5Key.txt)|BC5Key.txt|Text files (*.txt)|*.txt|All files (*.*)|*.*'
        if ($dlg.ShowDialog() -eq $true) { $txtBc5KeyFile.Text = $dlg.FileName }
    }.GetNewClosure())

    # =============================================
    # LOCAL INSTALLER SOURCES
    # =============================================
    $localSourceBoxes = @{}
    $localSourcePackagers = @(Get-Packagers -Root $PackagersRoot | Where-Object { $_.LocalSource })
    if ($localSourcePackagers.Count -gt 0) {
        & $addDivider
        & $addHeader "Local Installer Sources"
        foreach ($lsp in $localSourcePackagers) {
            $lspBase = [System.IO.Path]::GetFileNameWithoutExtension([string]$lsp.Script)
            $lspBox = New-Object System.Windows.Controls.TextBox
            $lspBox.FontSize = 13
            $lspBox.Width = 350
            if ($script:Prefs.LocalSourceFolders -and $script:Prefs.LocalSourceFolders.PSObject.Properties[$lspBase]) {
                $lspBox.Text = [string]$script:Prefs.LocalSourceFolders.$lspBase
            }
            $lspBrowse = New-Object System.Windows.Controls.Button
            $lspBrowse.Content = 'Browse...'
            $lspBrowse.MinWidth = 90
            $lspBrowse.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
            $lspBrowse.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, 'MahApps.Styles.Button.Square')
            [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($lspBrowse, [System.Windows.Controls.CharacterCasing]::Normal)
            $lspRow = New-Object System.Windows.Controls.StackPanel
            $lspRow.Orientation = [System.Windows.Controls.Orientation]::Horizontal
            [void]$lspRow.Children.Add($lspBox)
            [void]$lspRow.Children.Add($lspBrowse)
            & $addLabelRow ([string]$lsp.Application + ':') $lspRow ("Folder that holds the {0} installer. Its download requires a sign-in, so Stage uses the newest matching installer in this folder. Stage asks for it when this is empty." -f [string]$lsp.Application)
            $lspBrowse.Add_Click({
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
                $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                if ($lspBox.Text -and (Test-Path -LiteralPath $lspBox.Text -PathType Container)) { $dlg.SelectedPath = $lspBox.Text }
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $lspBox.Text = $dlg.SelectedPath }
            }.GetNewClosure())
            $localSourceBoxes[$lspBase] = $lspBox
        }
    }

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

        if (-not $prefsRef.DBeaverInstallOptions) {
            $prefsRef.DBeaverInstallOptions = [pscustomobject]@{
                InstallScope = "System"
                DisableAI    = $false
            }
        }
        $selectedDbvScope = [string]$cmbDbvScope.SelectedItem
        if ($selectedDbvScope -notin @('System','User')) { $selectedDbvScope = 'System' }
        $prefsRef.DBeaverInstallOptions.InstallScope = $selectedDbvScope
        $prefsRef.DBeaverInstallOptions.DisableAI    = ($chkDbvDisableAI.IsChecked -eq $true)

        $prefsRef.BeyondCompareKeyFile = [string]$txtBc5KeyFile.Text.Trim()

        if ($localSourceBoxes.Count -gt 0) {
            $sourceProps = [ordered]@{}
            if ($prefsRef.LocalSourceFolders) {
                foreach ($p in $prefsRef.LocalSourceFolders.PSObject.Properties) { $sourceProps[$p.Name] = $p.Value }
            }
            foreach ($key in $localSourceBoxes.Keys) {
                $value = [string]$localSourceBoxes[$key].Text.Trim()
                if ($value) { $sourceProps[$key] = $value }
                elseif ($sourceProps.Contains($key)) { $sourceProps.Remove($key) }
            }
            $prefsRef.LocalSourceFolders = [pscustomobject]$sourceProps
        }

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

function Show-ExistingConflictDialog {
    # Modal choice for one existing application, at any version. Returns
    # @{ Choice = 'Skip'|'Overwrite'|'Cancel'; ApplyToAll = [bool] }.
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Version,
        [string]$IncomingVersion = '',
        [Parameter(Mandatory)]$Owner
    )
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="560" SizeToContent="Height" MinWidth="460"
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
                <Setter Property="MinWidth" Value="120"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="120"/><Setter Property="Height" Value="32"/>
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
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtIntro" Grid.Row="0" TextWrapping="Wrap" FontSize="12" Margin="0,4,0,12"/>
        <CheckBox  x:Name="chkAll" Grid.Row="1" FontSize="12" Margin="0,0,0,16"
                   Content="Do this for all remaining conflicts in this run"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnSkip"      Content="Skip"          Style="{StaticResource DialogButton}"/>
            <Button x:Name="btnOverwrite" Content="Overwrite"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel"    Content="Cancel run"    Style="{StaticResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    $dlg.Title = 'Application already exists'
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner
    $dlg.FindName('txtIntro').Text = "$AppName already exists at version $Version. The staged version is $IncomingVersion. Overwrite updates its content and deployment types while keeping the application and its deployments. Skip leaves it unchanged; Cancel stops the remaining run."
    $chkAll = $dlg.FindName('chkAll')
    # ShowDialog keeps this function on the stack while the handlers run, so
    # they keep its scope; a GetNewClosure handler writes the choice into its
    # own module and the value returned below stays at Skip.
    $script:ConflictDialogChoice = 'Skip'
    $dlg.FindName('btnSkip').Add_Click({ $script:ConflictDialogChoice = 'Skip'; $dlg.Close() })
    $dlg.FindName('btnOverwrite').Add_Click({ $script:ConflictDialogChoice = 'Overwrite'; $dlg.Close() })
    $dlg.FindName('btnCancel').Add_Click({ $script:ConflictDialogChoice = 'Cancel'; $dlg.Close() })
    [void]$dlg.ShowDialog()
    return @{ Choice = [string]$script:ConflictDialogChoice; ApplyToAll = [bool]$chkAll.IsChecked }
}

function New-AboutPanel {
    $xaml = @'
<ScrollViewer xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      VerticalScrollBarVisibility="Auto">
    <StackPanel>
        <TextBlock Text="AppPackager" FontSize="20" FontWeight="Bold" Margin="0,0,0,2"/>
        <TextBlock x:Name="txtAboutVersion" FontSize="13" Margin="0,0,0,14"
                   Foreground="{DynamicResource MahApps.Brushes.Gray3}"/>
        <TextBlock TextWrapping="Wrap" FontSize="12" Margin="0,0,0,14"
                   Text="Automated application packaging for ConfigMgr and Intune, built entirely in in-box PowerShell 5.1."/>
        <Grid Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="130"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Grid.Column="0" Text="License" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="0" Grid.Column="1" Text="MIT" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="1" Grid.Column="0" Text="Repository" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="1" Grid.Column="1" FontSize="12" Margin="0,0,0,6">
                <Hyperlink x:Name="lnkAboutRepo" ToolTip="Open the project on GitHub">
                    <Run x:Name="runAboutRepo" Text="github.com"/>
                </Hyperlink>
            </TextBlock>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="Last update check" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="txtAboutLastCheck" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="Latest release" FontSize="12"/>
            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="txtAboutLatest" FontSize="12"/>
        </Grid>
        <StackPanel Orientation="Horizontal">
            <Button x:Name="btnAboutUpdate" Content="Update now" MinWidth="110" Height="30" Margin="0,0,8,0"
                    ToolTip="Download and install the latest release, then close and relaunch AppPackager"/>
            <Button x:Name="btnAboutReleases" Content="Release notes" MinWidth="110" Height="30"
                    ToolTip="Open the releases page on GitHub"/>
        </StackPanel>
    </StackPanel>
</ScrollViewer>
'@

    [xml]$xml = $xaml
    $reader  = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $txtAboutVersion   = $element.FindName('txtAboutVersion')
    $txtAboutLastCheck = $element.FindName('txtAboutLastCheck')
    $txtAboutLatest    = $element.FindName('txtAboutLatest')
    $lnkAboutRepo      = $element.FindName('lnkAboutRepo')
    $runAboutRepo      = $element.FindName('runAboutRepo')
    $btnAboutUpdate    = $element.FindName('btnAboutUpdate')
    $btnAboutReleases  = $element.FindName('btnAboutReleases')

    $version = $script:AppVersion
    if ([string]::IsNullOrWhiteSpace($version)) { $version = Get-AppVersion }
    $txtAboutVersion.Text = if ($version) { "Version $version" } else { 'Version unknown' }

    $repoUrl = "https://github.com/$($script:UpdateRepo)"
    $runAboutRepo.Text = $script:UpdateRepo
    $lnkAboutRepo.Add_Click({ try { Start-Process $repoUrl } catch { } }.GetNewClosure())

    $cache = Read-UpdateCheckCache
    $lastCheckText = 'Never'
    $latestText    = 'Unknown'
    if ($cache) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse(
                ([string]$cache.LastCheckUtc), [cultureinfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
            $lastCheckText = $parsed.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        }
        if (-not [string]::IsNullOrWhiteSpace($cache.LatestVersion)) {
            $latestText = "v$($cache.LatestVersion)"
            if (-not (Test-UpdateAvailable -CurrentVersion $version -LatestVersion $cache.LatestVersion)) {
                $latestText += ' (up to date)'
            }
        }
    }
    $txtAboutLastCheck.Text = $lastCheckText
    $txtAboutLatest.Text    = $latestText

    # Update now stays inert until a check has actually found a newer release,
    # so the button can never install over a current folder for no reason.
    $btnAboutUpdate.IsEnabled = [bool]$script:UpdateLatestVersion
    $btnAboutUpdate.Add_Click({
        Invoke-SelfUpdate -Owner ([System.Windows.Window]::GetWindow($btnAboutUpdate))
    }.GetNewClosure())

    $releasesUrl = "$repoUrl/releases"
    $btnAboutReleases.Add_Click({ try { Start-Process $releasesUrl } catch { } }.GetNewClosure())

    return @{ Name = 'About'; Element = $element; Commit = { } }
}

function Show-OptionsDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$InitialSection = 'ConfigMgr Preferences'
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="1160" Height="760"
    MinWidth="1000" MinHeight="600"
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
        (New-ScriptSigningPanel),
        (New-AboutPanel)
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
            $siteBefore = '{0}|{1}' -f $script:Prefs.SiteCode, $script:Prefs.ProviderMachineName
            foreach ($p in $panels) { if ($p.Commit) { & $p.Commit } }
            $siteChanged = ('{0}|{1}' -f $script:Prefs.SiteCode, $script:Prefs.ProviderMachineName) -ne $siteBefore
            Save-Preferences -Prefs $script:Prefs
            # Panels that mutate sibling JSON configs expose the refs on
            # the panel hash; master persists them here so panel commits
            # stay free of function calls (GetNewClosure-safe).
            foreach ($p in $panels) {
                if ($p.CwaSwitches) { Save-CwaSwitches -Switches $p.CwaSwitches }
                if ($p.TvConfig)    { Save-TvHostConfig -Config $p.TvConfig }
            }
            Invoke-RefreshGrid -DiscardSiteResults:$siteChanged
            Update-SidebarForDeploymentTarget
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
# First-run setup wizard - deployment target plus the minimum settings that
# target needs. Writes through the same preference keys and persistence path
# as the Options window; the target selection shows or hides the two groups.
# =============================================================================
function Show-FirstRunWizard {
    param([Parameter(Mandatory)]$Owner)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Setup"
    Width="720" Height="560"
    MinWidth="640" MinHeight="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ShowMaxRestoreButton="False"
    ShowMinButton="False"
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
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ScrollViewer Grid.Row="0" Margin="20,18,20,12" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel>
                <TextBlock Text="Welcome to AppPackager" FontSize="18" FontWeight="Bold" Margin="0,0,0,6"/>
                <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="{DynamicResource MahApps.Brushes.Gray3}" Margin="0,0,0,16"
                           Text="Pick where packaged applications should land, then fill in the settings that target needs. Everything here can be changed later in Options - ConfigMgr Preferences."/>

                <TextBlock Text="Environment" FontSize="13" FontWeight="Bold" Margin="0,0,0,6"/>
                <ComboBox x:Name="cboTarget" FontSize="13" Width="280" HorizontalAlignment="Left" Margin="0,0,0,16"
                          ToolTip="Where Package creates applications. ConfigMgr only: today's flow. ConfigMgr + Intune: ConfigMgr app plus a Graph publish of the .intunewin. Intune only: stage, build the .intunewin, and publish via Graph - no ConfigMgr console, site, or file share needed. Repeat publishes update the existing Intune app.">
                    <ComboBoxItem Content="ConfigMgr only" Tag="MECM"/>
                    <ComboBoxItem Content="ConfigMgr + Intune" Tag="MECMAndIntune"/>
                    <ComboBoxItem Content="Intune only" Tag="IntuneOnly"/>
                </ComboBox>

                <GroupBox x:Name="grpMecm" Header="ConfigMgr" Margin="0,0,0,14">
                    <Grid Margin="0,8,0,4">
                        <Grid.RowDefinitions>
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
                        <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtWizSC" Width="80" FontSize="13" HorizontalAlignment="Left" MaxLength="5" Margin="0,0,0,8" ToolTip="ConfigMgr site code PSDrive name (e.g., MCM)"/>

                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Provider Machine:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
                        <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtWizProvider" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="SMS Provider server from the ConfigMgr AdminUI connect script's ProviderMachineName value"/>

                        <TextBlock Grid.Row="2" Grid.Column="0" Text="File Share Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
                        <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtWizFS" FontSize="13" MaxLength="200" Margin="0,0,0,8" ToolTip="UNC path to the SCCM content file share"/>

                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Download Root:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,4"/>
                        <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtWizDL" FontSize="13" MaxLength="200" Margin="0,0,0,4" ToolTip="Local folder where installers are downloaded during staging"/>
                    </Grid>
                </GroupBox>

                <GroupBox x:Name="grpIntune" Header="Intune" Margin="0,0,0,4">
                    <Grid Margin="0,8,0,4">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="140"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Tenant ID:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="Entra tenant ID (GUID or domain) for Graph publishing."/>
                        <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtWizTenant" FontSize="13" Margin="0,0,0,8" ToolTip="Entra tenant ID (GUID or domain) for Graph publishing."/>

                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Client ID:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8" ToolTip="App registration (client) ID with application permission DeviceManagementApps.ReadWrite.All, admin-consented."/>
                        <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtWizClient" FontSize="13" Margin="0,0,0,8" ToolTip="App registration (client) ID with application permission DeviceManagementApps.ReadWrite.All, admin-consented."/>

                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Client Secret:" FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,4" ToolTip="Stored DPAPI-protected for the current Windows user; leave empty to keep the saved secret."/>
                        <PasswordBox Grid.Row="2" Grid.Column="1" x:Name="pwdWizSecret" FontSize="13" Margin="0,0,0,4" ToolTip="Stored DPAPI-protected for the current Windows user; leave empty to keep the saved secret."/>
                    </Grid>
                </GroupBox>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="1" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <Grid Margin="20,12,20,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="chkWizDontAsk" Content="Don't show this again" FontSize="12" VerticalAlignment="Center"
                          Controls:ControlsHelper.ContentCharacterCasing="Normal"
                          ToolTip="Suppresses this wizard on future launches even when nothing is configured here. Saving always suppresses it."/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnWizSave" Content="Save" MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True" Style="{DynamicResource MahApps.Styles.Button.Square.Accent}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <Button x:Name="btnWizSkip" Content="Skip" MinWidth="90" Height="32" IsCancel="True" Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $cboTarget      = $dlg.FindName('cboTarget')
    $grpMecm        = $dlg.FindName('grpMecm')
    $grpIntune      = $dlg.FindName('grpIntune')
    $txtWizSC       = $dlg.FindName('txtWizSC')
    $txtWizProvider = $dlg.FindName('txtWizProvider')
    $txtWizFS       = $dlg.FindName('txtWizFS')
    $txtWizDL       = $dlg.FindName('txtWizDL')
    $txtWizTenant   = $dlg.FindName('txtWizTenant')
    $txtWizClient   = $dlg.FindName('txtWizClient')
    $pwdWizSecret   = $dlg.FindName('pwdWizSecret')
    $chkWizDontAsk  = $dlg.FindName('chkWizDontAsk')
    $btnWizSave     = $dlg.FindName('btnWizSave')
    $btnWizSkip     = $dlg.FindName('btnWizSkip')

    $txtWizSC.Text       = [string]$script:Prefs.SiteCode
    $txtWizProvider.Text = [string]$script:Prefs.ProviderMachineName
    $txtWizFS.Text       = [string]$script:Prefs.FileShareRoot
    $txtWizDL.Text       = [string]$script:Prefs.DownloadRoot
    $txtWizTenant.Text   = [string]$script:Prefs.Intune.TenantId
    $txtWizClient.Text   = [string]$script:Prefs.Intune.ClientId

    $currentTarget = [string]$script:Prefs.Intune.DeploymentTarget
    foreach ($item in $cboTarget.Items) {
        if ([string]$item.Tag -eq $currentTarget) { $cboTarget.SelectedItem = $item; break }
    }
    if (-not $cboTarget.SelectedItem) { $cboTarget.SelectedIndex = 0 }

    $updateGroupState = {
        $tag = if ($cboTarget.SelectedItem) { [string]$cboTarget.SelectedItem.Tag } else { 'MECM' }
        $grpMecm.Visibility   = if ($tag -eq 'IntuneOnly') { 'Collapsed' } else { 'Visible' }
        $grpIntune.Visibility = if ($tag -eq 'MECM') { 'Collapsed' } else { 'Visible' }
    }.GetNewClosure()
    & $updateGroupState
    $cboTarget.Add_SelectionChanged($updateGroupState)

    $script:FirstRunDlgSaved = $false
    $prefsRef = $script:Prefs
    # ShowDialog keeps this function on the stack while these handlers run.
    # Keep its scope: GetNewClosure would hide the script-local save/refresh
    # helpers and give each handler a separate $script:FirstRunDlgSaved flag.
    $btnWizSave.Add_Click({
        try {
            $target = [string]$cboTarget.SelectedItem.Tag
            if ($target -ne 'IntuneOnly') {
                $prefsRef.SiteCode            = $txtWizSC.Text.Trim()
                $prefsRef.ProviderMachineName = $txtWizProvider.Text.Trim()
                $prefsRef.FileShareRoot       = $txtWizFS.Text.Trim()
                $prefsRef.DownloadRoot        = $txtWizDL.Text.Trim()
            }
            if ($target -ne 'MECM') {
                $prefsRef.Intune.TenantId = $txtWizTenant.Text.Trim()
                $prefsRef.Intune.ClientId = $txtWizClient.Text.Trim()
                # An empty box keeps the stored secret; a typed value replaces it,
                # protected with DPAPI for the current Windows user.
                if ($pwdWizSecret.SecurePassword.Length -gt 0) {
                    $prefsRef.Intune.ClientSecretProtected = ($pwdWizSecret.SecurePassword | ConvertFrom-SecureString)
                }
            }
            $prefsRef.Intune.DeploymentTarget = $target
            $prefsRef.Intune.PublishToIntune  = ($target -ne 'MECM')
            $prefsRef.FirstRunCompleted       = $true
            Save-Preferences -Prefs $prefsRef
            Invoke-RefreshGrid
            Update-SidebarForDeploymentTarget
            $script:FirstRunDlgSaved = $true
            $dlg.Close()
        } catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save Failed' -Message $_.Exception.Message -Buttons OK -Icon Error)
        }
    })

    # Skip and window close share one path: the flag persists only when the
    # suppression box is checked, otherwise the wizard returns next launch.
    $dlg.Add_Closing({
        if ($script:FirstRunDlgSaved) { return }
        if ([bool]$chkWizDontAsk.IsChecked) {
            $prefsRef.FirstRunCompleted = $true
            try { Save-Preferences -Prefs $prefsRef } catch { }
        }
    })

    $btnWizSkip.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog()

    if ($script:FirstRunDlgSaved) {
        Add-LogLine -Message ("Setup complete: deployment target = {0}." -f [string]$script:Prefs.Intune.DeploymentTarget)
    }
}

# =============================================================================
# Grid refresh helper
# =============================================================================
function Invoke-RefreshGrid {
    # ConfigMgr versions and compare results describe one site; a site change
    # must not carry them onto rows that now point elsewhere.
    param([switch]$DiscardSiteResults)

    $session = @{}
    foreach ($row in @($script:PackagerData)) { $session[[string]$row.Script] = $row }
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

        $newRow = [pscustomobject]@{
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
        }
        $prior = $session[[string]$m.Script]
        if ($prior) {
            $newRow.Selected = [bool]$prior.Selected
            if ([string]$prior.LatestVersion) { $newRow.LatestVersion = [string]$prior.LatestVersion }
            if ([string]$prior.LastChecked)   { $newRow.LastChecked   = [string]$prior.LastChecked }
            if (-not $DiscardSiteResults -and -not ([string]$m.Status).StartsWith('Read error')) {
                $newRow.CurrentVersion = [string]$prior.CurrentVersion
                $newRow.Status         = $prior.Status
            }
        }
        $script:PackagerData.Add($newRow)
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
    # The workbench module comes too: One Click freshness reads build
    # records from the background loop, not from the UI thread.
    $modulePath = @(
        (Join-Path $PSScriptRoot 'Packagers\AppPackagerCommon.psm1')
        (Join-Path $PSScriptRoot 'Packagers\AppPackagerWorkbench.psm1')
        (Join-Path $PSScriptRoot 'Packagers\AppPackagerSigning.psm1')
    )
    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath)
        $paths = @($ModulePath)
        Import-Module -Name $paths[0] -Force -DisableNameChecking
        foreach ($path in $paths[1..($paths.Count - 1)]) {
            if (Test-Path -LiteralPath $path) {
                Import-Module -Name $path -Force -DisableNameChecking -ErrorAction SilentlyContinue
            }
        }
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
        # Existing-application overwrite handshake. ConflictDecisionForAll
        # lives for this run only; nothing about it is persisted.
        ConflictRequest        = $null
        ConflictResponse       = $null
        ConflictDecisionForAll = ''
    })
    $script:ConflictPromptOpen = $false

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

                # The run plan is prebuilt on the UI thread; an application
                # with no entry keeps the legacy default-profile paths.
                $plan = $null
                if ($Ctx.RunPlanByApp -and $Ctx.RunPlanByApp.ContainsKey($baseName)) { $plan = $Ctx.RunPlanByApp[$baseName] }
                $planSnapshot = ''
                $planDataRoot = ''
                $planDownloadRoot = [string]$Ctx.DownloadRoot
                $planProfileId = 'default'
                $planRevision = 0
                $planApplicationId = ''
                if ($plan) {
                    $planSnapshot = [string]$plan.SnapshotPath
                    $planDataRoot = [string]$plan.DataRoot
                    $planProfileId = [string]$plan.ProfileId
                    $planRevision = [int]$plan.ProfileRevision
                    $planApplicationId = [string]$plan.ApplicationId
                    if ([string]$plan.DownloadRoot) { $planDownloadRoot = [string]$plan.DownloadRoot }
                }
                $planSigning = [string]$Ctx.SigningJson

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
                                -DownloadRoot $planDownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -SevenZipPath $Ctx.SevenZipPath `
                                -VariantsJson $(if ($Ctx.VariantsByApp) { [string]$Ctx.VariantsByApp[$baseName] } else { '' }) `
                                -InstallMode $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' }) `
                                -RunSnapshotPath $planSnapshot -SigningJson $planSigning -WorkbenchDataRoot $planDataRoot

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
                            $packageArgs = @{
                                PackagerPath         = $path
                                SiteCode             = $Ctx.SiteCode
                                ProviderMachineName  = $Ctx.ProviderMachineName
                                Comment              = $Ctx.Comment
                                FileServerPath       = $Ctx.FileShareRoot
                                LogFolder            = $Ctx.LogFolder
                                DownloadRoot         = $planDownloadRoot
                                RunSnapshotPath      = $planSnapshot
                                SigningJson          = $planSigning
                                WorkbenchDataRoot    = $planDataRoot
                                BuildId              = [string]$Ctx.BuildId
                                M365Channel          = $Ctx.M365Channel
                                M365DeployMode       = $Ctx.M365DeployMode
                                EstimatedRuntimeMins = $Ctx.EstimatedRuntimeMins
                                MaximumRuntimeMins   = $Ctx.MaximumRuntimeMins
                                SevenZipPath         = $Ctx.SevenZipPath
                                CreateIntuneWin      = [bool]$Ctx.IntuneWinCreate
                                IntuneWinToolPath    = [string]$Ctx.IntuneWinToolPath
                                IntunePublishConfig  = $Ctx.IntunePublishConfig
                                DeploymentTarget     = $(if ([string]$Ctx.DeploymentTarget -in @('MECM','MECMAndIntune','IntuneOnly')) { [string]$Ctx.DeploymentTarget } else { 'MECM' })
                                ContentLayout        = [string]$Ctx.ContentLayout
                                RequirementsJson     = $reqJson
                                VariantsJson         = $varJson
                                CommandsJson         = $cmdJson
                                InstallMode          = $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' })
                            }
                            $packageArgs['TitleMode'] = $(if ($Ctx.TitleModesByApp -and [string]$Ctx.TitleModesByApp[$baseName]) { [string]$Ctx.TitleModesByApp[$baseName] } else { [string]$Ctx.DefaultTitleMode })
                            $res = Invoke-PackagerPackageWithConflictPrompt -State $State -AppLabel $app -PackageArgs $packageArgs

                            if ($res.PSObject.Properties['PackageOutcome'] -and $res.PackageOutcome -in @('Skipped', 'Canceled')) {
                                $row.Status = [string]$res.PackageOutcome
                                $counts['Skipped']++
                            } elseif ($res.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                if ($res.PSObject.Properties['IntunePublish'] -and $res.IntunePublish) {
                                    [void]$State.LogQueue.Enqueue(('Intune publish: ' + $res.IntunePublish.Message))
                                }
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
                        # Cadence gate (Report only), ConfigMgr pre-flight, then Stage + optional Package.
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

                        # 1a. ConfigMgr pre-flight for Stage/StageAndPackage.
                        # IntuneOnly has no site to query; the query would open a
                        # provider connection, so it is skipped before that.
                        if ($Ctx.Action -in @('Stage','StageAndPackage') -and [string]$Ctx.DeploymentTarget -eq 'IntuneOnly') {
                            [void]$State.LogQueue.Enqueue(('ConfigMgr pre-flight skipped for {0}: deployment target is Intune only.' -f $app))
                        }
                        elseif ($Ctx.Action -in @('Stage','StageAndPackage')) {
                            $cmName = [string]$row.CMName
                            if (-not [string]::IsNullOrWhiteSpace($cmName) -and $Ctx.AdminUiFound) {
                                try {
                                    $mecmRes = Get-MecmCurrentVersionByCMName -SiteCode $Ctx.SiteCode -ProviderMachineName $Ctx.ProviderMachineName -CMName $cmName
                                    if ($mecmRes.Found -and -not [string]::IsNullOrWhiteSpace([string]$mecmRes.SoftwareVersion)) {
                                        $row.CurrentVersion = [string]$mecmRes.SoftwareVersion
                                        $cmp = Compare-SemVer -A ([string]$mecmRes.SoftwareVersion) -B $latest
                                        if ($cmp -eq 0) {
                                            $row.Status = 'Up to date (ConfigMgr)'
                                            [void]$State.LogQueue.Enqueue(('ConfigMgr already has {0} at {1} - skipping' -f $app, $latest))
                                            try { Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result NoChange } catch { }
                                            $counts['NoChange']++
                                            continue
                                        }
                                    }
                                } catch {
                                    [void]$State.LogQueue.Enqueue(('ConfigMgr pre-flight for {0} failed: {1}' -f $app, $_.Exception.Message))
                                }
                            }
                        }

                        $versionChanged = (-not $lastKnown) -or ($lastKnown -ne $latest)
                        $neverStaged    = ($Ctx.Action -eq 'Stage'           -and -not $lastStaged)
                        $neverPackaged  = ($Ctx.Action -eq 'StageAndPackage' -and -not $lastPackaged)

                        # A vendor-version match alone is not freshness: a
                        # changed profile revision or signing policy makes
                        # the last build stale even at the same version.
                        $buildStale = $false
                        if ($planApplicationId) {
                            $current = Test-WorkbenchBuildIsCurrent -ApplicationId $planApplicationId -ProfileId $planProfileId `
                                -Version $latest -ProfileRevision $planRevision -PolicyDigest ([string]$Ctx.SigningDigest)
                            if ($null -ne $current -and -not $current) { $buildStale = $true }
                        }
                        $shouldAct      = $versionChanged -or $Ctx.ForceFlag -or $neverStaged -or $neverPackaged -or $buildStale
                        if ($buildStale -and -not $versionChanged) {
                            [void]$State.LogQueue.Enqueue(('Rebuilding {0} at the same version: profile revision or signing policy changed.' -f $app))
                        }

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
                                -DownloadRoot $planDownloadRoot `
                                -M365Channel $Ctx.M365Channel `
                                -M365DeployMode $Ctx.M365DeployMode `
                                -SevenZipPath $Ctx.SevenZipPath `
                                -VariantsJson $(if ($Ctx.VariantsByApp) { [string]$Ctx.VariantsByApp[$baseName] } else { '' }) `
                                -InstallMode $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' }) `
                                -RunSnapshotPath $planSnapshot -SigningJson $planSigning -WorkbenchDataRoot $planDataRoot

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
                            $packageArgs = @{
                                PackagerPath         = $path
                                SiteCode             = $Ctx.SiteCode
                                ProviderMachineName  = $Ctx.ProviderMachineName
                                Comment              = $Ctx.Comment
                                FileServerPath       = $Ctx.FileShareRoot
                                LogFolder            = $Ctx.LogFolder
                                DownloadRoot         = $planDownloadRoot
                                RunSnapshotPath      = $planSnapshot
                                SigningJson          = $planSigning
                                WorkbenchDataRoot    = $planDataRoot
                                BuildId              = [string]$Ctx.BuildId
                                M365Channel          = $Ctx.M365Channel
                                M365DeployMode       = $Ctx.M365DeployMode
                                EstimatedRuntimeMins = $Ctx.EstimatedRuntimeMins
                                MaximumRuntimeMins   = $Ctx.MaximumRuntimeMins
                                SevenZipPath         = $Ctx.SevenZipPath
                                CreateIntuneWin      = [bool]$Ctx.IntuneWinCreate
                                IntuneWinToolPath    = [string]$Ctx.IntuneWinToolPath
                                IntunePublishConfig  = $Ctx.IntunePublishConfig
                                DeploymentTarget     = $(if ([string]$Ctx.DeploymentTarget -in @('MECM','MECMAndIntune','IntuneOnly')) { [string]$Ctx.DeploymentTarget } else { 'MECM' })
                                ContentLayout        = [string]$Ctx.ContentLayout
                                RequirementsJson     = $reqJson
                                VariantsJson         = $varJson
                                CommandsJson         = $cmdJson
                                InstallMode          = $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' })
                            }
                            $packageArgs['TitleMode'] = $(if ($Ctx.TitleModesByApp -and [string]$Ctx.TitleModesByApp[$baseName]) { [string]$Ctx.TitleModesByApp[$baseName] } else { [string]$Ctx.DefaultTitleMode })
                            $pkg = Invoke-PackagerPackageWithConflictPrompt -State $State -AppLabel $app -PackageArgs $packageArgs

                            if ($pkg.PSObject.Properties['PackageOutcome'] -and $pkg.PackageOutcome -in @('Skipped', 'Canceled')) {
                                $row.Status = [string]$pkg.PackageOutcome
                                $counts['Skipped']++
                            } elseif ($pkg.ExitCode -eq 0) {
                                $row.Status = 'Packaged'
                                [void]$State.LogQueue.Enqueue(('Packaged. Logs: ' + (Split-Path -Leaf $pkg.OutLog)))
                                if ($pkg.PSObject.Properties['IntunePublish'] -and $pkg.IntunePublish) {
                                    [void]$State.LogQueue.Enqueue(('Intune publish: ' + $pkg.IntunePublish.Message))
                                }
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
        # Existing-application conflict: the bg loop is parked waiting for an
        # answer. ConflictPromptOpen guards against a second modal being
        # opened by the next tick while this one is still on screen.
        if ($script:BgState -and $script:BgState.ConflictRequest -and -not $script:ConflictPromptOpen) {
            $script:ConflictPromptOpen = $true
            try {
                $req = $script:BgState.ConflictRequest
                $answer = Show-ExistingConflictDialog -AppName ([string]$req.AppName) -Version ([string]$req.Version) -IncomingVersion ([string]$req.IncomingVersion) -Owner $window
                $script:BgState.ConflictResponse = [pscustomobject]@{
                    Choice     = [string]$answer.Choice
                    ApplyToAll = [bool]$answer.ApplyToAll
                }
            }
            catch {
                # A dialog failure must not park the pipeline forever.
                $script:BgState.ConflictResponse = [pscustomobject]@{ Choice = 'Skip'; ApplyToAll = $false }
                Add-LogLine -Message ('Conflict prompt failed, keeping the existing application: ' + $_.Exception.Message)
            }
            finally {
                $script:ConflictPromptOpen = $false
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
    Width="720" Height="560"
    MinWidth="660" MinHeight="480"
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
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Application" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtDropAppName" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Publisher" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtDropPublisher" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="2" Grid.Column="0" Text="Version" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtDropVersion" Height="28" Margin="0,0,0,8"/>

            <!-- Shown only when the installer script accepts a mode switch; the
                 two radios swap every mode-dependent field together. -->
            <TextBlock Grid.Row="3" Grid.Column="0" x:Name="lblDropMode" Text="Install for" VerticalAlignment="Center" Margin="0,0,12,8" Visibility="Collapsed"/>
            <StackPanel Grid.Row="3" Grid.Column="1" x:Name="pnlDropMode" Orientation="Horizontal" Margin="0,0,0,8" Visibility="Collapsed">
                <RadioButton x:Name="radDropCurrentUser" Content="Current user" GroupName="DropMode" VerticalAlignment="Center" Margin="0,0,18,0" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <RadioButton x:Name="radDropAllUsers" Content="All users (system)" GroupName="DropMode" VerticalAlignment="Center" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <TextBlock x:Name="txtDropModeSwitch" VerticalAlignment="Center" Margin="14,0,0,0" Opacity="0.75"/>
            </StackPanel>

            <TextBlock Grid.Row="4" Grid.Column="0" Text="Install args" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="4" Grid.Column="1" x:Name="txtDropInstallArgs" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="5" Grid.Column="0" x:Name="lblDropUninstall" Text="Uninstall command" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="5" Grid.Column="1" x:Name="txtDropUninstallCmd" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="6" Grid.Column="0" Text="Detection" VerticalAlignment="Top" Margin="0,4,12,0"/>
            <TextBlock Grid.Row="6" Grid.Column="1" x:Name="txtDropDetection" TextWrapping="Wrap" Opacity="0.75" Margin="0,4,0,0"/>
        </Grid>

        <CheckBox Grid.Row="3" x:Name="chkDropConfirm" Margin="0,12,0,0"
                  Content="I verified the predicted silent switches and detection for this installer"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnDropSaveApplication" Content="Save as Application" MinWidth="150" Height="32" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource MahApps.Styles.Button.Square}"
                    ToolTip="Keep this installer as a durable application so it reopens in the Application Workbench without a download script. Its update source stays manual."/>
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

    $lblMode      = $dlg.FindName('lblDropMode')
    $pnlMode      = $dlg.FindName('pnlDropMode')
    $radCurrent   = $dlg.FindName('radDropCurrentUser')
    $radAllUsers  = $dlg.FindName('radDropAllUsers')
    $txtModeSw    = $dlg.FindName('txtDropModeSwitch')

    $isMsi = ([string]$Analysis.InstallerType -eq 'MSI')
    $script:DropActiveAnalysis = $Analysis

    $txtFile.Text      = [string]$Analysis.FileName
    $txtAppName.Text   = [string]$Analysis.AppName
    $txtPublisher.Text = [string]$Analysis.Publisher
    $txtVersion.Text   = [string]$Analysis.SoftwareVersion

    # Every mode-dependent field is written from one analysis object, so a
    # switch of install mode replaces install arguments, uninstall command,
    # context and detection together and never leaves a mixed pair behind.
    $applyAnalysis = {
        param($a)
        $script:DropActiveAnalysis = $a
        $contextText = ''
        if ($a.PSObject.Properties['InstallContext'] -and [string]$a.InstallContext) {
            $contextText = if ([string]$a.InstallContext -eq 'PerUser') { '   Context: per-user' } else { '   Context: per-machine' }
        }
        $txtDetected.Text = ('Detected: {0}   Confidence: {1}   Architecture: {2}{3}' -f `
            $a.InstallerType, $a.Confidence, $a.Architecture, $contextText)
        $txtArgs.Text   = [string]$a.InstallArgs
        $txtUninst.Text = [string]$a.UninstallCommand

        if ($isMsi) {
            # MSI wrappers always run msiexec /qn /norestart against the file;
            # detection is the ProductCode ARP key. Nothing to edit or confirm.
            $txtArgs.IsReadOnly = $true
            $txtUninst.Visibility = [System.Windows.Visibility]::Collapsed
            $lblUninst.Visibility = [System.Windows.Visibility]::Collapsed
            $chkConfirm.Visibility = [System.Windows.Visibility]::Collapsed
            $txtDetect.Text = 'Registry: ARP key for ProductCode ' + $a.ProductCode + ' (DisplayVersion match)'
        }
        else {
            $predicted = [string]$a.UninstallRegistryKey
            $source = 'predicted'
            if ($a.PSObject.Properties['PackageMetadata'] -and $a.PackageMetadata -and
                $a.PackageMetadata.PSObject.Properties['HeaderAvailable'] -and $a.PackageMetadata.HeaderAvailable -and
                $a.PackageMetadata.PSObject.Properties['UninstallRegistryKey'] -and $a.PackageMetadata.UninstallRegistryKey) {
                $source = 'from the installer script'
            }
            if ([string]::IsNullOrWhiteSpace($predicted)) { $predicted = 'ARP key derived from the application name (verify after first install)' }
            $viewNote = if ($a.PSObject.Properties['RegistryView'] -and [string]$a.RegistryView -eq '32') { ', 32-bit view' } else { '' }
            $ctxNote = if ([string]$a.InstallContext -eq 'PerUser') { '. Per-user: the deployment type installs and detects in the user context' } else { '' }
            $txtDetect.Text = ('Registry ({0}): {1} (DisplayVersion match{2}){3}' -f $source, $predicted, $viewNote, $ctxNote)
        }
    }
    & $applyAnalysis $Analysis

    # The mode toggle appears only for an installer whose script accepts a
    # mode switch (NSIS /allusers, Inno Setup /ALLUSERS); the default branch
    # is what the installer does without the switch.
    $modes = @()
    if ($Analysis.PSObject.Properties['InstallModes'] -and $Analysis.InstallModes) { $modes = @($Analysis.InstallModes) }
    if ($modes.Count -gt 1) {
        $lblMode.Visibility = [System.Windows.Visibility]::Visible
        $pnlMode.Visibility = [System.Windows.Visibility]::Visible
        $describeSwitch = {
            param($mode)
            $v = $Analysis.ModeVariants[$mode]
            $defaultMode = [string]$Analysis.InstallMode
            if ($mode -eq $defaultMode) { 'installer default' } else { 'adds ' + ((([string]$v.InstallArgs) -split '\s+') | Where-Object { $_ -match '^/(allusers|currentuser)$' } | Select-Object -First 1) }
        }
        if ([string]$Analysis.InstallMode -eq 'AllUsers') { $radAllUsers.IsChecked = $true } else { $radCurrent.IsChecked = $true }
        $txtModeSw.Text = (& $describeSwitch ([string]$Analysis.InstallMode))
        $switchMode = {
            param($mode)
            try {
                $applied = Set-InstallerAnalysisMode -Analysis $Analysis -Mode $mode
                & $applyAnalysis $applied
                $txtModeSw.Text = (& $describeSwitch $mode)
            }
            catch {
                [void](Show-ThemedMessage -Owner $dlg -Title 'Install Mode' -Message $_.Exception.Message -Buttons OK -Icon Warning)
            }
        }
        $radCurrent.Add_Checked({ & $switchMode 'CurrentUser' })
        $radAllUsers.Add_Checked({ & $switchMode 'AllUsers' })
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
        $script:DropIntakeResult = @{ Action = $action; Values = (& $readValues); Analysis = $script:DropActiveAnalysis }
        $dlg.Close()
    }

    $btnStageOnly.Add_Click({ & $chooseAction 'Stage' })
    $btnStagePkg.Add_Click({ & $chooseAction 'StageAndPackage' })
    $btnSave.Add_Click({ & $chooseAction 'SavePackager' })

    # Persisting the drop as an application creates no download script and
    # deploys nothing; the installer becomes source revision 1.
    $btnSaveApp = $dlg.FindName('btnDropSaveApplication')
    $btnSaveApp.IsEnabled = [bool](Get-Command -Name 'Save-ByoApplication' -ErrorAction SilentlyContinue)
    if (-not $btnSaveApp.IsEnabled) {
        $btnSaveApp.ToolTip = 'The workbench definition model is not loaded, so a persistent application cannot be saved.'
    }
    $btnSaveApp.Add_Click({
        $problem = & $validate
        if ($problem) {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Missing Value' -Message $problem -Buttons OK -Icon Warning)
            return
        }
        $v = & $readValues
        try {
            $saved = Save-ByoApplication -InstallerPath ([string]$Analysis.Path) -DisplayName ([string]$v.AppName) `
                -Publisher ([string]$v.Publisher) -SoftwareVersion ([string]$v.SoftwareVersion) -Analysis $script:DropActiveAnalysis
            $script:DropIntakeResult = @{ Action = 'SaveApplication'; Values = $v; Analysis = $script:DropActiveAnalysis; ApplicationId = [string]$saved.ApplicationId }
            $dlg.Close()
        }
        catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save as Application' -Message $_.Exception.Message -Buttons OK -Icon Error)
        }
    })
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

        # The ad-hoc stage runs in-process, so the policy has to sit on the
        # process environment for the duration of this run only: leaving it
        # behind would hand the next build a policy nobody selected for it.
        $savedSigning = $env:APP_PACKAGER_SIGNING
        $savedWorkbenchRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
        if ([string]$Ctx.SigningJson) { $env:APP_PACKAGER_SIGNING = [string]$Ctx.SigningJson }
        if ([string]$Ctx.WorkbenchDataRoot) { $env:APP_PACKAGER_WORKBENCH_ROOT = [string]$Ctx.WorkbenchDataRoot }

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
            if ($null -eq $savedSigning) { Remove-Item Env:\APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }
            else { $env:APP_PACKAGER_SIGNING = $savedSigning }
            if ($null -eq $savedWorkbenchRoot) { Remove-Item Env:\APP_PACKAGER_WORKBENCH_ROOT -ErrorAction SilentlyContinue }
            else { $env:APP_PACKAGER_WORKBENCH_ROOT = $savedWorkbenchRoot }
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
            -Message 'Staging a dropped installer needs a Download Root. Open OPTIONS -> ConfigMgr Preferences to configure it.' `
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
        $packageAvailable = $false; $packageReason = 'SiteCode is not configured (OPTIONS -> ConfigMgr Preferences).'
    }
    elseif ([string]::IsNullOrWhiteSpace($script:Prefs.FileShareRoot)) {
        $packageAvailable = $false; $packageReason = 'File Share Root is not configured (OPTIONS -> ConfigMgr Preferences).'
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
        # The dialog hands back the analysis for the install mode it shows.
        if ($choice.Analysis) { $analysis = $choice.Analysis }

        if ($choice.Action -eq 'SaveApplication') {
            Add-LogLine -Message ('Saved as application {0}. It appears in the workbench with manual updates; nothing was staged or deployed.' -f [string]$choice.ApplicationId)
            $answer = Show-ThemedMessage -Owner $window -Title 'Saved' `
                -Message ('"{0}" is saved as an application. Open it in the Application Workbench now?' -f [string]$choice.Values.AppName) `
                -Buttons YesNo -Icon Question
            if ($answer -eq 'Yes') {
                Show-ApplicationWorkbench -Owner $window -PreselectApplicationId ([string]$choice.ApplicationId)
            }
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
        SigningJson          = Get-WorkbenchSigningPolicyJson
        WorkbenchDataRoot    = $(if (Test-WorkbenchModuleAvailable) { Get-WorkbenchDataRoot } else { '' })
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
# Browse fallback for the drop target: a drag from Explorer is silently
# blocked when the processes run at different elevation levels.
$btnAddInstaller.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title  = 'Select installer(s) to package'
    $dlg.Filter = 'Installers (*.msi;*.exe)|*.msi;*.exe'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog()) { Invoke-DropIntake -Paths ([string[]]$dlg.FileNames) }
})

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

# --- 2. Check ConfigMgr ---
$btnCheckMECM.Add_Click({
    if (-not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Check ConfigMgr requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and reboot if you just installed) before running Check ConfigMgr." `
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
        $txtStatus.Text = "Querying ConfigMgr for selected products..."

        foreach ($row in $selectedRows) {
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )

            $app    = [string]$row.Application
            $cmName = [string]$row.CMName

            Add-LogLine -Message ("ConfigMgr: {0}" -f $app)
            $row.Status = "Querying ConfigMgr..."
            $dataGrid.Items.Refresh()

            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -ProviderMachineName $script:Prefs.ProviderMachineName -CMName $cmName

                if (-not $res.Found) {
                    $row.CurrentVersion = ""
                    $row.Status = "Not found in ConfigMgr"
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
                    $row.Status = "ConfigMgr version retrieved"
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

        # Auto-discovery: offer to hide apps not found in ConfigMgr
        if (@($script:Prefs.HiddenApplications).Count -eq 0) {
            $notFound = @()
            foreach ($item in $script:PackagerData) {
                if ([string]$item.Status -eq "Not found in ConfigMgr") {
                    $notFound += [string]$item.Script
                }
            }
            if ($notFound.Count -gt 0 -and $notFound.Count -lt $script:PackagerData.Count) {
                $answer = Show-ThemedMessage -Owner $window -Title "Hide Unused Applications" `
                    -Message ("{0} application(s) were not found in ConfigMgr.`n`nHide them from the grid? You can change this later via Product Filter." -f $notFound.Count) `
                    -Buttons YesNo -Icon Question
                if ($answer -eq 'Yes') {
                    $script:Prefs.HiddenApplications = $notFound
                    Save-Preferences -Prefs $script:Prefs
                    Invoke-RefreshGrid
                    Add-LogLine -Message ("{0} application(s) hidden. Manage via Product Filter." -f $notFound.Count)
                }
            }
        }

        $txtStatus.Text = "ConfigMgr query complete."
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
    $selectedRows = @(Confirm-LocalSourceFolders -Rows $selectedRows)
    if ($selectedRows.Count -eq 0) {
        $txtStatus.Text = "Nothing to stage."
        return
    }

    $txtStatus.Text = "Staging selected packages..."
    Invoke-MultiAppPipeline -Operation Stage -Rows $selectedRows -Context @{
        DownloadRoot   = $dlRootValue
        M365Channel    = $script:Prefs.M365Channel
        M365DeployMode = $script:Prefs.M365DeployMode
        LogFolder      = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath   = Get-SevenZipPathForContext
        RunPlanByApp   = Get-WorkbenchRunPlanForContext -Rows $selectedRows -Target ([string]$script:Prefs.Intune.DeploymentTarget)
        SigningJson    = Get-WorkbenchSigningPolicyJson
        SigningDigest  = Get-WorkbenchSigningPolicyDigest
    }
})

# --- 4. Package Apps ---
$btnPackage.Add_Click({
    # Intune-only runs never touch the site or the share, so the console,
    # SiteCode, and File Share Root gates apply only to ConfigMgr targets.
    $intuneOnlyRun = ([string]$script:Prefs.Intune.DeploymentTarget -eq 'IntuneOnly')
    if (-not $intuneOnlyRun -and -not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Package requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and reboot if you just installed) before packaging." `
            -Buttons OK -Icon Warning)
        return
    }

    $siteCodeValue = $script:Prefs.SiteCode
    if (-not $intuneOnlyRun -and [string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $fsPathValue = $script:Prefs.FileShareRoot
    if (-not $intuneOnlyRun -and [string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -Message "File Share Root is required. Open Preferences to configure."
        $txtStatus.Text = "File Share Root is required."
        return
    }

    if ($intuneOnlyRun -and -not (Get-IntunePublishConfigForContext)) {
        Add-LogLine -Message "Publish requires Intune credentials. Open ConfigMgr Preferences to configure Tenant ID, Client ID, and Client Secret."
        $txtStatus.Text = "Intune credentials required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    $txtStatus.Text = "Packaging selected applications..."
    $rowsForPlan = $selectedRows
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
        InstallModesByApp    = Get-InstallModesMapForContext
        TitleModesByApp      = Get-TitleModesMapForContext
        DefaultTitleMode     = Get-DefaultTitleModeForContext
        IntunePublishConfig  = Get-IntunePublishConfigForContext
        DeploymentTarget     = [string]$script:Prefs.Intune.DeploymentTarget
        RunPlanByApp         = Get-WorkbenchRunPlanForContext -Rows $rowsForPlan -Target ([string]$script:Prefs.Intune.DeploymentTarget)
        SigningJson          = Get-WorkbenchSigningPolicyJson
        SigningDigest        = Get-WorkbenchSigningPolicyDigest
    }
})

# --- 5. Full Run (one-click tracked-apps flow) ---
# Thin dispatch: validates prefs + ConfigMgr availability + tracked set, then
# routes to Invoke-MultiAppPipeline -Operation FullRun. The bg scriptblock
# there mirrors the original per-row cadence / ConfigMgr pre-flight / Stage /
# Package logic so history entries and row.Status flips stay identical.
$btnFullRun.Add_Click({
    # Intune-only runs never touch the site, so the SiteCode and console
    # gates apply only to ConfigMgr targets.
    $intuneOnlyRun = ([string]$script:Prefs.Intune.DeploymentTarget -eq 'IntuneOnly')
    $siteCodeValue = $script:Prefs.SiteCode
    if (-not $intuneOnlyRun -and [string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "SiteCode is required. Open ConfigMgr Preferences to configure."
        $txtStatus.Text = "SiteCode is required."
        return
    }

    $actionPlanned = $script:Prefs.AppFlow.Action
    if (-not $intuneOnlyRun -and $actionPlanned -eq 'StageAndPackage' -and -not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
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
        Add-LogLine -Message ("File Share Root is required for action '{0}'. Open ConfigMgr Preferences." -f $action)
        $txtStatus.Text = "File Share Root is required."
        return
    }
    if ($action -in @('Stage','StageAndPackage') -and [string]::IsNullOrWhiteSpace($dlRootValue)) {
        Add-LogLine -Message ("Download Root is required for action '{0}'. Open ConfigMgr Preferences." -f $action)
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
    if ($action -in @('Stage','StageAndPackage')) {
        $rows = @(Confirm-LocalSourceFolders -Rows $rows)
        if ($rows.Count -eq 0) {
            $txtStatus.Text = "Nothing to run."
            return
        }
    }

    Add-LogSeparator
    Add-LogLine -Message ("One Click: {0} app(s), action={1}{2}" -f $rows.Count, $action, $(if ($forceFlag) { ', force=on' } else { '' }))
    $txtStatus.Text = ("One Click: {0} app(s)..." -f $rows.Count)

    $rowsForPlan = $rows
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
        InstallModesByApp    = Get-InstallModesMapForContext
        TitleModesByApp      = Get-TitleModesMapForContext
        DefaultTitleMode     = Get-DefaultTitleModeForContext
        IntunePublishConfig  = Get-IntunePublishConfigForContext
        DeploymentTarget     = [string]$script:Prefs.Intune.DeploymentTarget
        RunPlanByApp         = Get-WorkbenchRunPlanForContext -Rows $rowsForPlan -Target ([string]$script:Prefs.Intune.DeploymentTarget)
        SigningJson          = Get-WorkbenchSigningPolicyJson
        SigningDigest        = Get-WorkbenchSigningPolicyDigest
    }
})

# =============================================================================
# Application Workbench
# =============================================================================
# The definition and build model lives in AppPackagerWorkbench.psm1; the
# signing service lives in AppPackagerSigning.psm1. Signing calls stay
# guarded because a build host without the signing module must report that
# no script was signed rather than appear to have signed one.

function Get-WorkbenchCommand {
    param([Parameter(Mandatory)][string]$Name)
    return (Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Test-WorkbenchModuleAvailable {
    return [bool](Get-Command -Name 'Get-WorkbenchApplications' -ErrorAction SilentlyContinue)
}

function ConvertTo-WorkbenchUiHashtable {
    # The editor mutates nested members by name, so a profile loaded as a
    # PSCustomObject tree is rebuilt as ordered hashtables.
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $out = [ordered]@{}
        foreach ($k in @($InputObject.Keys)) { $out[[string]$k] = ConvertTo-WorkbenchUiHashtable -InputObject $InputObject[$k] }
        return $out
    }
    if ($InputObject -is [string] -or $InputObject -is [ValueType]) { return $InputObject }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        return @(foreach ($item in $InputObject) { ConvertTo-WorkbenchUiHashtable -InputObject $item })
    }
    if ($InputObject -is [psobject] -and $InputObject.PSObject.Properties.Count -gt 0) {
        $out = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) { $out[$p.Name] = ConvertTo-WorkbenchUiHashtable -InputObject $p.Value }
        return $out
    }
    return $InputObject
}

function Get-WorkbenchWindowStatePath {
    Join-Path $PSScriptRoot 'AppPackager.workbench.windowstate.json'
}

function Get-WorkbenchXamlPath {
    # The override exists so a UI probe can load the markup from a checkout
    # while the functions themselves run outside the shell's file scope.
    if ($script:WorkbenchXamlPathOverride) { return [string]$script:WorkbenchXamlPathOverride }
    return (Join-Path $PSScriptRoot 'WorkbenchWindow.xaml')
}

function Get-WorkbenchCustomScriptRoot {
    # User-authored packagers live outside the replaceable install tree.
    # Discovery and legacy migration have to agree on this path, or a
    # migrated profile lands under an id the picker never reads.
    return (Join-Path (Get-WorkbenchDataRoot) 'scripts')
}

function Get-WorkbenchUiApplications {
    # One list of catalog packagers, user-authored scripts and saved BYO
    # applications, shaped for the picker. Manual sources report manual
    # updates rather than a fictitious latest-version check.
    $result = New-Object System.Collections.ArrayList
    $packagerMeta = @{}
    foreach ($p in @(Get-Packagers -Root $PackagersRoot)) {
        $packagerMeta[[string]$p.FullPath] = $p
    }
    foreach ($a in @(Get-WorkbenchApplications -PackagersRoot $PackagersRoot -CustomScriptRoot (Get-WorkbenchCustomScriptRoot))) {
        $meta = $null
        if ($a.ScriptPath -and $packagerMeta.ContainsKey([string]$a.ScriptPath)) { $meta = $packagerMeta[[string]$a.ScriptPath] }
        $sourceType = switch ([string]$a.Kind) {
            'Catalog' { 'Catalog packager' }
            'Custom'  { 'User-authored packager' }
            default   { 'Bring your own installer' }
        }
        $display = [string]$a.DisplayName
        if ($meta -and [string]$meta.Application) { $display = [string]$meta.Application }
        $publisher = [string]$a.Publisher
        if (-not $publisher -and $meta) { $publisher = [string]$meta.Vendor }
        [void]$result.Add([pscustomobject]@{
            ApplicationId        = [string]$a.ApplicationId
            DisplayName          = $display
            Publisher            = $publisher
            SourceType           = $sourceType
            Kind                 = [string]$a.Kind
            PackagerBase         = $(if ($a.ScriptName) { [System.IO.Path]::GetFileNameWithoutExtension([string]$a.ScriptName) } else { '' })
            ScriptPath           = [string]$a.ScriptPath
            ActiveProfileId      = [string]$a.ActiveProfileId
            Description          = $(if ($meta) { [string]$meta.Description } else { '' })
            LatestVersionText    = $(if ([string]$a.Kind -eq 'Byo') { 'Manual updates' } else { '' })
            SupportsVariants     = $(if ($meta) { @($meta.SupportsVariants) } else { @() })
            SupportsInstallModes = $(if ($meta) { @($meta.SupportsInstallModes) } else { @() })
        })
    }
    return $result
}

function Get-WorkbenchInheritedSettings {
    # Everything below the profile: global defaults plus packager defaults.
    # A value only a completed Stage can resolve is reported as such and
    # never replaced with an invented default.
    param([Parameter(Mandatory)]$Application)

    $needsStaging = '(needs staging)'
    $globals = [pscustomobject]@{
        EstimatedRuntimeMins = 15
        MaximumRuntimeMins   = 30
        Description          = [string]$Application.Description
        TitleMode            = Get-DefaultTitleModeForContext
    }
    try {
        $globals.EstimatedRuntimeMins = [int]$script:Prefs.EstimatedRuntimeMins
        $globals.MaximumRuntimeMins   = [int]$script:Prefs.MaximumRuntimeMins
    } catch { }

    $resolved = $null
    try { $resolved = Resolve-EffectiveSettings -GlobalDefaults $globals -BaseManifest $null -Profile $null } catch { }

    $pick = {
        param([string]$field, $fallback)
        if ($resolved -and $resolved.Contains($field)) {
            $entry = $resolved[$field]
            if ([string]$entry.Source -eq 'NeedsStaging') { return $needsStaging }
            if ($null -ne $entry.Value -and [string]$entry.Value -ne '') { return $entry.Value }
        }
        return $fallback
    }

    return [pscustomobject]@{
        Source             = 'Packager default'
        DisplayName        = $(if ([string]$Application.DisplayName) { [string]$Application.DisplayName } else { & $pick 'DisplayName' $needsStaging })
        Publisher          = $(if ([string]$Application.Publisher) { [string]$Application.Publisher } else { & $pick 'Publisher' $needsStaging })
        Description        = [string](& $pick 'Description' '')
        TitleMode          = $(if (Get-DefaultTitleModeForContext) { 'Include version' } else { 'Packager default' })
        InstallCommand     = [string](& $pick 'InstallCommand' $needsStaging)
        UninstallCommand   = [string](& $pick 'UninstallCommand' $needsStaging)
        DetectionSummary   = [string](& $pick 'Detection' $needsStaging)
        RebootPolicy       = 'Packager default'
        WorkingDirectory   = ''
        EstimatedMinutes   = [int](& $pick 'EstimatedMinutes' $globals.EstimatedRuntimeMins)
        MaximumMinutes     = [int](& $pick 'MaximumMinutes' $globals.MaximumRuntimeMins)
        ExecutionContext   = [string](& $pick 'Context' 'System')
        LogonRequirement   = [string](& $pick 'LogonRequirement' 'Whether or not a user is logged on')
        UserInteraction    = [string](& $pick 'UserInteraction' 'Hidden')
        ScriptHost         = [string](& $pick 'ScriptHost' 'x64')
        NeedsStagingMarker = $needsStaging
    }
}

function Get-WorkbenchOverrideValue {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)][string]$Path)
    $node = $Profile
    foreach ($part in ($Path -split '\.')) {
        if ($node -isnot [System.Collections.IDictionary] -or -not $node.Contains($part)) { return $null }
        $node = $node[$part]
    }
    return $node
}

function Test-WorkbenchOverridePresent {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)][string]$Path)
    $node = $Profile
    foreach ($part in ($Path -split '\.')) {
        if ($node -isnot [System.Collections.IDictionary] -or -not $node.Contains($part)) { return $false }
        $node = $node[$part]
    }
    return $true
}

function Set-WorkbenchOverrideValue {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)][string]$Path, $Value)
    $parts = @($Path -split '\.')
    $node = $Profile
    for ($i = 0; $i -lt $parts.Count - 1; $i++) {
        if (-not $node.Contains($parts[$i]) -or $node[$parts[$i]] -isnot [System.Collections.IDictionary]) {
            $node[$parts[$i]] = [ordered]@{}
        }
        $node = $node[$parts[$i]]
    }
    $node[$parts[-1]] = $Value
}

function Clear-WorkbenchOverrideValue {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)][string]$Path)
    $parts = @($Path -split '\.')
    $node = $Profile
    for ($i = 0; $i -lt $parts.Count - 1; $i++) {
        if ($node -isnot [System.Collections.IDictionary] -or -not $node.Contains($parts[$i])) { return }
        $node = $node[$parts[$i]]
    }
    if ($node -is [System.Collections.IDictionary] -and $node.Contains($parts[-1])) { $node.Remove($parts[-1]) }
}

function Get-WorkbenchFieldSourceText {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)][string]$Path, $InheritedValue)
    $inheritedText = if ($null -eq $InheritedValue -or [string]$InheritedValue -eq '') { '(none)' } else { [string]$InheritedValue }
    if (Test-WorkbenchOverridePresent -Profile $Profile -Path $Path) {
        $stored = Get-WorkbenchOverrideValue -Profile $Profile -Path $Path
        if ($null -eq $stored) { return 'Custom: removed (inherited ' + $inheritedText + ')' }
        return 'Custom (inherited ' + $inheritedText + ')'
    }
    return 'Inherited: ' + $inheritedText
}

function Get-WorkbenchLineNumberText {
    # Gutter content for the script editors: one label per physical line, so
    # the column stays aligned with an unwrapped monospace TextBox.
    param([AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { $Text = '' }
    $count = (@($Text -split "`n")).Count
    if ($count -lt 1) { $count = 1 }
    $sb = New-Object System.Text.StringBuilder
    for ($i = 1; $i -le $count; $i++) { [void]$sb.AppendLine([string]$i) }
    return $sb.ToString().TrimEnd("`r", "`n")
}

function Get-WorkbenchParseDiagnostics {
    # Static inspection only: the text is parsed, never executed.
    param([AllowEmptyString()][string]$Text)
    $results = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Text)) { return $results }
    $tokens = $null
    $errors = $null
    try {
        [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
    } catch {
        $results.Add('Parser failure: ' + $_.Exception.Message)
        return $results
    }
    foreach ($err in @($errors)) {
        $results.Add(('Line {0}, column {1}: {2}' -f $err.Extent.StartLineNumber, $err.Extent.StartColumnNumber, $err.Message))
    }
    return $results
}

function Test-WorkbenchDetectionScriptOutput {
    # Intune reads exit 0 with no STDOUT as "not installed", so a detector
    # that only exits successfully reports every endpoint as missing.
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $tokens = $null
    $errors = $null
    $ast = $null
    try { $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors) } catch { return $false }
    if ($null -eq $ast) { return $false }
    $writers = $ast.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.CommandAst]) -and
        (@('write-output', 'write-host', 'echo') -contains ([string]$n.GetCommandName()).ToLowerInvariant())
    }, $true)
    if (@($writers).Count -gt 0) { return $true }
    # A bare expression statement is STDOUT in PowerShell too.
    $pipelines = $ast.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.PipelineAst]) -and
        ($n.Parent -is [System.Management.Automation.Language.NamedBlockAst])
    }, $true)
    foreach ($p in @($pipelines)) {
        if ($p.PipelineElements.Count -eq 1 -and $p.PipelineElements[0] -is [System.Management.Automation.Language.CommandExpressionAst]) { return $true }
    }
    return $false
}

function Show-WorkbenchUnsavedDialog {
    # Three-way answer; a two-button themed message cannot express Cancel
    # alongside Discard. Returns Save, Discard or Cancel.
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][string]$Message)

    if ($script:WorkbenchUnsavedPromptOverride) {
        return [string](& $script:WorkbenchUnsavedPromptOverride $Message)
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Unsaved Changes" Width="480" SizeToContent="Height" MinWidth="420"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ResizeMode="NoResize" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="18,16,18,14">
        <TextBlock x:Name="txtUnsaved" TextWrapping="Wrap" FontSize="12" Margin="0,0,0,16"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnUnsavedSave" Content="Save" MinWidth="100" Height="30" Margin="0,0,8,0"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnUnsavedDiscard" Content="Discard" MinWidth="100" Height="30" Margin="0,0,8,0"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnUnsavedCancel" Content="Cancel" MinWidth="100" Height="30" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner
    $dlg.FindName('txtUnsaved').Text = $Message
    # ShowDialog keeps this frame alive while the handlers run, so the
    # answer travels through script scope rather than a local.
    $script:WorkbenchUnsavedChoice = 'Cancel'
    $dlg.FindName('btnUnsavedSave').Add_Click({ $script:WorkbenchUnsavedChoice = 'Save'; $dlg.Close() })
    $dlg.FindName('btnUnsavedDiscard').Add_Click({ $script:WorkbenchUnsavedChoice = 'Discard'; $dlg.Close() })
    $dlg.FindName('btnUnsavedCancel').Add_Click({ $script:WorkbenchUnsavedChoice = 'Cancel'; $dlg.Close() })
    [void]$dlg.ShowDialog()
    return [string]$script:WorkbenchUnsavedChoice
}

function Show-WorkbenchNameDialog {
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][string]$Title, [string]$Value = '')

    if ($script:WorkbenchNamePromptOverride) {
        return [string](& $script:WorkbenchNamePromptOverride $Title $Value)
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Profile" Width="440" SizeToContent="Height" MinWidth="380"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ResizeMode="NoResize" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="18,16,18,14">
        <TextBlock x:Name="txtNamePrompt" FontSize="12" Margin="0,0,0,8"/>
        <TextBox x:Name="txtNameValue" FontSize="12" Height="28" Margin="0,0,0,16"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnNameOk" Content="OK" MinWidth="90" Height="30" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnNameCancel" Content="Cancel" MinWidth="90" Height="30" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner
    $dlg.FindName('txtNamePrompt').Text = $Title
    $box = $dlg.FindName('txtNameValue')
    $box.Text = $Value
    $script:WorkbenchNameResult = $null
    $dlg.FindName('btnNameOk').Add_Click({
        $v = ([string]$box.Text).Trim()
        if ($v) { $script:WorkbenchNameResult = $v }
        $dlg.Close()
    })
    $dlg.FindName('btnNameCancel').Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog()
    return $script:WorkbenchNameResult
}

function Get-WorkbenchLocalFindings {
    # Local validation only: it inspects the profile and never contacts a
    # target. The Intune adapter's own findings are merged on top.
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)]$Inherited)

    $findings = New-Object System.Collections.ArrayList
    $add = {
        param($sev, $code, $msg)
        [void]$findings.Add([pscustomobject]@{ Severity = $sev; Code = $code; Message = $msg })
    }

    $est = $null; $max = $null
    if (Test-WorkbenchOverridePresent -Profile $Profile -Path 'Timing.EstimatedMinutes') { $est = Get-WorkbenchOverrideValue -Profile $Profile -Path 'Timing.EstimatedMinutes' }
    if (Test-WorkbenchOverridePresent -Profile $Profile -Path 'Timing.MaximumMinutes')   { $max = Get-WorkbenchOverrideValue -Profile $Profile -Path 'Timing.MaximumMinutes' }
    if ($null -eq $est) { $est = $Inherited.EstimatedMinutes }
    if ($null -eq $max) { $max = $Inherited.MaximumMinutes }
    $estInt = 0; $maxInt = 0
    [void][int]::TryParse([string]$est, [ref]$estInt)
    [void][int]::TryParse([string]$max, [ref]$maxInt)
    if ($estInt -lt 1 -or $maxInt -lt 1) {
        & $add 'Blocking' 'TIMING-RANGE' 'Estimated duration and maximum runtime must both be at least one minute.'
    }
    elseif ($estInt -gt $maxInt) {
        & $add 'Blocking' 'TIMING-ORDER' 'Estimated install duration is greater than the maximum runtime.'
    }
    & $add 'Info' 'TIMING-INTUNE' 'Intune install timeout is not sent by this publisher; the value is recorded in the definition only.'

    foreach ($section in @('Install', 'Uninstall')) {
        if ([string](Get-WorkbenchOverrideValue -Profile $Profile -Path ($section + '.Mode')) -eq 'Custom') {
            & $add 'Review' 'INSTALL-CUSTOM' ($section + ' uses a custom script, so later changes to the packager wrapper no longer reach this profile and need review.')
        }
    }

    if ([string](Get-WorkbenchOverrideValue -Profile $Profile -Path 'Detection.Mode') -eq 'Custom') {
        $rule = Get-WorkbenchOverrideValue -Profile $Profile -Path 'Detection.Rule'
        $type = if ($rule -is [System.Collections.IDictionary] -and $rule.Contains('Type')) { [string]$rule['Type'] } else { '' }
        if ($type -eq 'Script') {
            $text = if ($rule.Contains('ScriptText')) { [string]$rule['ScriptText'] } else { '' }
            if (-not (Test-WorkbenchDetectionScriptOutput -Text $text)) {
                & $add 'Blocking' 'DETECT-NO-OUTPUT' 'The detection script writes nothing to STDOUT, so an installed endpoint still reports not installed.'
            }
            $diag = Get-WorkbenchParseDiagnostics -Text $text
            if (@($diag).Count -gt 0) {
                & $add 'Blocking' 'DETECT-PARSE' ('The detection script does not parse: ' + @($diag)[0])
            }
        }
        $logic = if ($rule -is [System.Collections.IDictionary] -and $rule.Contains('Logic')) { [string]$rule['Logic'] } else { '' }
        $conv = [bool](Get-WorkbenchOverrideValue -Profile $Profile -Path 'Detection.IntuneScriptConversion')
        if ($logic -in @('Or', 'TwoGroup') -and -not $conv) {
            & $add 'Blocking' 'DETECT-INTUNE-COMPOUND' 'Intune cannot express this grouped or OR detection; choose script conversion explicitly, or publish to ConfigMgr only.'
        }
    }

    if (@(Get-WorkbenchOverrideValue -Profile $Profile -Path 'Requirements.Operations').Count -gt 0) {
        & $add 'Review' 'REQ-INTUNE' 'Requirement rules are not translated by the Intune publisher; the ConfigMgr deployment type carries them.'
    }

    if (@(Get-WorkbenchOverrideValue -Profile $Profile -Path 'Variants.Split').Count -gt 0) {
        & $add 'Blocking' 'VARIANT-INTUNE' 'A variant split produces several deployment types; Intune publishing of split applications is not supported.'
    }

    foreach ($f in @($Profile.SourceFiles)) {
        $dest = ''
        if ($f -is [System.Collections.IDictionary] -and $f.Contains('Destination')) { $dest = [string]$f['Destination'] }
        if ([System.IO.Path]::IsPathRooted($dest) -or $dest -match '(^|[\\/])\.\.([\\/]|$)') {
            & $add 'Blocking' 'SOURCE-PATH' ('Source file destination escapes the content root: ' + $dest)
        }
    }

    return $findings
}

function Get-WorkbenchChipState {
    # Three independent results: an Intune finding never changes the ConfigMgr
    # or content-build verdict.
    param([Parameter(Mandatory)]$Findings, [bool]$Validated)

    $state = [ordered]@{ Content = 'Not validated'; Mecm = 'Not validated'; Intune = 'Not validated' }
    if (-not $Validated) { return $state }

    $all = @($Findings)
    $contentIssues = @($all | Where-Object { $_.Severity -in @('Blocking', 'Review') -and ([string]$_.Code -like 'SOURCE-*' -or [string]$_.Code -like 'INSTALL-*') })
    $mecmIssues    = @($all | Where-Object { $_.Severity -in @('Blocking', 'Review') -and [string]$_.Code -notlike '*INTUNE*' })
    $intuneBlock   = @($all | Where-Object { $_.Severity -eq 'Blocking' -and [string]$_.Code -like '*INTUNE*' })
    $intuneReview  = @($all | Where-Object { $_.Severity -eq 'Review' -and [string]$_.Code -like '*INTUNE*' })

    $state.Content = $(if ($contentIssues.Count -gt 0) { 'Needs review' } else { 'Ready' })
    $state.Mecm    = $(if ($mecmIssues.Count -gt 0) { 'Needs review' } else { 'Ready' })
    if ($intuneBlock.Count -gt 0) { $state.Intune = 'Unsupported' }
    elseif ($intuneReview.Count -gt 0) { $state.Intune = 'Needs review' }
    else { $state.Intune = 'Ready' }
    return $state
}

function Get-WorkbenchRunPlanForContext {
    # Prebuilt on the UI thread: the background runspace has its own session
    # state and cannot read the preferences object or the data root.
    param([array]$Rows, [string]$Target = 'MECM')

    $plan = @{}
    if (-not (Test-WorkbenchModuleAvailable)) { return $plan }
    if ([string]$Target -notin @('ContentOnly', 'MECM', 'MECMAndIntune', 'IntuneOnly')) { $Target = 'MECM' }
    $signing = Get-WorkbenchSigningPolicy
    $downloadRoot = [string]$script:Prefs.DownloadRoot
    $dataRoot = Get-WorkbenchDataRoot

    foreach ($row in @($Rows)) {
        $scriptPath = [string]$row.FullPath
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script)
        if (-not $base) { continue }
        $appId = ''
        try { $appId = [string](New-ApplicationId -Kind Catalog -ScriptPath $scriptPath) } catch { continue }
        $definition = Get-ApplicationDefinition -ApplicationId $appId
        $profileId = [string]$definition.ActiveProfileId
        if (-not $profileId) { $profileId = 'default' }

        $revision = 0
        $title = ''
        if ($profileId -ne 'default') {
            try {
                $prof = Get-Profile -ApplicationId $appId -ProfileId $profileId
                $revision = [int]$prof.Revision
                $appBlock = $prof.Application
                if ($appBlock) {
                    $member = $null
                    if ($appBlock -is [System.Collections.IDictionary]) { $member = $appBlock['DisplayName'] }
                    else { $member = $appBlock.DisplayName }
                    if ($member) { $title = [string]$member }
                }
            } catch { }
        }

        $snapshotPath = ''
        $buildId = ''
        try {
            $snapshot = New-RunSnapshot -ApplicationId $appId -ProfileId $profileId -Target $Target `
                -SigningPolicy $signing -PackagerScriptPath $scriptPath `
                -DownloadRoot (Get-WorkbenchProfileDownloadRoot -DownloadRoot $downloadRoot -ProfileId $profileId)
            $snapshotPath = [string]$snapshot.Path
            $buildId = [string]$snapshot.BuildId
        }
        catch {
            Add-LogLine -Message ('Run snapshot not created for {0}: {1}' -f $base, $_.Exception.Message)
        }

        $plan[$base] = @{
            ApplicationId   = $appId
            ProfileId       = $profileId
            ProfileRevision = $revision
            DisplayTitle    = $title
            SnapshotPath    = $snapshotPath
            BuildId         = $buildId
            DataRoot        = $dataRoot
            DownloadRoot    = (Get-WorkbenchProfileDownloadRoot -DownloadRoot $downloadRoot -ProfileId $profileId)
        }
    }
    return $plan
}

function Get-WorkbenchPackagerStageRoot {
    # The subtree the packager child reads through its own staged-version.txt.
    # Scoping the search here keeps one packager's builds from matching
    # another's stage folder under the same download root.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$DownloadRoot, [AllowEmptyString()][string]$PackagerPath)

    if ([string]::IsNullOrWhiteSpace($DownloadRoot)) { return '' }
    if (-not [string]::IsNullOrWhiteSpace($PackagerPath) -and (Test-Path -LiteralPath $PackagerPath)) {
        $info = Get-PackagerFolderInfo -ScriptPath $PackagerPath
        if ($info.DownloadSubfolder) {
            $scoped = Join-Path $DownloadRoot $info.DownloadSubfolder
            if (Test-Path -LiteralPath $scoped) { return $scoped }
        }
    }
    return $DownloadRoot
}

function Assert-WorkbenchBuildSelection {
    <#
        Refuses a Package run whose selected build is not what the child
        would resolve. The packager child reads staged-version.txt itself,
        so verifying the selection exists is not enough: the newest stage
        in the tree the child will read has to be the selected build.
        Returns the resolved manifest record.
    #>
    param(
        [Parameter(Mandatory)][string]$BuildId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$DownloadRoot,
        [AllowEmptyString()][string]$PackagerPath = ''
    )

    if (-not (Get-Command -Name 'Resolve-StageManifestForBuild' -ErrorAction SilentlyContinue)) {
        throw "Build selection needs the workbench definition model; Resolve-StageManifestForBuild is not available."
    }
    $searchRoot = Get-WorkbenchPackagerStageRoot -DownloadRoot $DownloadRoot -PackagerPath $PackagerPath
    if ([string]::IsNullOrWhiteSpace($searchRoot)) {
        throw "stale build: a build id was selected but no download root is configured, so the staged content cannot be identified."
    }

    $selected = Resolve-StageManifestForBuild -BuildId $BuildId -SearchRoot $searchRoot

    # The child picks the newest stage, so anything newer than the selection
    # would be packaged instead of it.
    $newest = @(Get-ChildItem -LiteralPath $searchRoot -Filter 'stage-manifest.json' -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1)
    if ($newest.Count -gt 0) {
        $newestBuildId = ''
        try { $newestBuildId = [string]((Get-Content -LiteralPath $newest[0].FullName -Raw -ErrorAction Stop | ConvertFrom-Json).BuildId) } catch { }
        if ($newestBuildId -and $newestBuildId -ne $BuildId) {
            throw ("stale build: the selected build '{0}' is not the staged content under '{1}'; the newest stage there is build '{2}'. Stage the selected profile again, or select build '{2}'." -f $BuildId, $searchRoot, $newestBuildId)
        }
    }
    return $selected
}

function Test-WorkbenchBuildIsCurrent {
    # One Click freshness: a vendor-version match alone must not skip an
    # application whose profile or signing policy changed since the build.
    # $null means "no build record model available"; the caller then keeps
    # its legacy version-only decision.
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$ProfileId = 'default',
        [AllowEmptyString()][string]$Version,
        [int]$ProfileRevision = 0,
        [AllowEmptyString()][string]$PolicyDigest
    )
    if (-not (Get-Command -Name 'Get-LatestBuildRecord' -ErrorAction SilentlyContinue)) { return $null }
    $record = $null
    try { $record = Get-LatestBuildRecord -ApplicationId $ApplicationId -ProfileId $ProfileId } catch { return $null }
    if (-not $record) { return $false }
    if ([string]$record.SoftwareVersion -ne [string]$Version) { return $false }
    if ([int]$record.ProfileRevision -ne [int]$ProfileRevision) { return $false }
    if ([string]$record.PolicyDigest -ne [string]$PolicyDigest) { return $false }
    return $true
}

function New-WorkbenchFieldDescriptor {
    param(
        [Parameter(Mandatory)]$Control,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$InheritedKey,
        $SourceLabel = $null,
        $ResetButton = $null,
        [ValidateSet('Text', 'Combo', 'Check', 'Int')][string]$Kind = 'Text'
    )
    return [pscustomobject]@{
        Control      = $Control
        Path         = $Path
        InheritedKey = $InheritedKey
        SourceLabel  = $SourceLabel
        ResetButton  = $ResetButton
        Kind         = $Kind
    }
}

function Show-ApplicationWorkbench {
    <#
        Editor for one application definition: identity, commands and
        scripts, detection, requirements and variants, timing and
        execution, extra source files, and the review pane. Opening or
        saving here deploys nothing.
    #>
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$PreselectApplicationId = '',
        [string]$PreselectPackagerBase = '',
        [scriptblock]$Probe
    )

    if (-not (Test-WorkbenchModuleAvailable)) {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Application Workbench' `
            -Message 'The workbench definition model (Packagers\AppPackagerWorkbench.psm1) is not loaded, so no application definition can be read or saved.' `
            -Buttons OK -Icon Error)
        return
    }

    # Legacy per-app preference maps become profiles on first open. The
    # migration is idempotent and leaves the legacy keys in place.
    if (-not $script:WorkbenchMigrationDone) {
        $script:WorkbenchMigrationDone = $true
        try {
            # Both roots travel so a legacy key resolves to the id its script
            # is actually discovered under: catalog:<name> or custom:<name>.
            $migration = Invoke-LegacyPreferenceMigration -Preferences $script:Prefs -PackagersRoot $PackagersRoot -CustomScriptRoot (Get-WorkbenchCustomScriptRoot)
            if ([int]$migration.MigratedCount -gt 0) {
                Add-LogLine -Message ('Migrated {0} application(s) from the legacy preference maps into workbench profiles.' -f [int]$migration.MigratedCount)
            }
        }
        catch {
            Add-LogLine -Message ('Legacy preference migration failed: ' + $_.Exception.Message)
        }
    }

    $xamlPath = Get-WorkbenchXamlPath
    if (-not (Test-Path -LiteralPath $xamlPath)) {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Application Workbench' `
            -Message ('WorkbenchWindow.xaml was not found at ' + $xamlPath + '.') -Buttons OK -Icon Error)
        return
    }
    [xml]$wbXaml = Get-Content -LiteralPath $xamlPath -Raw
    $reader = New-Object System.Xml.XmlNodeReader $wbXaml
    $win = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $win
    Set-DialogChromeFromOwner -Dialog $win -Owner $Owner

    $ctl = {
        param([string]$n)
        $found = $win.FindName($n)
        if ($null -eq $found) { throw ('WorkbenchWindow.xaml is missing the control ' + $n + '.') }
        return $found
    }

    # Header
    $cboApplication = & $ctl 'cboApplication'
    $cboProfile     = & $ctl 'cboProfile'
    $btnProfileRename = & $ctl 'btnProfileRename'
    $txtSourceType  = & $ctl 'txtSourceType'
    $cboBuild       = & $ctl 'cboBuild'
    $lstSections    = & $ctl 'lstSections'

    # Section panels, in the order the section list shows them.
    $sectionNames = @('Application', 'Install & uninstall', 'Detection', 'Requirements & variants', 'Timing & execution', 'Source files', 'Review & build')
    $sectionPanels = @(
        (& $ctl 'pnlApplication'), (& $ctl 'pnlInstall'), (& $ctl 'pnlDetection'),
        (& $ctl 'pnlRequirements'), (& $ctl 'pnlTiming'), (& $ctl 'pnlSources'), (& $ctl 'pnlReview')
    )

    # Application section
    $txtDisplayName = & $ctl 'txtDisplayName'
    $txtPublisher   = & $ctl 'txtPublisher'
    $txtDescription = & $ctl 'txtDescription'
    $cboTitleMode   = & $ctl 'cboTitleMode'
    $txtProfileName = & $ctl 'txtProfileName'
    $txtIdentity    = & $ctl 'txtIdentity'
    $txtProvenance  = & $ctl 'txtProvenance'
    $imgIcon        = & $ctl 'imgIcon'
    $btnIconChoose  = & $ctl 'btnIconChoose'
    $btnIconReset   = & $ctl 'btnIconReset'
    $btnIconRemove  = & $ctl 'btnIconRemove'
    $txtIconState   = & $ctl 'txtIconState'
    $txtAppNote     = & $ctl 'txtAppNote'

    # Install section
    $cboScriptTarget     = & $ctl 'cboScriptTarget'
    $cboInstallMode      = & $ctl 'cboInstallMode'
    $lblInstallModeSrc   = & $ctl 'lblInstallModeSrc'
    $btnInstallReset     = & $ctl 'btnInstallReset'
    $txtEffectiveCommand = & $ctl 'txtEffectiveCommand'
    $txtReturnCodes      = & $ctl 'txtReturnCodes'
    $cboRebootPolicy     = & $ctl 'cboRebootPolicy'
    $txtWorkingDirectory = & $ctl 'txtWorkingDirectory'
    $lblScriptCaption    = & $ctl 'lblScriptCaption'
    $txtScript           = & $ctl 'txtScript'
    $txtScriptGutter     = & $ctl 'txtScriptGutter'
    $txtHookBefore       = & $ctl 'txtHookBefore'
    $txtHookAfter        = & $ctl 'txtHookAfter'
    $chkAfterOnFailure   = & $ctl 'chkAfterOnFailure'
    $txtScriptNote       = & $ctl 'txtScriptNote'
    $lstScriptDiagnostics = & $ctl 'lstScriptDiagnostics'
    $pnlFind             = & $ctl 'pnlFind'
    $txtFind             = & $ctl 'txtFind'
    $btnFindNext         = & $ctl 'btnFindNext'
    $btnFindClose        = & $ctl 'btnFindClose'

    # Detection section
    $cboDetectionMode      = & $ctl 'cboDetectionMode'
    $cboDetectionType      = & $ctl 'cboDetectionType'
    $lblDetectionSrc       = & $ctl 'lblDetectionSrc'
    $btnDetectionReset     = & $ctl 'btnDetectionReset'
    $txtDetectionInherited = & $ctl 'txtDetectionInherited'
    $pnlDetectionTyped     = & $ctl 'pnlDetectionTyped'
    $pnlDetectionScript    = & $ctl 'pnlDetectionScript'
    $cboDetHive            = & $ctl 'cboDetHive'
    $cboDetView            = & $ctl 'cboDetView'
    $txtDetKey             = & $ctl 'txtDetKey'
    $txtDetValueName       = & $ctl 'txtDetValueName'
    $cboDetOperator        = & $ctl 'cboDetOperator'
    $txtDetExpected        = & $ctl 'txtDetExpected'
    $cboDetVersionBinding  = & $ctl 'cboDetVersionBinding'
    $txtDetFilePath        = & $ctl 'txtDetFilePath'
    $txtDetFileName        = & $ctl 'txtDetFileName'
    $cboDetFileProperty    = & $ctl 'cboDetFileProperty'
    $cboDetLogic           = & $ctl 'cboDetLogic'
    $dgDetClauses          = & $ctl 'dgDetClauses'
    $btnDetClauseAdd       = & $ctl 'btnDetClauseAdd'
    $btnDetClauseRemove    = & $ctl 'btnDetClauseRemove'
    $txtIntuneContract     = & $ctl 'txtIntuneContract'
    $txtDetectScript       = & $ctl 'txtDetectScript'
    $txtDetectGutter       = & $ctl 'txtDetectGutter'
    $chkIntuneScriptConversion = & $ctl 'chkIntuneScriptConversion'
    $lstDetectDiagnostics  = & $ctl 'lstDetectDiagnostics'

    # Requirements section
    $cboReqTemplate   = & $ctl 'cboReqTemplate'
    $btnReqAdd        = & $ctl 'btnReqAdd'
    $btnReqReplace    = & $ctl 'btnReqReplace'
    $btnReqRemove     = & $ctl 'btnReqRemove'
    $dgRequirements   = & $ctl 'dgRequirements'
    $cboVariantSplit  = & $ctl 'cboVariantSplit'
    $cboInstallForMode = & $ctl 'cboInstallForMode'
    $txtSiteArchGc      = & $ctl 'txtSiteArchGc'
    $txtSiteLangGc      = & $ctl 'txtSiteLangGc'
    $txtSiteVpnGc       = & $ctl 'txtSiteVpnGc'
    $txtSiteVpnPatterns = & $ctl 'txtSiteVpnPatterns'

    # Site condition names are shared by every application and every
    # profile, so they persist on focus loss instead of with Save.
    $loadSiteConditions = {
        try { $doc = Get-ConditionTemplates } catch { return }
        $byId = @{}
        foreach ($c in @($doc.Conditions)) { $byId[[string]$c.Id] = $c }
        $txtSiteArchGc.Text      = if ($byId['cpu-arch'])      { [string]$byId['cpu-arch'].GlobalConditionName } else { '' }
        $txtSiteLangGc.Text      = if ($byId['os-language'])   { [string]$byId['os-language'].GlobalConditionName } else { '' }
        $txtSiteVpnGc.Text       = if ($byId['vpn-connected']) { [string]$byId['vpn-connected'].GlobalConditionName } else { '' }
        $txtSiteVpnPatterns.Text = if ($byId['vpn-connected']) { (@($byId['vpn-connected'].AdapterPatterns) -join ', ') } else { '' }
    }
    $saveSiteConditions = {
        try {
            $doc = Get-ConditionTemplates
            foreach ($c in @($doc.Conditions)) {
                switch ([string]$c.Id) {
                    'cpu-arch'      { if ($txtSiteArchGc.Text.Trim()) { $c.GlobalConditionName = $txtSiteArchGc.Text.Trim() } }
                    'os-language'   { if ($txtSiteLangGc.Text.Trim()) { $c.GlobalConditionName = $txtSiteLangGc.Text.Trim() } }
                    'vpn-connected' {
                        if ($txtSiteVpnGc.Text.Trim()) { $c.GlobalConditionName = $txtSiteVpnGc.Text.Trim() }
                        $patterns = @([string]$txtSiteVpnPatterns.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        if ($patterns.Count -gt 0) { $c.AdapterPatterns = $patterns }
                    }
                }
            }
            [void](Save-ConditionTemplates -Templates $doc)
            & $loadSiteConditions
        } catch { }
    }
    foreach ($tb in @($txtSiteArchGc, $txtSiteLangGc, $txtSiteVpnGc, $txtSiteVpnPatterns)) { $tb.Add_LostFocus($saveSiteConditions) }
    $lblVariantSrc    = & $ctl 'lblVariantSrc'
    $dgVariants       = & $ctl 'dgVariants'

    # Timing section
    $chkEstimatedDefault  = & $ctl 'chkEstimatedDefault'
    $txtEstimatedMinutes  = & $ctl 'txtEstimatedMinutes'
    $lblEstimatedEffective = & $ctl 'lblEstimatedEffective'
    $chkMaximumDefault    = & $ctl 'chkMaximumDefault'
    $txtMaximumMinutes    = & $ctl 'txtMaximumMinutes'
    $lblMaximumEffective  = & $ctl 'lblMaximumEffective'
    $txtTimingTargets     = & $ctl 'txtTimingTargets'
    $cboExecContext       = & $ctl 'cboExecContext'
    $lblExecContextSrc    = & $ctl 'lblExecContextSrc'
    $cboExecLogon         = & $ctl 'cboExecLogon'
    $cboExecInteraction   = & $ctl 'cboExecInteraction'
    $cboExecScriptHost    = & $ctl 'cboExecScriptHost'
    $btnTimingReset       = & $ctl 'btnTimingReset'

    # Source files section
    $dgSourceFiles      = & $ctl 'dgSourceFiles'
    $btnSourceAddFile   = & $ctl 'btnSourceAddFile'
    $btnSourceAddFolder = & $ctl 'btnSourceAddFolder'
    $btnSourceReplace   = & $ctl 'btnSourceReplace'
    $btnSourceRemove    = & $ctl 'btnSourceRemove'
    $txtSourceNote      = & $ctl 'txtSourceNote'

    # Review section
    $dgDiff          = & $ctl 'dgDiff'
    $dgFindings      = & $ctl 'dgFindings'
    $txtResolvedPlan = & $ctl 'txtResolvedPlan'
    $dgBuilds        = & $ctl 'dgBuilds'

    # Footer
    $chipContent   = & $ctl 'chipContent'
    $chipMecm      = & $ctl 'chipMecm'
    $chipIntune    = & $ctl 'chipIntune'
    $txtSaveState  = & $ctl 'txtSaveState'
    $btnUseDefaults = & $ctl 'btnUseDefaults'
    $btnSave       = & $ctl 'btnSave'
    $btnSaveAs     = & $ctl 'btnSaveAs'
    $btnValidate   = & $ctl 'btnValidate'
    $btnStage      = & $ctl 'btnStage'
    $btnPackage    = & $ctl 'btnPackage'

    # Static option lists.
    foreach ($n in $sectionNames) { [void]$lstSections.Items.Add($n) }
    foreach ($v in @('Packager default', 'Include version', 'No version')) { [void]$cboTitleMode.Items.Add($v) }
    foreach ($v in @('Install', 'Uninstall')) { [void]$cboScriptTarget.Items.Add($v) }
    foreach ($v in @('Generated', 'Extend generated', 'Custom')) { [void]$cboInstallMode.Items.Add($v) }
    foreach ($v in @('Inherit', 'No reboot', 'Soft reboot', 'Hard reboot', 'Force restart')) { [void]$cboRebootPolicy.Items.Add($v) }
    foreach ($v in @('Inherit', 'Custom')) { [void]$cboDetectionMode.Items.Add($v) }
    foreach ($v in @('Registry', 'File', 'PowerShell', 'Compound')) { [void]$cboDetectionType.Items.Add($v) }
    foreach ($v in @('HKLM', 'HKCU', 'HKCR')) { [void]$cboDetHive.Items.Add($v) }
    foreach ($v in @('64-bit', '32-bit')) { [void]$cboDetView.Items.Add($v) }
    foreach ($v in @('Exists', 'Equals', 'NotEquals', 'GreaterEqual', 'Greater', 'LessEqual', 'Less', 'Contains')) { [void]$cboDetOperator.Items.Add($v) }
    foreach ($v in @('Follows staged version', 'Pinned')) { [void]$cboDetVersionBinding.Items.Add($v) }
    foreach ($v in @('Exists', 'Version', 'Size', 'DateModified')) { [void]$cboDetFileProperty.Items.Add($v) }
    foreach ($v in @('And', 'Or', 'Two groups')) { [void]$cboDetLogic.Items.Add($v) }
    foreach ($v in @('System', 'User')) { [void]$cboExecContext.Items.Add($v) }
    foreach ($v in @('Whether or not a user is logged on', 'Only when a user is logged on', 'Only when no user is logged on')) { [void]$cboExecLogon.Items.Add($v) }
    foreach ($v in @('Hidden', 'Normal', 'Minimized', 'Maximized')) { [void]$cboExecInteraction.Items.Add($v) }
    foreach ($v in @('x64', 'x86')) { [void]$cboExecScriptHost.Items.Add($v) }
    $txtIntuneContract.Text = 'Intune reads the detection result from the script: exit code 0 plus output on STDOUT means installed. Any output on STDERR is a negative result, and exit 0 with no output means not installed. A detector that only exits successfully is rejected here.'
    $txtSourceNote.Text = 'Bundling a file places it inside the package content. Copying it onto the endpoint is the install script''s job; a detector must not depend on package cache content.'
    $txtTimingTargets.Text = 'ConfigMgr receives both values on the deployment type. The Intune publisher sends neither; the values stay in the definition and the review pane reports the gap.'

    # Session state. Handlers read and write it rather than closing over
    # a dozen separate variables.
    $wb = @{
        Applications   = @()
        Application    = $null
        Inherited      = $null
        Profile        = $null
        ProfileId      = 'default'
        Baseline       = ''
        Dirty          = $false
        Loading        = $true
        ScriptTarget   = 'Install'
        Validated      = $false
        Findings       = @()
        Suppress       = $false
    }

    $sourceFileRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $requirementRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $clauseRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $variantRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $dgSourceFiles.ItemsSource = $sourceFileRows
    $dgRequirements.ItemsSource = $requirementRows
    $dgDetClauses.ItemsSource = $clauseRows
    $dgVariants.ItemsSource = $variantRows

    # Debounced parse: the parser runs on the UI thread but only after the
    # user stops typing, so a large script never stalls each keystroke.
    $parseTimer = New-Object System.Windows.Threading.DispatcherTimer
    $parseTimer.Interval = [TimeSpan]::FromMilliseconds(500)

    $fieldDescriptors = @(
        (New-WorkbenchFieldDescriptor -Control $txtDisplayName -Path 'Application.DisplayName' -InheritedKey 'DisplayName' -SourceLabel (& $ctl 'lblDisplayNameSrc') -ResetButton (& $ctl 'btnDisplayNameReset')),
        (New-WorkbenchFieldDescriptor -Control $txtPublisher   -Path 'Application.Publisher'   -InheritedKey 'Publisher'   -SourceLabel (& $ctl 'lblPublisherSrc')   -ResetButton (& $ctl 'btnPublisherReset')),
        (New-WorkbenchFieldDescriptor -Control $txtDescription -Path 'Application.Description' -InheritedKey 'Description' -SourceLabel (& $ctl 'lblDescriptionSrc') -ResetButton (& $ctl 'btnDescriptionReset')),
        (New-WorkbenchFieldDescriptor -Control $cboTitleMode   -Path 'Application.TitleMode'   -InheritedKey 'TitleMode'   -SourceLabel (& $ctl 'lblTitleModeSrc')   -ResetButton (& $ctl 'btnTitleModeReset') -Kind 'Combo')
    )

    $setStatus = {
        param([string]$text)
        $txtSaveState.Text = $text
    }

    $markDirty = {
        if ($wb.Loading -or $wb.Suppress) { return }
        $wb.Dirty = $true
        $wb.Validated = $false
        $txtSaveState.Text = 'Unsaved changes. Save writes the profile; nothing is deployed.'
        $chipContent.Text = 'Content build: Not validated'
        $chipMecm.Text = 'ConfigMgr: Not validated'
        $chipIntune.Text = 'Intune: Not validated'
        if ($wb.ProfileId -ne 'default') {
            try { [void](Save-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId -Draft ([pscustomobject]$wb.Profile)) } catch { }
        }
    }

    $refreshFieldIndicators = {
        foreach ($d in $fieldDescriptors) {
            if (-not $d.SourceLabel) { continue }
            $inheritedValue = $null
            try { $inheritedValue = $wb.Inherited.($d.InheritedKey) } catch { }
            $d.SourceLabel.Text = Get-WorkbenchFieldSourceText -Profile $wb.Profile -Path $d.Path -InheritedValue $inheritedValue
        }
        $lblInstallModeSrc.Text = Get-WorkbenchFieldSourceText -Profile $wb.Profile -Path ($wb.ScriptTarget + '.Mode') -InheritedValue 'Generated'
        $lblDetectionSrc.Text = Get-WorkbenchFieldSourceText -Profile $wb.Profile -Path 'Detection.Mode' -InheritedValue 'Inherit'
        $lblExecContextSrc.Text = Get-WorkbenchFieldSourceText -Profile $wb.Profile -Path 'Execution.Context' -InheritedValue $wb.Inherited.ExecutionContext
        $lblVariantSrc.Text = Get-WorkbenchFieldSourceText -Profile $wb.Profile -Path 'InstallMode' -InheritedValue 'Packager default'
        $lblEstimatedEffective.Text = 'Default: ' + [string]$wb.Inherited.EstimatedMinutes + ' minutes'
        $lblMaximumEffective.Text = 'Default: ' + [string]$wb.Inherited.MaximumMinutes + ' minutes'
    }

    $refreshScriptEditor = {
        $target = $wb.ScriptTarget
        $lblScriptCaption.Text = $target + ' script'
        $wb.Suppress = $true
        try {
            $mode = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.Mode'))
            if (-not $mode) { $mode = 'Generated' }
            $display = switch ($mode) { 'Extend' { 'Extend generated' } 'Custom' { 'Custom' } default { 'Generated' } }
            $cboInstallMode.SelectedItem = $display
            $txtScript.Text = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.ScriptText'))
            $txtHookBefore.Text = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.BeforeText'))
            $txtHookAfter.Text = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.AfterText'))
            $chkAfterOnFailure.IsChecked = [bool](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.AfterRunsOnFailure'))
            $txtReturnCodes.Text = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.ReturnCodes'))
            $txtWorkingDirectory.Text = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.WorkingDirectory'))
            $reboot = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.RebootPolicy'))
            $cboRebootPolicy.SelectedItem = $(if ($reboot) { $reboot } else { 'Inherit' })

            $custom = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.Command'))
            if ($custom) { $txtEffectiveCommand.Text = $custom }
            elseif ($target -eq 'Install') { $txtEffectiveCommand.Text = [string]$wb.Inherited.InstallCommand }
            else { $txtEffectiveCommand.Text = [string]$wb.Inherited.UninstallCommand }

            $scriptEditable = ($mode -ne 'Generated')
            $txtScript.IsReadOnly = -not $scriptEditable
            $txtHookBefore.IsEnabled = ($mode -eq 'Extend')
            $txtHookAfter.IsEnabled = ($mode -eq 'Extend')
            $chkAfterOnFailure.IsEnabled = ($mode -eq 'Extend')
            $txtScriptNote.Text = switch ($mode) {
                'Extend' { 'Extend generated runs the before hook, the generated wrapper, and the after hook as separate processes with explicit exit-code propagation, so an exit in one step cannot skip the rest.' }
                'Custom' { 'Custom replaces the generated wrapper. Changes to the packager''s own wrapper at a later release no longer reach this profile and need review.' }
                default  { 'Generated keeps the packager''s wrapper. Switch to Extend generated for hooks, or Custom to take ownership of the whole script.' }
            }
            $txtScriptGutter.Text = Get-WorkbenchLineNumberText -Text $txtScript.Text
        }
        finally { $wb.Suppress = $false }
    }

    $refreshDetectionEditor = {
        $wb.Suppress = $true
        try {
            $mode = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Mode')
            if (-not $mode) { $mode = 'Inherit' }
            $cboDetectionMode.SelectedItem = $mode
            $rule = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Rule'
            $get = {
                param([string]$name, $fallback)
                if ($rule -is [System.Collections.IDictionary] -and $rule.Contains($name) -and $null -ne $rule[$name]) { return $rule[$name] }
                return $fallback
            }
            $type = [string](& $get 'Type' 'Registry')
            $cboDetectionType.SelectedItem = $(if ($type -eq 'Script') { 'PowerShell' } else { $type })
            $cboDetHive.SelectedItem = [string](& $get 'Hive' 'HKLM')
            $cboDetView.SelectedItem = $(if ([string](& $get 'View' '64') -eq '32') { '32-bit' } else { '64-bit' })
            $txtDetKey.Text = [string](& $get 'Key' '')
            $txtDetValueName.Text = [string](& $get 'ValueName' '')
            $cboDetOperator.SelectedItem = [string](& $get 'Operator' 'Exists')
            $txtDetExpected.Text = [string](& $get 'ExpectedValue' '')
            $cboDetVersionBinding.SelectedItem = $(if ([string](& $get 'VersionBinding' 'FollowsStaged') -eq 'Pinned') { 'Pinned' } else { 'Follows staged version' })
            $txtDetFilePath.Text = [string](& $get 'FilePath' '')
            $txtDetFileName.Text = [string](& $get 'FileName' '')
            $cboDetFileProperty.SelectedItem = [string](& $get 'FileProperty' 'Exists')
            $logic = [string](& $get 'Logic' 'And')
            $cboDetLogic.SelectedItem = $(if ($logic -eq 'TwoGroup') { 'Two groups' } elseif ($logic -eq 'Or') { 'Or' } else { 'And' })
            $txtDetectScript.Text = [string](& $get 'ScriptText' '')
            $chkIntuneScriptConversion.IsChecked = [bool](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.IntuneScriptConversion')
            $txtDetectGutter.Text = Get-WorkbenchLineNumberText -Text $txtDetectScript.Text

            $clauseRows.Clear()
            foreach ($c in @(& $get 'Clauses' @())) {
                if ($c -isnot [System.Collections.IDictionary]) { continue }
                $clauseRows.Add([pscustomobject]@{
                    Group    = [string]$c['Group']
                    Type     = [string]$c['Type']
                    Subject  = [string]$c['Subject']
                    Operator = [string]$c['Operator']
                    Expected = [string]$c['Expected']
                })
            }

            $isScript = ([string]$cboDetectionType.SelectedItem -eq 'PowerShell')
            $pnlDetectionScript.Visibility = $(if ($isScript) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
            $pnlDetectionTyped.Visibility  = $(if ($isScript) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible })
            $enabled = ($mode -eq 'Custom')
            foreach ($c in @($cboDetectionType, $cboDetHive, $cboDetView, $txtDetKey, $txtDetValueName, $cboDetOperator,
                             $txtDetExpected, $cboDetVersionBinding, $txtDetFilePath, $txtDetFileName, $cboDetFileProperty,
                             $cboDetLogic, $dgDetClauses, $btnDetClauseAdd, $btnDetClauseRemove, $txtDetectScript,
                             $chkIntuneScriptConversion)) {
                $c.IsEnabled = $enabled
            }
            $txtDetectionInherited.Text = 'Inherited rule: ' + [string]$wb.Inherited.DetectionSummary +
                '. Detection establishes installed state; eligibility rules belong in Requirements & variants.'
        }
        finally { $wb.Suppress = $false }
    }

    $refreshSourceFiles = {
        $sourceFileRows.Clear()
        foreach ($f in @($wb.Profile.SourceFiles)) {
            if ($f -isnot [System.Collections.IDictionary]) { continue }
            $size = 0
            [void][int64]::TryParse([string]$f['Size'], [ref]$size)
            $hash = [string]$f['Sha256']
            $sourceFileRows.Add([pscustomobject]@{
                Asset       = [string]$f['Asset']
                Provenance  = [string]$f['Provenance']
                Destination = [string]$f['Destination']
                SizeText    = ('{0:N0} bytes' -f $size)
                ShortHash   = $(if ($hash.Length -ge 16) { $hash.Substring(0, 16) } else { $hash })
                Sha256      = $hash
                Size        = $size
                Linked      = [bool]$f['Linked']
            })
        }
    }

    $refreshRequirements = {
        $requirementRows.Clear()
        foreach ($op in @(Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Requirements.Operations')) {
            if ($op -isnot [System.Collections.IDictionary]) { continue }
            $rule = $op['Rule']
            $summary = ''
            $value = ''
            if ($rule -is [System.Collections.IDictionary]) {
                $summary = [string]$rule['ConditionId']
                if ($rule.Contains('Value')) { $value = [string]$rule['Value'] }
                elseif ($rule.Contains('Cultures')) { $value = (@($rule['Cultures']) -join ', ') }
            }
            $requirementRows.Add([pscustomobject]@{
                Op        = [string]$op['Op']
                RuleId    = [string]$op['RuleId']
                Summary   = $summary
                Value     = $value
                AppliesTo = (@($op['AppliesTo']) -join ', ')
            })
        }
    }

    $refreshVariants = {
        $variantRows.Clear()
        $overrides = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Variants.Overrides'
        if ($overrides -is [System.Collections.IDictionary]) {
            foreach ($suffix in @($overrides.Keys)) {
                $o = $overrides[$suffix]
                $variantRows.Add([pscustomobject]@{
                    Suffix         = [string]$suffix
                    InstallCommand = $(if ($o -is [System.Collections.IDictionary]) { [string]$o['InstallCommand'] } else { '' })
                    Detection      = $(if ($o -is [System.Collections.IDictionary]) { [string]$o['Detection'] } else { '' })
                    Estimated      = $(if ($o -is [System.Collections.IDictionary]) { [string]$o['EstimatedMinutes'] } else { '' })
                    Maximum        = $(if ($o -is [System.Collections.IDictionary]) { [string]$o['MaximumMinutes'] } else { '' })
                })
            }
        }
    }

    $refreshBuilds = {
        $cboBuild.Items.Clear()
        $records = @()
        try { $records = @(Get-BuildRecords -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId) } catch { }
        $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
        if (@($records).Count -eq 0) {
            [void]$cboBuild.Items.Add('No sealed builds')
            $cboBuild.SelectedIndex = 0
        }
        else {
            foreach ($r in @($records)) {
                [void]$cboBuild.Items.Add([string]$r.BuildId)
                $rows.Add([pscustomobject]@{
                    BuildId  = [string]$r.BuildId
                    Version  = [string]$r.Version
                    Revision = [string]$r.ProfileRevision
                    Result   = [string]$r.Result
                })
            }
            $cboBuild.SelectedIndex = 0
        }
        $dgBuilds.ItemsSource = $rows
    }

    $refreshReview = {
        $diff = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
        foreach ($d in $fieldDescriptors) {
            if (-not (Test-WorkbenchOverridePresent -Profile $wb.Profile -Path $d.Path)) { continue }
            $inheritedValue = ''
            try { $inheritedValue = [string]$wb.Inherited.($d.InheritedKey) } catch { }
            $custom = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path
            $diff.Add([pscustomobject]@{
                Field   = $d.Path
                Default = $inheritedValue
                Custom  = $(if ($null -eq $custom) { '(removed)' } else { [string]$custom })
            })
        }
        foreach ($path in @('Install.Mode', 'Install.Command', 'Uninstall.Mode', 'Uninstall.Command',
                            'Detection.Mode', 'InstallMode', 'Timing.EstimatedMinutes', 'Timing.MaximumMinutes',
                            'Execution.Context', 'Execution.ScriptHost')) {
            if (-not (Test-WorkbenchOverridePresent -Profile $wb.Profile -Path $path)) { continue }
            $custom = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path $path
            $default = switch ($path) {
                'Timing.EstimatedMinutes' { [string]$wb.Inherited.EstimatedMinutes }
                'Timing.MaximumMinutes'   { [string]$wb.Inherited.MaximumMinutes }
                'Execution.Context'       { [string]$wb.Inherited.ExecutionContext }
                'Execution.ScriptHost'    { [string]$wb.Inherited.ScriptHost }
                'Install.Command'         { [string]$wb.Inherited.InstallCommand }
                'Uninstall.Command'       { [string]$wb.Inherited.UninstallCommand }
                default                   { 'Inherit' }
            }
            $diff.Add([pscustomobject]@{
                Field   = $path
                Default = $default
                Custom  = $(if ($null -eq $custom) { '(removed)' } else { [string]$custom })
            })
        }
        if (@($wb.Profile.SourceFiles).Count -gt 0) {
            $diff.Add([pscustomobject]@{ Field = 'SourceFiles'; Default = '(none)'; Custom = ('{0} file(s)' -f @($wb.Profile.SourceFiles).Count) })
        }
        $dgDiff.ItemsSource = $diff

        $plan = New-Object System.Text.StringBuilder
        [void]$plan.AppendLine('Application : ' + [string]$wb.Application.ApplicationId)
        [void]$plan.AppendLine('Profile     : ' + [string]$wb.Profile.Name + ' (' + $wb.ProfileId + ', revision ' + [string]$wb.Profile.Revision + ')')
        [void]$plan.AppendLine('Title        : ' + $(if ([string]$txtDisplayName.Text) { [string]$txtDisplayName.Text } else { [string]$wb.Inherited.DisplayName }))
        [void]$plan.AppendLine('Install      : ' + [string]$txtEffectiveCommand.Text)
        [void]$plan.AppendLine('Detection    : ' + $(if ([string]$cboDetectionMode.SelectedItem -eq 'Custom') { 'Custom ' + [string]$cboDetectionType.SelectedItem } else { [string]$wb.Inherited.DetectionSummary }))
        [void]$plan.AppendLine('Timing       : estimated ' + [string]$txtEstimatedMinutes.Text + ' min, maximum ' + [string]$txtMaximumMinutes.Text + ' min')
        [void]$plan.AppendLine('Execution    : ' + [string]$cboExecContext.SelectedItem + ', script host ' + [string]$cboExecScriptHost.SelectedItem)
        [void]$plan.AppendLine('Source files : ' + [string]@($wb.Profile.SourceFiles).Count)
        $signCmd = Get-WorkbenchCommand -Name 'Get-SigningPolicy'
        if ($signCmd) { [void]$plan.AppendLine('Signing      : policy from preferences, applied at build time') }
        else { [void]$plan.AppendLine('Signing      : signing service not loaded; no script will be signed in this build') }
        $txtResolvedPlan.Text = $plan.ToString()
    }

    $refreshEditors = {
        & $refreshFieldIndicators
        & $refreshScriptEditor
        & $refreshDetectionEditor
        & $refreshSourceFiles
        & $refreshRequirements
        & $refreshVariants
        & $refreshReview
    }

    $loadProfileIntoUi = {
        $wb.Loading = $true
        try {
            $wb.Inherited = Get-WorkbenchInheritedSettings -Application $wb.Application
            $txtSourceType.Text = [string]$wb.Application.SourceType
            $txtIdentity.Text = 'Application id ' + [string]$wb.Application.ApplicationId +
                '; output folders and vendor discovery keep using the packager identity, not the displayed title.'
            $prov = [string]$wb.Application.ScriptPath
            if (-not $prov) { $prov = 'Manual updates; no vendor discovery source is configured for this application.' }
            $txtProvenance.Text = $prov
            $txtProfileName.Text = [string]$wb.Profile.Name

            foreach ($d in $fieldDescriptors) {
                $stored = $null
                $present = Test-WorkbenchOverridePresent -Profile $wb.Profile -Path $d.Path
                if ($present) { $stored = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path }
                if ($d.Kind -eq 'Combo') {
                    $display = 'Packager default'
                    if ($present -and $stored) {
                        $display = switch ([string]$stored) { 'IncludeVersion' { 'Include version' } 'NoVersion' { 'No version' } default { [string]$stored } }
                    }
                    $d.Control.SelectedItem = $display
                }
                else {
                    $d.Control.Text = [string]$stored
                }
            }

            $estPresent = Test-WorkbenchOverridePresent -Profile $wb.Profile -Path 'Timing.EstimatedMinutes'
            $maxPresent = Test-WorkbenchOverridePresent -Profile $wb.Profile -Path 'Timing.MaximumMinutes'
            $chkEstimatedDefault.IsChecked = -not $estPresent
            $chkMaximumDefault.IsChecked = -not $maxPresent
            $txtEstimatedMinutes.Text = $(if ($estPresent) { [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.EstimatedMinutes') } else { [string]$wb.Inherited.EstimatedMinutes })
            $txtMaximumMinutes.Text = $(if ($maxPresent) { [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.MaximumMinutes') } else { [string]$wb.Inherited.MaximumMinutes })
            $txtEstimatedMinutes.IsEnabled = $estPresent
            $txtMaximumMinutes.IsEnabled = $maxPresent

            $ctx = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Execution.Context')
            $cboExecContext.SelectedItem = $(if ($ctx) { $ctx } else { [string]$wb.Inherited.ExecutionContext })
            $logon = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Execution.LogonRequirement')
            $cboExecLogon.SelectedItem = $(if ($logon) { $logon } else { [string]$wb.Inherited.LogonRequirement })
            $inter = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Execution.UserInteraction')
            $cboExecInteraction.SelectedItem = $(if ($inter) { $inter } else { [string]$wb.Inherited.UserInteraction })
            $host32 = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Execution.ScriptHost')
            $cboExecScriptHost.SelectedItem = $(if ($host32) { $host32 } else { [string]$wb.Inherited.ScriptHost })

            $cboVariantSplit.Items.Clear()
            [void]$cboVariantSplit.Items.Add('None')
            foreach ($v in @($wb.Application.SupportsVariants)) { [void]$cboVariantSplit.Items.Add([string]$v) }
            $split = @(Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Variants.Split')
            $cboVariantSplit.SelectedItem = $(if ($split.Count -gt 0 -and $split[0]) { [string]$split[0] } else { 'None' })
            $cboVariantSplit.IsEnabled = (@($wb.Application.SupportsVariants).Count -gt 0)

            $cboInstallForMode.Items.Clear()
            [void]$cboInstallForMode.Items.Add('Packager default')
            if (@($wb.Application.SupportsInstallModes) -contains 'AllUsers') { [void]$cboInstallForMode.Items.Add('System') }
            if (@($wb.Application.SupportsInstallModes) -contains 'CurrentUser') { [void]$cboInstallForMode.Items.Add('User') }
            $im = [string](Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'InstallMode')
            $cboInstallForMode.SelectedItem = $(switch ($im) { 'AllUsers' { 'System' } 'CurrentUser' { 'User' } default { 'Packager default' } })
            $cboInstallForMode.IsEnabled = (@($wb.Application.SupportsInstallModes).Count -gt 0)

            $iconAsset = $null
            $iconPresent = Test-WorkbenchOverridePresent -Profile $wb.Profile -Path 'Application.Icon'
            if ($iconPresent) { $iconAsset = Get-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Application.Icon' }
            $imgIcon.Source = $null
            $previewPath = ''
            if (-not $iconPresent) {
                $downloadRoot = ''
                try { $downloadRoot = [string]$script:Prefs.DownloadRoot } catch { }
                $previewPath = Get-WorkbenchInheritedIconPath -ScriptPath ([string]$wb.Application.ScriptPath) -DownloadRoot $downloadRoot
                $txtIconState.Text = $(if ($previewPath) { 'Inherited from the packager: ' + $previewPath } else { 'Inherited from the packager. No staged icon or icon pack entry was found.' })
            }
            elseif ($null -eq $iconAsset) { $txtIconState.Text = 'Removed: this profile publishes no icon.' }
            else {
                if ($iconAsset -is [System.Collections.IDictionary]) { $previewPath = [string]$iconAsset['Path'] }
                $txtIconState.Text = 'Custom icon: ' + $previewPath
            }
            if ($previewPath -and (Test-Path -LiteralPath $previewPath)) {
                try {
                    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmp.BeginInit()
                    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bmp.UriSource = New-Object System.Uri($previewPath)
                    $bmp.EndInit()
                    $imgIcon.Source = $bmp
                } catch { }
            }
            $txtAppNote.Text = 'The displayed title is separate from the stable identity: renaming a profile never renames an existing site application, and the output folder keeps the packager name.'

            $cboReqTemplate.Items.Clear()
            try {
                foreach ($c in @((Get-ConditionTemplates).Conditions)) { [void]$cboReqTemplate.Items.Add([string]$c.Id) }
            } catch { }
            if ($cboReqTemplate.Items.Count -gt 0) { $cboReqTemplate.SelectedIndex = 0 }
            & $loadSiteConditions

            & $refreshEditors
            & $refreshBuilds
            $wb.Baseline = ($wb.Profile | ConvertTo-Json -Depth 12 -Compress)
            $wb.Dirty = $false
            $wb.Validated = $false
            $chipContent.Text = 'Content build: Not validated'
            $chipMecm.Text = 'ConfigMgr: Not validated'
            $chipIntune.Text = 'Intune: Not validated'
            & $setStatus 'No unsaved changes.'
        }
        finally { $wb.Loading = $false }
    }

    $loadProfileList = {
        param([string]$selectProfileId)
        $wb.Loading = $true
        try {
            $cboProfile.Items.Clear()
            $script:WorkbenchProfileIds = @()
            foreach ($p in @(Get-Profiles -ApplicationId ([string]$wb.Application.ApplicationId))) {
                [void]$cboProfile.Items.Add([string]$p.Name)
                $script:WorkbenchProfileIds += [string]$p.ProfileId
            }
            [void]$cboProfile.Items.Add('Save as...')
            $idx = [array]::IndexOf($script:WorkbenchProfileIds, $selectProfileId)
            if ($idx -lt 0) { $idx = 0 }
            $cboProfile.SelectedIndex = $idx
        }
        finally { $wb.Loading = $false }
    }

    $selectApplication = {
        param($application, [string]$profileId)
        $wb.Application = $application
        if (-not $profileId) {
            $definition = Get-ApplicationDefinition -ApplicationId ([string]$application.ApplicationId)
            $profileId = [string]$definition.ActiveProfileId
            if (-not $profileId) { $profileId = 'default' }
        }
        $wb.ProfileId = $profileId
        $stored = $null
        try { $stored = ConvertTo-WorkbenchUiHashtable -InputObject (Get-Profile -ApplicationId ([string]$application.ApplicationId) -ProfileId $profileId) }
        catch { $stored = $null }
        if (-not $stored) {
            $wb.ProfileId = 'default'
            $stored = ConvertTo-WorkbenchUiHashtable -InputObject (New-WorkbenchProfileObject -ApplicationId ([string]$application.ApplicationId) -ProfileId 'default' -Name 'Packager default')
        }
        $wb.Profile = $stored
        if ($wb.ProfileId -ne 'default') {
            $draft = $null
            try { $draft = Get-Draft -ApplicationId ([string]$application.ApplicationId) -ProfileId $wb.ProfileId } catch { }
            if ($draft) {
                $answer = Show-ThemedMessage -Owner $win -Title 'Recover Draft' `
                    -Message ('Unsaved changes to "' + [string]$stored.Name + '" were found from an earlier session. Recover them?') `
                    -Buttons YesNo -Icon Question
                if ($answer -eq 'Yes') { $wb.Profile = ConvertTo-WorkbenchUiHashtable -InputObject $draft }
                else { [void](Remove-Draft -ApplicationId ([string]$application.ApplicationId) -ProfileId $wb.ProfileId) }
            }
        }
        & $loadProfileList $wb.ProfileId
        & $loadProfileIntoUi
    }

    $commitUi = {
        # Reads every editable control back into the working profile.
        # Empty text clears the override so the field returns to inherit;
        # an explicit removal is set by the Remove buttons, not by blanking.
        foreach ($d in $fieldDescriptors) {
            $value = if ($d.Kind -eq 'Combo') { [string]$d.Control.SelectedItem } else { [string]$d.Control.Text }
            if ($d.Kind -eq 'Combo') {
                $value = switch ($value) { 'Packager default' { '' } 'Include version' { 'IncludeVersion' } 'No version' { 'NoVersion' } default { $value } }
            }
            if ([string]::IsNullOrWhiteSpace($value)) {
                if (-not (Test-WorkbenchOverridePresent -Profile $wb.Profile -Path $d.Path) -or
                    $null -ne (Get-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path)) {
                    Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path
                }
            }
            else { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path -Value $value }
        }
        $wb.Profile.Name = ([string]$txtProfileName.Text).Trim()
        if (-not $wb.Profile.Name) { $wb.Profile.Name = 'Custom' }

        $target = $wb.ScriptTarget
        $mode = switch ([string]$cboInstallMode.SelectedItem) { 'Extend generated' { 'Extend' } 'Custom' { 'Custom' } default { 'Generated' } }
        if ($mode -eq 'Generated') { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.Mode') }
        else { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.Mode') -Value $mode }
        foreach ($pair in @(@{ P = 'ScriptText'; V = [string]$txtScript.Text }, @{ P = 'BeforeText'; V = [string]$txtHookBefore.Text },
                            @{ P = 'AfterText'; V = [string]$txtHookAfter.Text }, @{ P = 'ReturnCodes'; V = [string]$txtReturnCodes.Text },
                            @{ P = 'WorkingDirectory'; V = [string]$txtWorkingDirectory.Text })) {
            if ([string]::IsNullOrWhiteSpace($pair.V)) { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.' + $pair.P) }
            else { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.' + $pair.P) -Value $pair.V }
        }
        if ($chkAfterOnFailure.IsChecked -eq $true) { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.AfterRunsOnFailure') -Value $true }
        else { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.AfterRunsOnFailure') }
        $reboot = [string]$cboRebootPolicy.SelectedItem
        if ($reboot -and $reboot -ne 'Inherit') { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.RebootPolicy') -Value $reboot }
        else { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path ($target + '.RebootPolicy') }

        $detMode = [string]$cboDetectionMode.SelectedItem
        if ($detMode -eq 'Custom') {
            Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Mode' -Value 'Custom'
            $type = [string]$cboDetectionType.SelectedItem
            $rule = [ordered]@{
                Type           = $(if ($type -eq 'PowerShell') { 'Script' } else { $type })
                Hive           = [string]$cboDetHive.SelectedItem
                View           = $(if ([string]$cboDetView.SelectedItem -eq '32-bit') { '32' } else { '64' })
                Key            = [string]$txtDetKey.Text
                ValueName      = [string]$txtDetValueName.Text
                Operator       = [string]$cboDetOperator.SelectedItem
                ExpectedValue  = [string]$txtDetExpected.Text
                VersionBinding = $(if ([string]$cboDetVersionBinding.SelectedItem -eq 'Pinned') { 'Pinned' } else { 'FollowsStaged' })
                FilePath       = [string]$txtDetFilePath.Text
                FileName       = [string]$txtDetFileName.Text
                FileProperty   = [string]$cboDetFileProperty.SelectedItem
                Logic          = $(switch ([string]$cboDetLogic.SelectedItem) { 'Or' { 'Or' } 'Two groups' { 'TwoGroup' } default { 'And' } })
                ScriptText     = [string]$txtDetectScript.Text
                Clauses        = @(foreach ($r in $clauseRows) {
                    [ordered]@{ Group = [string]$r.Group; Type = [string]$r.Type; Subject = [string]$r.Subject; Operator = [string]$r.Operator; Expected = [string]$r.Expected }
                })
            }
            Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Rule' -Value $rule
            Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.IntuneScriptConversion' -Value ($chkIntuneScriptConversion.IsChecked -eq $true)
        }
        else {
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Mode'
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.Rule'
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Detection.IntuneScriptConversion'
        }

        if ($chkEstimatedDefault.IsChecked -eq $true) { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.EstimatedMinutes' }
        else {
            $n = 0
            [void][int]::TryParse([string]$txtEstimatedMinutes.Text, [ref]$n)
            Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.EstimatedMinutes' -Value $n
        }
        if ($chkMaximumDefault.IsChecked -eq $true) { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.MaximumMinutes' }
        else {
            $n = 0
            [void][int]::TryParse([string]$txtMaximumMinutes.Text, [ref]$n)
            Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Timing.MaximumMinutes' -Value $n
        }

        foreach ($pair in @(@{ P = 'Execution.Context'; C = $cboExecContext; D = [string]$wb.Inherited.ExecutionContext },
                            @{ P = 'Execution.LogonRequirement'; C = $cboExecLogon; D = [string]$wb.Inherited.LogonRequirement },
                            @{ P = 'Execution.UserInteraction'; C = $cboExecInteraction; D = [string]$wb.Inherited.UserInteraction },
                            @{ P = 'Execution.ScriptHost'; C = $cboExecScriptHost; D = [string]$wb.Inherited.ScriptHost })) {
            $v = [string]$pair.C.SelectedItem
            if (-not $v -or $v -eq $pair.D) { Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path $pair.P }
            else { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path $pair.P -Value $v }
        }

        $split = [string]$cboVariantSplit.SelectedItem
        if ($split -and $split -ne 'None') { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Variants.Split' -Value @($split) }
        else { Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Variants.Split' -Value @() }

        switch ([string]$cboInstallForMode.SelectedItem) {
            'System' { $wb.Profile.InstallMode = 'AllUsers' }
            'User'   { $wb.Profile.InstallMode = 'CurrentUser' }
            default  { $wb.Profile.InstallMode = $null }
        }

        $ops = New-Object System.Collections.ArrayList
        foreach ($r in $requirementRows) {
            $rule = [ordered]@{ ConditionId = [string]$r.Summary }
            if ([string]$r.Value) {
                if ([string]$r.Summary -eq 'os-language') { $rule['Cultures'] = @([string]$r.Value -split '\s*,\s*' | Where-Object { $_ }) }
                else { $rule['Value'] = [string]$r.Value }
            }
            [void]$ops.Add([ordered]@{
                Op        = [string]$r.Op
                RuleId    = [string]$r.RuleId
                Rule      = $rule
                AppliesTo = @([string]$r.AppliesTo -split '\s*,\s*' | Where-Object { $_ })
            })
        }
        Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Requirements.Operations' -Value @($ops)

        $files = New-Object System.Collections.ArrayList
        foreach ($r in $sourceFileRows) {
            [void]$files.Add([ordered]@{
                Asset       = [string]$r.Asset
                Destination = [string]$r.Destination
                Sha256      = [string]$r.Sha256
                Size        = [int64]$r.Size
                Provenance  = [string]$r.Provenance
                Linked      = [bool]$r.Linked
            })
        }
        $wb.Profile.SourceFiles = @($files)
    }

    $validateProfile = {
        & $commitUi
        $findings = @(Get-WorkbenchLocalFindings -Profile $wb.Profile -Inherited $wb.Inherited)
        $manifestCmd = Get-WorkbenchCommand -Name 'Get-IntuneCompatibilityFindings'
        if (-not $manifestCmd) {
            $findings += [pscustomobject]@{ Severity = 'Info'; Code = 'INTUNE-ADAPTER'; Message = 'The Intune compatibility check is not loaded; only local validation ran.' }
        }
        $wb.Findings = $findings
        $wb.Validated = $true
        $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
        foreach ($f in $findings) { $rows.Add($f) }
        $dgFindings.ItemsSource = $rows
        $chips = Get-WorkbenchChipState -Findings $findings -Validated $true
        $chipContent.Text = 'Content build: ' + $chips.Content
        $chipMecm.Text = 'ConfigMgr: ' + $chips.Mecm
        $chipIntune.Text = 'Intune: ' + $chips.Intune
        & $refreshReview
        return $findings
    }

    $saveProfile = {
        param([string]$newName)
        & $commitUi
        # The packager default profile is never stored or edited: the first
        # save of a default-profile edit becomes a named profile.
        if ($newName -or $wb.ProfileId -eq 'default') {
            $name = $newName
            if (-not $name) {
                $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for this profile' -Value 'Managed'
                if (-not $name) { return $false }
            }
            $wb.ProfileId = ([guid]::NewGuid().ToString('N'))
            $wb.Profile['ProfileId'] = $wb.ProfileId
            $wb.Profile['Name'] = $name
            $wb.Profile['Revision'] = 0
        }
        $wb.Profile['IsDefault'] = $false
        try {
            Import-WorkbenchProfileAssets -Profile $wb.Profile
            $saved = Save-Profile -Profile ([pscustomobject]$wb.Profile) -SetActive
        }
        catch {
            [void](Show-ThemedMessage -Owner $win -Title 'Save Failed' -Message $_.Exception.Message -Buttons OK -Icon Error)
            return $false
        }
        $wb.Profile = ConvertTo-WorkbenchUiHashtable -InputObject $saved
        $wb.ProfileId = [string]$wb.Profile['ProfileId']
        try { [void](Remove-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId) } catch { }
        $wb.Baseline = ($wb.Profile | ConvertTo-Json -Depth 12 -Compress)
        $wb.Dirty = $false
        & $loadProfileList $wb.ProfileId
        $wb.Loading = $true
        try { $txtProfileName.Text = [string]$wb.Profile['Name'] } finally { $wb.Loading = $false }
        & $refreshEditors
        & $setStatus ('Saved "' + [string]$wb.Profile['Name'] + '" at revision ' + [string]$wb.Profile['Revision'] + '. Existing builds are now stale relative to this profile; nothing was deployed.')
        return $true
    }

    $confirmDiscard = {
        param([string]$message)
        if (-not $wb.Dirty) { return $true }
        $answer = Show-WorkbenchUnsavedDialog -Owner $win -Message $message
        switch ($answer) {
            'Save'    { return [bool](& $saveProfile '') }
            'Discard' {
                try { [void](Remove-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId) } catch { }
                $wb.Dirty = $false
                return $true
            }
            default   { return $false }
        }
    }

    # ---- Handlers -----------------------------------------------------
    $lstSections.Add_SelectionChanged({
        $idx = $lstSections.SelectedIndex
        for ($i = 0; $i -lt $sectionPanels.Count; $i++) {
            $sectionPanels[$i].Visibility = $(if ($i -eq $idx) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
        }
        if ($idx -eq 6) { & $refreshReview }
    })

    $cboApplication.Add_SelectionChanged({
        if ($wb.Loading) { return }
        $sel = $cboApplication.SelectedItem
        if (-not $sel) { return }
        $target = @($wb.Applications | Where-Object { [string]$_.DisplayLabel -eq [string]$sel })
        if ($target.Count -eq 0) { return }
        if ([string]$target[0].ApplicationId -eq [string]$wb.Application.ApplicationId) { return }
        if (-not (& $confirmDiscard 'This application has unsaved changes. Save them before switching?')) {
            $wb.Loading = $true
            try { $cboApplication.SelectedItem = [string]$wb.Application.DisplayLabel } finally { $wb.Loading = $false }
            return
        }
        & $selectApplication $target[0] ''
    })

    $cboProfile.Add_SelectionChanged({
        if ($wb.Loading) { return }
        $idx = $cboProfile.SelectedIndex
        if ($idx -lt 0) { return }
        if ($idx -ge @($script:WorkbenchProfileIds).Count) {
            # "Save as..." entry.
            $wb.Loading = $true
            try { $cboProfile.SelectedIndex = [array]::IndexOf($script:WorkbenchProfileIds, $wb.ProfileId) } finally { $wb.Loading = $false }
            $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for the new profile' -Value ([string]$wb.Profile.Name + ' copy')
            if ($name) { [void](& $saveProfile $name) }
            return
        }
        $newId = [string]@($script:WorkbenchProfileIds)[$idx]
        if ($newId -eq $wb.ProfileId) { return }
        if (-not (& $confirmDiscard 'This profile has unsaved changes. Save them before switching?')) {
            $wb.Loading = $true
            try { $cboProfile.SelectedIndex = [array]::IndexOf($script:WorkbenchProfileIds, $wb.ProfileId) } finally { $wb.Loading = $false }
            return
        }
        & $selectApplication $wb.Application $newId
    })

    $btnProfileRename.Add_Click({
        if ($wb.ProfileId -eq 'default') {
            [void](Show-ThemedMessage -Owner $win -Title 'Rename Profile' -Message 'The packager default profile cannot be renamed. Use Save as to create a named profile.' -Buttons OK -Icon Info)
            return
        }
        $name = Show-WorkbenchNameDialog -Owner $win -Title 'New name for this profile' -Value ([string]$wb.Profile.Name)
        if (-not $name) { return }
        $txtProfileName.Text = $name
        & $markDirty
    })

    foreach ($descriptor in $fieldDescriptors) {
        $d = $descriptor
        if ($d.Kind -eq 'Combo') { $d.Control.Add_SelectionChanged({ & $markDirty }.GetNewClosure()) }
        else { $d.Control.Add_TextChanged({ & $markDirty }.GetNewClosure()) }
        if ($d.ResetButton) {
            $d.ResetButton.Add_Click({
                Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path $d.Path
                $wb.Suppress = $true
                try {
                    if ($d.Kind -eq 'Combo') { $d.Control.SelectedItem = 'Packager default' } else { $d.Control.Text = '' }
                }
                finally { $wb.Suppress = $false }
                & $markDirty
                & $refreshFieldIndicators
            }.GetNewClosure())
        }
    }

    $cboScriptTarget.Add_SelectionChanged({
        if ($wb.Loading -or $wb.Suppress) { return }
        & $commitUi
        $wb.ScriptTarget = [string]$cboScriptTarget.SelectedItem
        & $refreshScriptEditor
        & $refreshFieldIndicators
    })
    $cboScriptTarget.SelectedIndex = 0

    $cboInstallMode.Add_SelectionChanged({
        if ($wb.Loading -or $wb.Suppress) { return }
        & $markDirty
        & $commitUi
        & $refreshScriptEditor
        & $refreshFieldIndicators
    })

    foreach ($box in @($txtReturnCodes, $txtWorkingDirectory, $txtHookBefore, $txtHookAfter)) {
        $box.Add_TextChanged({ & $markDirty }.GetNewClosure())
    }
    $chkAfterOnFailure.Add_Checked({ & $markDirty })
    $chkAfterOnFailure.Add_Unchecked({ & $markDirty })
    $cboRebootPolicy.Add_SelectionChanged({ & $markDirty })

    $parseTimer.Add_Tick({
        $parseTimer.Stop()
        $lstScriptDiagnostics.Items.Clear()
        foreach ($line in (Get-WorkbenchParseDiagnostics -Text ([string]$txtScript.Text))) { [void]$lstScriptDiagnostics.Items.Add($line) }
        if ($lstScriptDiagnostics.Items.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$txtScript.Text)) {
            [void]$lstScriptDiagnostics.Items.Add('No parse errors. Validation inspects the script; it never runs it.')
        }
        $lstDetectDiagnostics.Items.Clear()
        foreach ($line in (Get-WorkbenchParseDiagnostics -Text ([string]$txtDetectScript.Text))) { [void]$lstDetectDiagnostics.Items.Add($line) }
        if ($lstDetectDiagnostics.Items.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$txtDetectScript.Text)) {
            if (Test-WorkbenchDetectionScriptOutput -Text ([string]$txtDetectScript.Text)) {
                [void]$lstDetectDiagnostics.Items.Add('No parse errors; the script writes to STDOUT.')
            }
            else {
                [void]$lstDetectDiagnostics.Items.Add('No parse errors, but nothing is written to STDOUT: Intune would read this as not installed.')
            }
        }
    })

    $txtScript.Add_TextChanged({
        $txtScriptGutter.Text = Get-WorkbenchLineNumberText -Text ([string]$txtScript.Text)
        $parseTimer.Stop()
        $parseTimer.Start()
        & $markDirty
    })
    $txtDetectScript.Add_TextChanged({
        $txtDetectGutter.Text = Get-WorkbenchLineNumberText -Text ([string]$txtDetectScript.Text)
        $parseTimer.Stop()
        $parseTimer.Start()
        & $markDirty
    })
    # The gutter is a separate TextBox, so it has to follow the editor's
    # own vertical offset rather than its own scrollbar.
    $txtScript.Add_TextChanged({ $txtScriptGutter.ScrollToVerticalOffset($txtScript.VerticalOffset) })
    $txtScript.AddHandler([System.Windows.Controls.ScrollViewer]::ScrollChangedEvent, [System.Windows.RoutedEventHandler]{
        $txtScriptGutter.ScrollToVerticalOffset($txtScript.VerticalOffset)
    })
    $txtDetectScript.AddHandler([System.Windows.Controls.ScrollViewer]::ScrollChangedEvent, [System.Windows.RoutedEventHandler]{
        $txtDetectGutter.ScrollToVerticalOffset($txtDetectScript.VerticalOffset)
    })

    $findNext = {
        $needle = [string]$txtFind.Text
        if (-not $needle) { return }
        $box = if ($pnlDetectionScript.IsVisible) { $txtDetectScript } else { $txtScript }
        $start = $box.SelectionStart + [Math]::Max($box.SelectionLength, 1)
        if ($start -ge $box.Text.Length) { $start = 0 }
        $idx = $box.Text.IndexOf($needle, $start, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -lt 0) { $idx = $box.Text.IndexOf($needle, 0, [System.StringComparison]::OrdinalIgnoreCase) }
        if ($idx -ge 0) {
            $box.Focus()
            $box.Select($idx, $needle.Length)
            $box.ScrollToLine([Math]::Max(0, $box.GetLineIndexFromCharacterIndex($idx)))
        }
    }
    $btnFindNext.Add_Click({ & $findNext })
    $btnFindClose.Add_Click({ $pnlFind.Visibility = [System.Windows.Visibility]::Collapsed })
    $txtFind.Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) { & $findNext; $e.Handled = $true }
    })
    $win.Add_PreviewKeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::F -and
            ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
            $pnlFind.Visibility = [System.Windows.Visibility]::Visible
            $txtFind.Focus()
            $e.Handled = $true
        }
    })

    $cboDetectionMode.Add_SelectionChanged({
        if ($wb.Loading -or $wb.Suppress) { return }
        & $markDirty
        & $commitUi
        & $refreshDetectionEditor
        & $refreshFieldIndicators
    })
    $cboDetectionType.Add_SelectionChanged({
        if ($wb.Loading -or $wb.Suppress) { return }
        & $markDirty
        & $commitUi
        & $refreshDetectionEditor
    })
    foreach ($c in @($txtDetKey, $txtDetValueName, $txtDetExpected, $txtDetFilePath, $txtDetFileName)) {
        $c.Add_TextChanged({ & $markDirty }.GetNewClosure())
    }
    foreach ($c in @($cboDetHive, $cboDetView, $cboDetOperator, $cboDetVersionBinding, $cboDetFileProperty, $cboDetLogic)) {
        $c.Add_SelectionChanged({ & $markDirty }.GetNewClosure())
    }
    $chkIntuneScriptConversion.Add_Checked({ & $markDirty })
    $chkIntuneScriptConversion.Add_Unchecked({ & $markDirty })

    $btnDetClauseAdd.Add_Click({
        $clauseRows.Add([pscustomobject]@{ Group = '1'; Type = 'Registry'; Subject = ''; Operator = 'Exists'; Expected = '' })
        & $markDirty
    })
    $btnDetClauseRemove.Add_Click({
        $sel = $dgDetClauses.SelectedItem
        if ($sel) { [void]$clauseRows.Remove($sel); & $markDirty }
    })

    $btnReqAdd.Add_Click({
        $id = [string]$cboReqTemplate.SelectedItem
        if (-not $id) { return }
        $requirementRows.Add([pscustomobject]@{ Op = 'Add'; RuleId = $id; Summary = $id; Value = ''; AppliesTo = '*' })
        & $markDirty
    })
    $btnReqReplace.Add_Click({
        $sel = $dgRequirements.SelectedItem
        if (-not $sel) { return }
        $sel.Op = 'Replace'
        $dgRequirements.Items.Refresh()
        & $markDirty
    })
    $btnReqRemove.Add_Click({
        $sel = $dgRequirements.SelectedItem
        if (-not $sel) { return }
        # An explicit Remove operation is not the same as dropping the row:
        # it suppresses a rule the packager would otherwise contribute.
        $sel.Op = 'Remove'
        $dgRequirements.Items.Refresh()
        & $markDirty
    })
    $cboVariantSplit.Add_SelectionChanged({ & $markDirty })
    $cboInstallForMode.Add_SelectionChanged({ & $markDirty })

    $chkEstimatedDefault.Add_Checked({ $txtEstimatedMinutes.IsEnabled = $false; $txtEstimatedMinutes.Text = [string]$wb.Inherited.EstimatedMinutes; & $markDirty })
    $chkEstimatedDefault.Add_Unchecked({ $txtEstimatedMinutes.IsEnabled = $true; & $markDirty })
    $chkMaximumDefault.Add_Checked({ $txtMaximumMinutes.IsEnabled = $false; $txtMaximumMinutes.Text = [string]$wb.Inherited.MaximumMinutes; & $markDirty })
    $chkMaximumDefault.Add_Unchecked({ $txtMaximumMinutes.IsEnabled = $true; & $markDirty })
    $txtEstimatedMinutes.Add_TextChanged({ & $markDirty })
    $txtMaximumMinutes.Add_TextChanged({ & $markDirty })
    foreach ($c in @($cboExecContext, $cboExecLogon, $cboExecInteraction, $cboExecScriptHost)) {
        $c.Add_SelectionChanged({ & $markDirty }.GetNewClosure())
    }
    $btnTimingReset.Add_Click({
        foreach ($p in @('Timing.EstimatedMinutes', 'Timing.MaximumMinutes', 'Execution.Context',
                         'Execution.LogonRequirement', 'Execution.UserInteraction', 'Execution.ScriptHost')) {
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path $p
        }
        & $loadProfileIntoUi
        & $markDirty
    })

    $btnInstallReset.Add_Click({
        foreach ($p in @('Mode', 'Command', 'ScriptText', 'BeforeText', 'AfterText', 'AfterRunsOnFailure',
                         'ReturnCodes', 'RebootPolicy', 'WorkingDirectory')) {
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path ($wb.ScriptTarget + '.' + $p)
        }
        & $refreshScriptEditor
        & $refreshFieldIndicators
        & $markDirty
    })
    $btnDetectionReset.Add_Click({
        foreach ($p in @('Detection.Mode', 'Detection.Rule', 'Detection.IntuneScriptConversion')) {
            Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path $p
        }
        & $refreshDetectionEditor
        & $refreshFieldIndicators
        & $markDirty
    })

    $addSourceEntry = {
        param([string]$path)
        if (-not (Test-Path -LiteralPath $path)) { return }
        $item = Get-Item -LiteralPath $path -Force
        if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            [void](Show-ThemedMessage -Owner $win -Title 'Source Files' -Message ('Reparse points are not accepted as source files: ' + $path) -Buttons OK -Icon Warning)
            return
        }
        $files = @()
        if ($item.PSIsContainer) { $files = @(Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue) }
        else { $files = @($item) }
        foreach ($f in $files) {
            $dest = $f.Name
            if ($item.PSIsContainer) {
                $rel = $f.FullName.Substring($item.FullName.Length).TrimStart('\')
                $dest = Join-Path $item.Name $rel
            }
            $hash = ''
            try { $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant() } catch { }
            $sourceFileRows.Add([pscustomobject]@{
                Asset       = ''
                Provenance  = $f.FullName
                Destination = $dest
                SizeText    = ('{0:N0} bytes' -f $f.Length)
                ShortHash   = $(if ($hash.Length -ge 16) { $hash.Substring(0, 16) } else { $hash })
                Sha256      = $hash
                Size        = [int64]$f.Length
                Linked      = $false
            })
        }
        & $markDirty
    }

    $btnSourceAddFile.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Add source file'
        $dlg.Multiselect = $true
        if ($dlg.ShowDialog() -eq $true) { foreach ($p in @($dlg.FileNames)) { & $addSourceEntry $p } }
    })
    $btnSourceAddFolder.Add_Click({
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Add every file under this folder'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { & $addSourceEntry $dlg.SelectedPath }
    })
    $btnSourceReplace.Add_Click({
        $sel = $dgSourceFiles.SelectedItem
        if (-not $sel) { return }
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Replace source file'
        if ($dlg.ShowDialog() -ne $true) { return }
        $f = Get-Item -LiteralPath $dlg.FileName
        $hash = ''
        try { $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant() } catch { }
        $sel.Provenance = $f.FullName
        $sel.Sha256 = $hash
        $sel.ShortHash = $(if ($hash.Length -ge 16) { $hash.Substring(0, 16) } else { $hash })
        $sel.Size = [int64]$f.Length
        $sel.SizeText = ('{0:N0} bytes' -f $f.Length)
        $sel.Asset = ''
        $dgSourceFiles.Items.Refresh()
        & $markDirty
    })
    $btnSourceRemove.Add_Click({
        $sel = $dgSourceFiles.SelectedItem
        if ($sel) { [void]$sourceFileRows.Remove($sel); & $markDirty }
    })

    $btnIconChoose.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Choose icon'
        $dlg.Filter = 'Image files|*.png;*.ico;*.jpg;*.jpeg;*.bmp|All files|*.*'
        if ($dlg.ShowDialog() -ne $true) { return }
        Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Application.Icon' -Value ([ordered]@{ Path = $dlg.FileName; Asset = '' })
        & $markDirty
        & $loadProfileIntoUi
        $wb.Dirty = $true
    })
    $btnIconReset.Add_Click({
        Clear-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Application.Icon'
        & $markDirty
        & $loadProfileIntoUi
        $wb.Dirty = $true
    })
    $btnIconRemove.Add_Click({
        # Explicit null is "publish no icon", distinct from inheriting one.
        Set-WorkbenchOverrideValue -Profile $wb.Profile -Path 'Application.Icon' -Value $null
        & $markDirty
        & $loadProfileIntoUi
        $wb.Dirty = $true
    })

    $btnUseDefaults.Add_Click({
        $answer = Show-ThemedMessage -Owner $win -Title 'Use Defaults' `
            -Message 'Clear every override in this profile and inherit the packager defaults? Saved builds are not touched.' `
            -Buttons YesNo -Icon Question
        if ($answer -ne 'Yes') { return }
        $name = [string]$wb.Profile.Name
        $revision = [int]$wb.Profile.Revision
        $wb.Profile = ConvertTo-WorkbenchUiHashtable -InputObject (New-WorkbenchProfileObject -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId -Name $name)
        $wb.Profile['Revision'] = $revision
        & $loadProfileIntoUi
        $wb.Dirty = $true
        & $markDirty
    })

    $btnSave.Add_Click({ [void](& $saveProfile '') })
    $btnSaveAs.Add_Click({
        $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for the new profile' -Value ([string]$wb.Profile.Name + ' copy')
        if ($name) { [void](& $saveProfile $name) }
    })
    $btnValidate.Add_Click({
        [void](& $validateProfile)
        $lstSections.SelectedIndex = 6
        & $setStatus 'Validation inspected the profile without running anything. Results are per target.'
    })

    $runFromWorkbench = {
        param([string]$operation)
        if ($wb.Dirty) {
            if (-not (& $confirmDiscard 'Save the profile before this run? A run uses the saved profile revision.')) { return }
        }
        $findings = @(& $validateProfile)
        $blocking = @($findings | Where-Object { $_.Severity -eq 'Blocking' -and [string]$_.Code -notlike '*INTUNE*' })
        if ($blocking.Count -gt 0) {
            [void](Show-ThemedMessage -Owner $win -Title 'Blocked' -Message ('This profile has blocking findings: ' + [string]$blocking[0].Message) -Buttons OK -Icon Warning)
            return
        }
        if (-not [string]$wb.Application.ScriptPath) {
            [void](Show-ThemedMessage -Owner $win -Title $operation `
                -Message 'This application has no packager script. Bring-your-own applications run through the installer intake flow.' -Buttons OK -Icon Info)
            return
        }
        $selectedBuild = [string]$cboBuild.SelectedItem
        if ($operation -eq 'Package' -and ($selectedBuild -eq 'No sealed builds' -or -not $selectedBuild)) {
            [void](Show-ThemedMessage -Owner $win -Title 'Package' `
                -Message 'No sealed build exists for this profile. Stage it first; Package consumes an exact build id, not the newest manifest on disk.' -Buttons OK -Icon Warning)
            return
        }
        $win.Close()
        $row = @($script:PackagerData | Where-Object { [string]$_.FullPath -eq [string]$wb.Application.ScriptPath })
        if ($row.Count -eq 0) {
            $cliHint = ('.\Invoke-AppPackagerBuild.ps1 -Application {0} -Profile {1} -{2}' -f [string]$wb.Application.ApplicationId, [string]$wb.ProfileId, $operation)
            Add-LogLine -Message ('Workbench run skipped: {0} is not in the current grid. Build it from the command line: {1}' -f [string]$wb.Application.DisplayName, $cliHint)
            [void](Show-ThemedMessage -Owner $win -Title $operation `
                -Message ('This application is not in the main grid, so it cannot run through the grid pipeline. Build it from the command line:' + "`r`n`r`n" + $cliHint) -Buttons OK -Icon Info)
            return
        }
        Add-LogLine -Message ('Workbench {0}: {1} (profile {2}, revision {3})' -f $operation, [string]$wb.Application.DisplayName, [string]$wb.Profile.Name, [string]$wb.Profile.Revision)
        Invoke-WorkbenchRun -Operation $operation -Rows $row -BuildId $(if ($operation -eq 'Package') { $selectedBuild } else { '' })
    }
    $btnStage.Add_Click({ & $runFromWorkbench 'Stage' })
    $btnPackage.Add_Click({ & $runFromWorkbench 'Package' })

    $win.Add_Closing({
        param($s, $e)
        if ($wb.Dirty) {
            if (-not (& $confirmDiscard 'This profile has unsaved changes. Save them before closing?')) {
                $e.Cancel = $true
                return
            }
        }
        $parseTimer.Stop()
        Save-WindowState -Window $win -Path (Get-WorkbenchWindowStatePath)
    })

    # ---- Population ---------------------------------------------------
    $apps = @(Get-WorkbenchUiApplications)
    foreach ($a in $apps) {
        $label = [string]$a.DisplayName
        if ([string]$a.Publisher) { $label = [string]$a.Publisher + ' - ' + $label }
        $a | Add-Member -NotePropertyName DisplayLabel -NotePropertyValue $label -Force
    }
    $wb.Applications = @($apps | Sort-Object DisplayLabel)
    if (@($wb.Applications).Count -eq 0) {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Application Workbench' `
            -Message 'No applications were found. Load packagers or save a bring-your-own installer first.' -Buttons OK -Icon Info)
        return
    }
    foreach ($a in $wb.Applications) { [void]$cboApplication.Items.Add([string]$a.DisplayLabel) }

    $initial = $null
    if ($PreselectApplicationId) { $initial = @($wb.Applications | Where-Object { [string]$_.ApplicationId -eq $PreselectApplicationId })[0] }
    if (-not $initial -and $PreselectPackagerBase) { $initial = @($wb.Applications | Where-Object { [string]$_.PackagerBase -eq $PreselectPackagerBase })[0] }
    if (-not $initial) { $initial = $wb.Applications[0] }

    $lstSections.SelectedIndex = 0
    & $selectApplication $initial ''
    $wb.Loading = $true
    try { $cboApplication.SelectedItem = [string]$initial.DisplayLabel } finally { $wb.Loading = $false }

    if ($Probe) {
        # UI probe seam: the window is fully built and populated but never
        # shown, so a headless run can drive the same handlers.
        & $Probe @{
            Window            = $win
            State             = $wb
            Sections          = $lstSections
            Panels            = $sectionPanels
            Find              = $ctl
            Commit            = $commitUi
            Save              = $saveProfile
            Validate          = $validateProfile
            Reload            = $loadProfileIntoUi
            SelectApplication = $selectApplication
            ConfirmDiscard    = $confirmDiscard
            MarkDirty         = $markDirty
            Refresh           = $refreshEditors
        }
        $parseTimer.Stop()
        return
    }

    Restore-WindowState -Window $win -Path (Get-WorkbenchWindowStatePath)
    [void]$win.ShowDialog()
}

function New-WorkbenchPipelineContext {
    # Every per-app map is built here on the UI thread: the background
    # runspace has its own session state and cannot read the preferences.
    param(
        [Parameter(Mandatory)][ValidateSet('Stage', 'Package')][string]$Operation,
        [array]$Rows = @(),
        [AllowEmptyString()][string]$BuildId = ''
    )
    $target = [string]$script:Prefs.Intune.DeploymentTarget
    $context = @{
        DownloadRoot   = $script:Prefs.DownloadRoot
        M365Channel    = $script:Prefs.M365Channel
        M365DeployMode = $script:Prefs.M365DeployMode
        LogFolder      = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath   = Get-SevenZipPathForContext
        RunPlanByApp   = Get-WorkbenchRunPlanForContext -Rows $Rows -Target $target
        SigningJson    = Get-WorkbenchSigningPolicyJson
        SigningDigest  = Get-WorkbenchSigningPolicyDigest
    }
    if ($Operation -eq 'Package') {
        $context['SiteCode']             = $script:Prefs.SiteCode
        $context['ProviderMachineName']  = $script:Prefs.ProviderMachineName
        $context['Comment']              = $txtComment.Text.Trim()
        $context['FileShareRoot']        = $script:Prefs.FileShareRoot
        $context['ContentLayout']        = $script:Prefs.ContentLayout
        $context['EstimatedRuntimeMins'] = $script:Prefs.EstimatedRuntimeMins
        $context['MaximumRuntimeMins']   = $script:Prefs.MaximumRuntimeMins
        $context['IntuneWinCreate']      = ([bool]$script:Prefs.Intune.CreateIntuneWin -and -not [string]::IsNullOrWhiteSpace((Get-IntuneWinToolPathForContext)))
        $context['IntuneWinToolPath']    = Get-IntuneWinToolPathForContext
        $context['RequirementsByApp']    = Get-RequirementsMapForContext
        $context['VariantsByApp']        = Get-VariantsMapForContext
        $context['CommandsByApp']        = Get-CommandsMapForContext
        $context['InstallModesByApp']    = Get-InstallModesMapForContext
        $context['TitleModesByApp']      = Get-TitleModesMapForContext
        $context['DefaultTitleMode']     = Get-DefaultTitleModeForContext
        $context['IntunePublishConfig']  = Get-IntunePublishConfigForContext
        $context['DeploymentTarget']     = $target
        $context['BuildId']                = $BuildId
    }
    return $context
}

function Invoke-WorkbenchRun {
    # Stage or Package a single application through the shared pipeline, so
    # a workbench run lands the same history and status a grid run does.
    param(
        [Parameter(Mandatory)][ValidateSet('Stage', 'Package')][string]$Operation,
        [Parameter(Mandatory)][array]$Rows,
        [AllowEmptyString()][string]$BuildId = ''
    )
    if ($Operation -eq 'Stage') {
        $Rows = @(Confirm-LocalSourceFolders -Rows $Rows)
        if ($Rows.Count -eq 0) { return }
    }
    $context = New-WorkbenchPipelineContext -Operation $Operation -Rows $Rows -BuildId $BuildId
    Invoke-MultiAppPipeline -Operation $Operation -Rows $Rows -Context $context
}

function New-ScriptSigningPanel {
    # Signing creation and signature enforcement are separate: every
    # switch here defaults off, and a Require flag blocks a publish rather
    # than falling back to unsigned content.
    $xaml = @'
<ScrollViewer xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
      VerticalScrollBarVisibility="Auto">
    <StackPanel>
        <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,0,0,12"
                   Text="Authenticode signing for the scripts AppPackager generates. Signing a detection script does not sign install wrappers, and signing deployment scripts does not sign requirement scripts. A local signature check does not prove the certificate is trusted on your clients."/>

        <TextBlock Text="Sign" FontSize="13" FontWeight="Bold" Margin="0,0,0,6"/>
        <CheckBox x:Name="chkSignDetection" FontSize="12" Margin="0,0,0,4"
                  Content="Sign detection scripts" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Signs generated and custom detection scripts. An unchanged import carrying a valid third-party signature is left alone."/>
        <CheckBox x:Name="chkSignRequirements" FontSize="12" Margin="0,0,0,4"
                  Content="Sign requirement scripts and script global conditions" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Covers script-based requirement rules including the VPN predicate. WQL and native rules need no signature."/>
        <CheckBox x:Name="chkSignDeployment" FontSize="12" Margin="0,0,0,12"
                  Content="Sign install and uninstall PowerShell scripts" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Signs the deployment execution chain and removes the execution-policy argument from its generated launchers, letting the configured endpoint policy govern. A vendor script that already carries an intact signature is left alone; an unsigned one is signed with this certificate."/>

        <TextBlock Text="Require valid signatures" FontSize="13" FontWeight="Bold" Margin="0,0,0,6"/>
        <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,0,0,6"
                   Text="A publishing constraint, not a change to client execution policy. When a required category fails verification the build stops; there is no unsigned fallback."/>
        <CheckBox x:Name="chkRequireDetection" FontSize="12" Margin="0,0,0,4"
                  Content="Require valid detection signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        <CheckBox x:Name="chkRequireRequirements" FontSize="12" Margin="0,0,0,4"
                  Content="Require valid requirement signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        <CheckBox x:Name="chkRequireDeployment" FontSize="12" Margin="0,0,0,12"
                  Content="Require valid deployment script signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

        <TextBlock Text="Signing identity" FontSize="13" FontWeight="Bold" Margin="0,0,0,6"/>
        <Grid Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Certificate store" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <ComboBox  Grid.Row="0" Grid.Column="1" x:Name="cboSignStore" FontSize="12" Height="28" Margin="0,0,0,6"
                       ToolTip="CurrentUser\My is the supported identity. LocalMachine\My needs an explicit choice and tested private-key access."/>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Certificate" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <ComboBox  Grid.Row="1" Grid.Column="1" x:Name="cboSignCertificate" FontSize="12" Height="28" Margin="0,0,0,6"
                       ToolTip="Selected by thumbprint. A renewed certificate is a new thumbprint and needs reselecting here."/>
            <Button    Grid.Row="1" Grid.Column="2" x:Name="btnSignRefresh" Content="Refresh" MinWidth="100" Height="28" Margin="8,0,0,6"
                       Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

            <TextBlock Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtSignCertDetail" FontSize="11" TextWrapping="Wrap"
                       Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="3" Grid.Column="0" Text="Timestamp server" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtSignTimestamp" FontSize="12" Height="28" VerticalContentAlignment="Center" Margin="0,0,0,6"
                       Controls:TextBoxHelper.Watermark="none"
                       ToolTip="A timestamped signature stays valid after the certificate expires. A signature that exists is not the same as a verified timestamp."/>
            <CheckBox  Grid.Row="4" Grid.Column="1" x:Name="chkSignTimestampRequired" FontSize="12" Margin="0,0,0,6"
                       Content="Timestamp required" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                       ToolTip="With this on, a timestamp server failure fails the signed build instead of producing an untimestamped signature."/>
        </Grid>

        <StackPanel Orientation="Horizontal" Margin="0,4,0,8">
            <Button x:Name="btnSignTest" Content="Test signing configuration" MinWidth="200" Height="28" Margin="0,0,10,0"
                    Style="{DynamicResource MahApps.Styles.Button.Square}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                    ToolTip="Signs and verifies a temporary file with the selected certificate. Nothing is published."/>
        </StackPanel>
        <TextBlock x:Name="txtSignTestResult" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,10"/>

        <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                   Text="Private keys are never exported and never reach profiles, manifests, logs or command lines. A key held on a hardware token may require a PIN and can prevent unattended runs."/>
    </StackPanel>
</ScrollViewer>
'@

    [xml]$xml = $xaml
    $reader  = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $chkSignDetection    = $element.FindName('chkSignDetection')
    $chkSignRequirements = $element.FindName('chkSignRequirements')
    $chkSignDeployment   = $element.FindName('chkSignDeployment')
    $chkRequireDetection = $element.FindName('chkRequireDetection')
    $chkRequireRequirements = $element.FindName('chkRequireRequirements')
    $chkRequireDeployment = $element.FindName('chkRequireDeployment')
    $cboSignStore        = $element.FindName('cboSignStore')
    $cboSignCertificate  = $element.FindName('cboSignCertificate')
    $btnSignRefresh      = $element.FindName('btnSignRefresh')
    $txtSignCertDetail   = $element.FindName('txtSignCertDetail')
    $txtSignTimestamp    = $element.FindName('txtSignTimestamp')
    $chkSignTimestampRequired = $element.FindName('chkSignTimestampRequired')
    $btnSignTest         = $element.FindName('btnSignTest')
    $txtSignTestResult   = $element.FindName('txtSignTestResult')

    foreach ($v in @('CurrentUser', 'LocalMachine')) { [void]$cboSignStore.Items.Add($v) }

    $signing = $script:Prefs.ScriptSigning
    $chkSignDetection.IsChecked    = [bool]$signing.SignDetection
    $chkSignRequirements.IsChecked = [bool]$signing.SignRequirements
    $chkSignDeployment.IsChecked   = [bool]$signing.SignDeployment
    $chkRequireDetection.IsChecked = [bool]$signing.RequireDetection
    $chkRequireRequirements.IsChecked = [bool]$signing.RequireRequirements
    $chkRequireDeployment.IsChecked = [bool]$signing.RequireDeployment
    $cboSignStore.SelectedItem     = $(if ([string]$signing.StoreLocation -eq 'LocalMachine') { 'LocalMachine' } else { 'CurrentUser' })
    $txtSignTimestamp.Text         = [string]$signing.TimestampServer
    $chkSignTimestampRequired.IsChecked = [bool]$signing.TimestampRequired

    # Thumbprints run parallel to the combo items; selection is by
    # thumbprint only, never by subject or list position.
    $state = @{ Thumbprints = @('') }

    $loadCertificates = {
        $cboSignCertificate.Items.Clear()
        $state.Thumbprints = @('')
        [void]$cboSignCertificate.Items.Add('(none selected)')
        $cmd = Get-WorkbenchCommand -Name 'Get-CodeSigningCertificateCandidates'
        if (-not $cmd) {
            $txtSignCertDetail.Text = 'The signing service is not loaded, so no certificates can be listed. Signing stays off until it is installed.'
            $cboSignCertificate.SelectedIndex = 0
            $cboSignCertificate.IsEnabled = $false
            $btnSignTest.IsEnabled = $false
            return
        }
        $cboSignCertificate.IsEnabled = $true
        $btnSignTest.IsEnabled = $true
        $candidates = @()
        try { $candidates = @(& $cmd -StoreLocation ([string]$cboSignStore.SelectedItem)) }
        catch { $txtSignCertDetail.Text = 'Certificate enumeration failed: ' + $_.Exception.Message }
        foreach ($c in $candidates) {
            $label = ('{0}  (expires {1})' -f [string]$c.Subject, ([datetime]$c.NotAfter).ToString('yyyy-MM-dd'))
            if (-not [bool]$c.Usable) { $label += '  [unusable]' }
            [void]$cboSignCertificate.Items.Add($label)
            $state.Thumbprints += [string]$c.Thumbprint
        }
        $state['Candidates'] = $candidates
        $idx = [array]::IndexOf($state.Thumbprints, [string]$signing.CertificateThumbprint)
        if ($idx -lt 0) { $idx = 0 }
        $cboSignCertificate.SelectedIndex = $idx
        if ($candidates.Count -eq 0) {
            $txtSignCertDetail.Text = 'No code-signing certificate with a private key was found in ' + [string]$cboSignStore.SelectedItem + '\My.'
        }
    }.GetNewClosure()

    $showDetail = {
        $idx = $cboSignCertificate.SelectedIndex
        if ($idx -le 0) {
            $txtSignCertDetail.Text = 'No certificate selected. Signing cannot run until one is chosen by thumbprint.'
            return
        }
        $thumb = [string]@($state.Thumbprints)[$idx]
        $c = @($state['Candidates'] | Where-Object { [string]$_.Thumbprint -eq $thumb })
        if ($c.Count -eq 0) { return }
        $c = $c[0]
        $lines = @(
            'Subject: ' + [string]$c.Subject
            'Issuer: ' + [string]$c.Issuer
            'Valid: ' + ([datetime]$c.NotBefore).ToString('yyyy-MM-dd') + ' to ' + ([datetime]$c.NotAfter).ToString('yyyy-MM-dd')
            'Thumbprint: ' + $thumb
            'Usable: ' + $(if ([bool]$c.Usable) { 'yes' } else { 'no - ' + [string]$c.Reason })
        )
        $txtSignCertDetail.Text = ($lines -join '   ')
    }.GetNewClosure()

    & $loadCertificates
    & $showDetail

    # The panel builder returns before any of these fire, so each handler
    # takes a closure over the builder's scope rather than relying on it.
    $cboSignStore.Add_SelectionChanged({ & $loadCertificates; & $showDetail }.GetNewClosure())
    $btnSignRefresh.Add_Click({ & $loadCertificates; & $showDetail }.GetNewClosure())
    $cboSignCertificate.Add_SelectionChanged({ & $showDetail }.GetNewClosure())

    $btnSignTest.Add_Click({
        $cmd = Get-WorkbenchCommand -Name 'Test-SigningConfiguration'
        if (-not $cmd) {
            $txtSignTestResult.Text = 'The signing service is not loaded; the test cannot run.'
            return
        }
        $idx = $cboSignCertificate.SelectedIndex
        $thumb = $(if ($idx -gt 0) { [string]@($state.Thumbprints)[$idx] } else { '' })
        $policy = [pscustomobject]@{
            SignDetection         = ($chkSignDetection.IsChecked -eq $true)
            SignRequirements      = ($chkSignRequirements.IsChecked -eq $true)
            SignDeployment        = ($chkSignDeployment.IsChecked -eq $true)
            RequireDetection      = ($chkRequireDetection.IsChecked -eq $true)
            RequireRequirements   = ($chkRequireRequirements.IsChecked -eq $true)
            RequireDeployment     = ($chkRequireDeployment.IsChecked -eq $true)
            CertificateThumbprint = $thumb
            StoreLocation         = [string]$cboSignStore.SelectedItem
            TimestampServer       = ([string]$txtSignTimestamp.Text).Trim()
            TimestampRequired     = ($chkSignTimestampRequired.IsChecked -eq $true)
            HashAlgorithm         = 'SHA256'
        }
        $txtSignTestResult.Text = 'Testing...'
        $btnSignTest.IsEnabled = $false
        try {
            $result = & $cmd -Policy $policy
            $txtSignTestResult.Text = ('Status: {0}. Timestamp verified: {1}. {2}' -f `
                [string]$result.Status, [string]$result.TimestampVerified, [string]$result.Reason)
        }
        catch {
            $txtSignTestResult.Text = 'Test failed: ' + $_.Exception.Message
        }
        finally { $btnSignTest.IsEnabled = $true }
    }.GetNewClosure())

    $prefsRef = $script:Prefs
    $commit = {
        $idx = $cboSignCertificate.SelectedIndex
        $thumb = $(if ($idx -gt 0) { [string]@($state.Thumbprints)[$idx] } else { '' })
        $prefsRef.ScriptSigning.SignDetection       = ($chkSignDetection.IsChecked -eq $true)
        $prefsRef.ScriptSigning.SignRequirements    = ($chkSignRequirements.IsChecked -eq $true)
        $prefsRef.ScriptSigning.SignDeployment      = ($chkSignDeployment.IsChecked -eq $true)
        $prefsRef.ScriptSigning.RequireDetection    = ($chkRequireDetection.IsChecked -eq $true)
        $prefsRef.ScriptSigning.RequireRequirements = ($chkRequireRequirements.IsChecked -eq $true)
        $prefsRef.ScriptSigning.RequireDeployment   = ($chkRequireDeployment.IsChecked -eq $true)
        $prefsRef.ScriptSigning.CertificateThumbprint = $thumb
        $prefsRef.ScriptSigning.StoreLocation       = [string]$cboSignStore.SelectedItem
        $prefsRef.ScriptSigning.TimestampServer     = ([string]$txtSignTimestamp.Text).Trim()
        $prefsRef.ScriptSigning.TimestampRequired   = ($chkSignTimestampRequired.IsChecked -eq $true)
        $prefsRef.ScriptSigning.HashAlgorithm       = 'SHA256'
    }.GetNewClosure()

    return @{ Name = 'Script Signing'; Element = $element; Commit = $commit }
}


# =============================================================================
# Window lifecycle
# =============================================================================
# Handlers built with GetNewClosure run in a dynamic module whose scope chain
# ends at the global scope. A launch that runs this file in a child scope (a
# call from a prompt or the Explorer context menu) leaves every function above
# invisible to them; publishing the functions to the global scope gives all
# launch styles the visibility a -File launch has. Script variables stay out
# of reach either way, so handlers capture the references they need before
# the closure is created.
$script:ScopeProbe = $true
if (-not (Test-Path -LiteralPath 'variable:global:ScopeProbe')) {
    Get-ChildItem -Path 'function:' |
        Where-Object { -not $_.Module -and $_.ScriptBlock.File -eq $PSCommandPath } |
        ForEach-Object { Set-Item -Path ('function:global:' + $_.Name) -Value $_.ScriptBlock }
}
Remove-Variable -Name ScopeProbe -Scope Script -ErrorAction SilentlyContinue

# =============================================================================
# Version display and update affordances
# =============================================================================
$script:AppVersion          = Get-AppVersion
$script:UpdateLatestVersion = $null
$script:UpdateReleaseUrl    = $null

if ($script:AppVersion) {
    $txtAppVersion.Text = "v$($script:AppVersion)"
    $window.Title = "{0}  v{1}" -f $window.Title, $script:AppVersion
}

$lnkUpdateAvailable.Add_Click({ Open-UpdateReleasePage })
$btnUpdateNow.Add_Click({ Invoke-SelfUpdate -Owner $window })

$window.Add_Loaded({
    Add-LogLine -Message ("Loading packagers from: {0}" -f $PackagersRoot)
    Invoke-RefreshGrid
    Add-LogLine -Message ("{0} packager(s) loaded. Ready." -f $script:PackagerData.Count)
    Update-SidebarForDeploymentTarget

    # A zip extracted through Explorer stamps every file with the
    # Mark-of-the-Web; module imports in child processes then fail while
    # the caller keeps running, surfacing as unknown-command errors
    # mid-stage. Detect and offer to clear it before anything runs.
    $blocked = @(Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue })
    if ($blocked.Count -gt 0) {
        Add-LogLine -Message ("{0} file(s) carry the Mark-of-the-Web (downloaded-file block)." -f $blocked.Count)
        $answer = Show-ThemedMessage -Owner $window -Title 'Blocked Files Detected' `
            -Message ("{0} file(s) in this folder are blocked because they came from a downloaded zip. Packager runs will fail with unknown-command errors until the block is cleared.`n`nUnblock all files now?" -f $blocked.Count) `
            -Buttons YesNo -Icon Warning
        if ($answer -eq 'Yes') {
            $blocked | Unblock-File -ErrorAction SilentlyContinue
            Add-LogLine -Message ("Unblocked {0} file(s)." -f $blocked.Count)
        }
        else {
            Add-LogLine -Message "Blocked files left in place. Clear them manually with: Get-ChildItem <app folder> -Recurse | Unblock-File"
        }
    }

    if (Test-FirstRunWizardNeeded -Prefs $script:Prefs) {
        Show-FirstRunWizard -Owner $window
    }

    Start-UpdateCheck
})

$window.Add_Closing({
    if ($script:UpdateCheckTimer) { $script:UpdateCheckTimer.Stop(); $script:UpdateCheckTimer = $null }

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
