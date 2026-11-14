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
    Version    : 2026.11.14.0139
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
    [switch]$Force,

    # Headless CVE monitor (with -BatchMode): refreshes the latest and the
    # ConfigMgr versions, looks up the CVEs of every application with an
    # update, writes the report and notifies. -Apps limits the set; without
    # it every packager runs.
    [switch]$CveReport,
    [switch]$SkipLatest,
    [switch]$SkipMecm,
    [switch]$SkipCve,
    [string]$ReportFolder,
    [ValidateSet('Html', 'Csv', 'Json', 'Xlsx', 'Markdown')][string[]]$ReportFormats = @('Html', 'Csv', 'Json', 'Xlsx', 'Markdown'),
    [string]$WebhookUrl,
    [string]$DropFolder,
    [switch]$NotifyAlways,
    [int]$KeepReportDays = 0
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
# second import of an already-loaded module is a no-op. The WSUS publisher
# loads only here and in the background runspace: packager children never
# publish.
foreach ($workbenchModule in @('AppPackagerWorkbench.psm1', 'AppPackagerSigning.psm1', 'AppPackagerWsus.psm1', 'AppPackagerOneClick.psm1', 'AppPackagerVulnerability.psm1', 'AppPackagerDiscovery.psm1')) {
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

function Get-DeploymentTargetNames {
    return @('MECM', 'MECMAndIntune', 'IntuneOnly', 'MECMAndWSUS', 'WSUSOnly')
}

function ConvertTo-DeploymentTargetName {
    # The nearest single target for the code paths that still read one; a
    # set of all three or WSUS+Intune has no name and maps to its ConfigMgr
    # or WSUS part.
    param([Parameter(Mandatory)]$Destinations)
    $cm = [bool]$Destinations.ConfigMgr; $wsus = [bool]$Destinations.WSUS; $intune = [bool]$Destinations.Intune
    if ($cm -and $wsus) { return 'MECMAndWSUS' }
    if ($cm -and $intune) { return 'MECMAndIntune' }
    if ($cm) { return 'MECM' }
    if ($wsus) { return 'WSUSOnly' }
    if ($intune) { return 'IntuneOnly' }
    return 'MECM'
}

function Test-DeploymentTargetSkipsSite {
    # Intune-only and WSUS-only runs stop at Stage: no site connection, no
    # share copy, no ConfigMgr application.
    param([AllowEmptyString()][string]$DeploymentTarget)
    return ([string]$DeploymentTarget -in @('IntuneOnly', 'WSUSOnly'))
}

function Test-DeploymentTargetPublishesToWsus {
    param([AllowEmptyString()][string]$DeploymentTarget)
    return ([string]$DeploymentTarget -in @('MECMAndWSUS', 'WSUSOnly'))
}

function Get-PublishRowOutcome {
    # Row status after one application's publish run. Kind is Success,
    # Skipped, NotSupported or Failed. An Intune or WSUS run succeeds only
    # when its publish succeeds; a ConfigMgr run that also publishes to WSUS
    # carries the WSUS result as a status suffix.
    param([Parameter(Mandatory)]$Result, [AllowEmptyString()][string]$DeploymentTarget)

    $wsus = if ($Result.PSObject.Properties['WsusPublish']) { $Result.WsusPublish } else { $null }
    $intune = if ($Result.PSObject.Properties['IntunePublish']) { $Result.IntunePublish } else { $null }
    $wsusRefused = [bool]($wsus -and $wsus.PSObject.Properties['Refused'] -and $wsus.Refused)
    $wsusWarned = [bool]($wsus -and $wsus.Ok -and $wsus.PSObject.Properties['Warnings'] -and @($wsus.Warnings).Count -gt 0)
    $wsusSuffix = ''
    if ($wsusRefused) { $wsusSuffix = ', WSUS not supported' }
    elseif ($wsus -and -not $wsus.Ok) { $wsusSuffix = ', WSUS failed' }
    elseif ($wsusWarned) { $wsusSuffix = ', WSUS warning' }

    $siteFree = Test-DeploymentTargetSkipsSite -DeploymentTarget $DeploymentTarget
    $make = { param($Kind, $Status, $Published) [pscustomobject]@{ Kind = $Kind; Status = $Status; Published = [bool]$Published } }

    if ($Result.PSObject.Properties['PackageOutcome'] -and $Result.PackageOutcome -in @('Skipped', 'Canceled')) {
        return (& $make 'Skipped' (([string]$Result.PackageOutcome) + $wsusSuffix) $false)
    }
    if ($Result.ExitCode -ne 0) {
        return (& $make 'Failed' $(if ($siteFree) { 'Publish error' } else { 'Package error' }) $false)
    }
    if ($DeploymentTarget -eq 'WSUSOnly') {
        if ($wsus -and $wsus.Ok) { return (& $make 'Success' $(if ($wsusWarned) { 'Published, warning' } else { 'Published' }) $true) }
        if ($wsusRefused) { return (& $make 'NotSupported' 'WSUS: not supported' $false) }
        return (& $make 'Failed' 'Publish error' $false)
    }
    if ($DeploymentTarget -eq 'IntuneOnly') {
        if ($intune -and $intune.Ok) { return (& $make 'Success' 'Published' $true) }
        return (& $make 'Failed' 'Publish error' $false)
    }
    return (& $make 'Success' ('Packaged' + $wsusSuffix) $false)
}

function Get-WsusClassificationNameList {
    # The WSUS module owns the classification names; a host where it failed
    # to load still reads every other preference section.
    if (Get-Command -Name Get-WsusClassificationNames -ErrorAction SilentlyContinue) {
        return @(Get-WsusClassificationNames)
    }
    return @('Updates')
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
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'A key that an earlier build saved in plain text is protected on read.')]
    param()
    $defaults = [pscustomobject]@{
        SiteCode             = "MCM"
        ProviderMachineName  = ""
        FileShareRoot        = "\\fileserver\sccm$"
        ContentLayout        = "Nested"
        DownloadRoot         = "C:\temp\ap"
        EstimatedRuntimeMins = 15
        MaximumRuntimeMins   = 30
        ContentFallback      = "Allow"
        SlowNetworkDeploymentMode = "Download"
        CompanyName          = ""
        M365Channel          = "MonthlyEnterprise"
        M365DeployMode       = "Managed"
        M365ExcludeApps      = @('Groove','Lync','OneDrive','Teams')
        # The Office Deployment Tool values that the Microsoft 365 packagers write. A value that the Office Deployment
        # Tool documents with a default of 0 or off stays off. OfficeMgmtCOM, Updates, FORCEAPPSHUTDOWN and RemoveMSI
        # keep the values that the packagers wrote before these were settings.
        M365Options          = [pscustomobject]@{
            Languages               = @('en-us')
            OfficeMgmtCom           = $true
            UpdatesEnabled          = $true
            ForceAppShutdown        = $true
            RemoveMsi               = $true
            MigrateArch             = $false
            SharedComputerLicensing = $false
            DeviceBasedLicensing    = $false
        }
        # A configuration.xml that the user imported replaces the generated one. The file is stored in Packagers.
        M365CustomConfig     = [pscustomobject]@{
            Enabled   = $false
            FileName  = ''
            SizeBytes = 0
        }
        SSMSInstallOptions   = [pscustomobject]@{
            UIMode             = "Quiet"
            DownloadThenInstall = $true
            NoUpdateInstaller  = $false
            IncludeRecommended = $false
            IncludeOptional    = $false
            RemoveOos          = $true
            ForceClose         = $false
            InstallPath        = ""
            Workloads          = @()
            Languages          = @()
            NoCache            = $false
        }
        DBeaverInstallOptions = [pscustomobject]@{
            InstallScope = "System"
            DisableAI    = $false
        }
        AdobeReaderInstallOptions = [pscustomobject]@{
            Edition                   = 'English'
            Languages                 = @()
            AcceptEula                = $false
            DisableDesktopShortcut    = $false
            DisableChromeExtension    = $false
            DisableBrowserIntegration = $false
            UpdateMode                = ''
        }
        BeyondCompareKeyFile = ""
        BeyondCompareOptions = [pscustomobject]@{
            DisableRemoteProfiles = $false
            DisableSavedPasswords = $false
        }
        LocalSourceFolders   = [pscustomobject]@{}
        HiddenApplications   = @()
        DiscoveryDismissed   = @()
        DiscoveryEnabled     = @()
        # Empty means the server and the CM_<site> database come from the site.
        DiscoverySqlServer   = ''
        DiscoverySqlInstance = ''
        # 0 means the default port of the instance.
        DiscoverySqlPort     = 0
        DiscoverySqlDatabase = ''
        FirstRunCompleted    = $false
        IncludeVersionInTitle = $false
        # Strips the release version from a vendor product name; applies only
        # while IncludeVersionInTitle is off.
        RemoveVersionFromTitle = $false
        # The systems the sidebar publishes to. Setup and Options set them;
        # a button of a system not in use stays disabled whatever else is set.
        Systems              = [pscustomobject]@{
            ConfigMgr = $true
            Intune    = $false
            Wsus      = $false
        }
        AppFlow              = [pscustomobject]@{
            Tracked          = @()
            Action           = 'Report'
            CadenceOverrides = [pscustomobject]@{}
            ForceOnLaunch    = $false
            # Destinations per tracked application; a row without an entry
            # uses DefaultDestinations. OnExisting decides a ConfigMgr
            # application that already exists without a prompt.
            Destinations        = [pscustomobject]@{}
            DefaultDestinations = [pscustomobject]@{ ConfigMgr = $true; WSUS = $false; Intune = $false }
            OnExisting          = 'Skip'
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
        # The token that the GitHub-backed packagers send. DPAPI-protected for the current Windows user; never plaintext.
        GitHub               = [pscustomobject]@{
            TokenProtected = ''
        }
        Intune               = [pscustomobject]@{
            # One Click destination. MECM = ConfigMgr app; MECMAndIntune =
            # ConfigMgr app + Graph publish; IntuneOnly = stage + .intunewin +
            # Graph publish, no site touch; MECMAndWSUS = ConfigMgr app + WSUS
            # publish; WSUSOnly = stage + WSUS publish, no site touch. The
            # sidebar publish buttons set their own destination per run.
            DeploymentTarget      = 'MECM'
            PublishToIntune       = $false
            TenantId              = ''
            ClientId              = ''
            # DPAPI-protected (ConvertFrom-SecureString); never plaintext.
            ClientSecretProtected = ''
        }
        Wsus                 = [pscustomobject]@{
            ServerName        = ''
            PortNumber        = 8530
            UseSsl            = $false
            Classification    = 'Updates'
            ApprovalGroup     = ''
            DeclineSuperseded = $false
        }
        DeploymentConditions = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
        CommandOverrides = [pscustomobject]@{
            Apps = [pscustomobject]@{}
        }
        WorkbenchDataRoot = ''
        Vulnerability = [pscustomobject]@{
            # NVD allows 5 requests per 30 seconds without a key, 50 with one.
            # DPAPI-protected (ConvertFrom-SecureString); never plaintext.
            NvdApiKeyProtected = ''
            CacheTtlMinutes   = 360
            LookupAfterCheckCm = $false
        }
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

        if ($null -ne $data.Vulnerability) {
            $vu = $data.Vulnerability
            if ($null -ne $vu.NvdApiKeyProtected) { $defaults.Vulnerability.NvdApiKeyProtected = [string]$vu.NvdApiKeyProtected }
            elseif ($null -ne $vu.NvdApiKey -and -not [string]::IsNullOrWhiteSpace([string]$vu.NvdApiKey)) {
                # A key saved in plain text by an earlier build is protected on read.
                try { $defaults.Vulnerability.NvdApiKeyProtected = (ConvertTo-SecureString -String ([string]$vu.NvdApiKey).Trim() -AsPlainText -Force | ConvertFrom-SecureString) } catch { }
            }
            $ttl = 0
            if ([int]::TryParse([string]$vu.CacheTtlMinutes, [ref]$ttl) -and $ttl -ge 0) { $defaults.Vulnerability.CacheTtlMinutes = $ttl }
            if ($null -ne $vu.LookupAfterCheckCm) { try { $defaults.Vulnerability.LookupAfterCheckCm = [bool]$vu.LookupAfterCheckCm } catch { <# Reading a boolean setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
        }

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
        if ([string]$data.ContentFallback -in @('Allow','Deny')) { $defaults.ContentFallback = [string]$data.ContentFallback }
        if ([string]$data.SlowNetworkDeploymentMode -in @('Download','DoNothing')) { $defaults.SlowNetworkDeploymentMode = [string]$data.SlowNetworkDeploymentMode }
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
            # Filter to only documented ExcludeApp IDs. Unknown values, such as the
            # retired Bing ID, are dropped silently.
            $validExcludes = @('Access','Excel','Groove','Lync','OneDrive','OneNote','Outlook','OutlookForWindows','PowerPoint','Publisher','Teams','Word')
            $defaults.M365ExcludeApps = @($data.M365ExcludeApps | Where-Object { $_ -in $validExcludes })
        }

        if ($null -ne $data.M365Options) {
            $mo = $data.M365Options
            foreach ($prop in @('OfficeMgmtCom','UpdatesEnabled','ForceAppShutdown','RemoveMsi','MigrateArch','SharedComputerLicensing','DeviceBasedLicensing')) {
                if ($null -ne $mo.$prop) { try { $defaults.M365Options.$prop = [bool]$mo.$prop } catch { <# Reading a boolean setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            }
            if ($null -ne $mo.Languages) {
                $codes = @($mo.Languages | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ -match '^[a-z]{2,3}(-[a-z0-9]{2,8}){1,2}$' })
                if ($codes.Count -gt 0) { $defaults.M365Options.Languages = $codes }
            }
        }
        if ($null -ne $data.M365CustomConfig) {
            $mc = $data.M365CustomConfig
            if ($null -ne $mc.Enabled) { try { $defaults.M365CustomConfig.Enabled = [bool]$mc.Enabled } catch { <# Reading a boolean setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            if ($null -ne $mc.FileName) { $defaults.M365CustomConfig.FileName = [string]$mc.FileName }
            $sizeBytes = 0
            if ([long]::TryParse([string]$mc.SizeBytes, [ref]$sizeBytes) -and $sizeBytes -ge 0) { $defaults.M365CustomConfig.SizeBytes = $sizeBytes }
        }

        if ($null -ne $data.SSMSInstallOptions) {
            $ssms = $data.SSMSInstallOptions
            if ($null -ne $ssms.UIMode) {
                $modeRaw = [string]$ssms.UIMode
                if ($modeRaw -in @('Quiet','Passive')) { $defaults.SSMSInstallOptions.UIMode = $modeRaw }
            }
            foreach ($prop in @('DownloadThenInstall','NoUpdateInstaller','IncludeRecommended','IncludeOptional','RemoveOos','ForceClose','NoCache')) {
                if ($null -ne $ssms.$prop) {
                    try { $defaults.SSMSInstallOptions.$prop = [bool]$ssms.$prop } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
                }
            }
            if ($null -ne $ssms.InstallPath) { $defaults.SSMSInstallOptions.InstallPath = [string]$ssms.InstallPath }
            if ($null -ne $ssms.Workloads) { $defaults.SSMSInstallOptions.Workloads = @($ssms.Workloads | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -match '^Microsoft\.SqlServer\.Workload\.SSMS\.[A-Za-z]+$' }) }
            if ($null -ne $ssms.Languages) { $defaults.SSMSInstallOptions.Languages = @($ssms.Languages | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ -match '^[a-z]{2}-[a-z]{2}$' }) }
        }

        if ($null -ne $data.DBeaverInstallOptions) {
            $dbv = $data.DBeaverInstallOptions
            if ([string]$dbv.InstallScope -in @('System','User')) { $defaults.DBeaverInstallOptions.InstallScope = [string]$dbv.InstallScope }
            if ($null -ne $dbv.DisableAI) {
                try { $defaults.DBeaverInstallOptions.DisableAI = [bool]$dbv.DisableAI } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
            }
        }

        if ($null -ne $data.AdobeReaderInstallOptions) {
            $ar = $data.AdobeReaderInstallOptions
            if ([string]$ar.Edition -in @('English', 'MUI')) { $defaults.AdobeReaderInstallOptions.Edition = [string]$ar.Edition }
            if ($null -ne $ar.Languages) { $defaults.AdobeReaderInstallOptions.Languages = @($ar.Languages | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^([a-z]{2}_[A-Z]{2}|All)$' }) }
            foreach ($prop in @('AcceptEula','DisableDesktopShortcut','DisableChromeExtension','DisableBrowserIntegration')) {
                if ($null -ne $ar.$prop) { try { $defaults.AdobeReaderInstallOptions.$prop = [bool]$ar.$prop } catch { <# Reading a boolean setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            }
            if ([string]$ar.UpdateMode -in @('', '0', '2', '3', '4')) { $defaults.AdobeReaderInstallOptions.UpdateMode = [string]$ar.UpdateMode }
        }

        if ($null -ne $data.BeyondCompareKeyFile)  { $defaults.BeyondCompareKeyFile = [string]$data.BeyondCompareKeyFile }
        if ($null -ne $data.BeyondCompareOptions) {
            foreach ($prop in @('DisableRemoteProfiles','DisableSavedPasswords')) {
                if ($null -ne $data.BeyondCompareOptions.$prop) { try { $defaults.BeyondCompareOptions.$prop = [bool]$data.BeyondCompareOptions.$prop } catch { <# Reading a boolean setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            }
        }
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
        foreach ($listName in 'DiscoveryDismissed', 'DiscoveryEnabled') {
            $stored = $data.PSObject.Properties[$listName]
            if (-not $stored -or $null -eq $stored.Value) { continue }
            $kept = New-Object System.Collections.Generic.List[string]
            $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($entry in @($stored.Value)) {
                $name = ([string]$entry).Trim()
                if ($name -notmatch '^package-.+\.(?:ps1|notps1)$') { continue }
                if ($seen.Add($name)) { $kept.Add($name) }
            }
            $defaults.$listName = $kept.ToArray()
        }
        if ($null -ne $data.DiscoverySqlServer)   { $defaults.DiscoverySqlServer   = ([string]$data.DiscoverySqlServer).Trim() }
        if ($null -ne $data.DiscoverySqlDatabase) { $defaults.DiscoverySqlDatabase = ([string]$data.DiscoverySqlDatabase).Trim() }
        if ($null -ne $data.DiscoverySqlInstance) { $defaults.DiscoverySqlInstance = ([string]$data.DiscoverySqlInstance).Trim() }
        $sqlPort = 0
        if ([int]::TryParse([string]$data.DiscoverySqlPort, [ref]$sqlPort) -and $sqlPort -ge 1 -and $sqlPort -le 65535) { $defaults.DiscoverySqlPort = $sqlPort }
        $defaults.FirstRunCompleted = Resolve-FirstRunCompleted -StoredValue $data.FirstRunCompleted -PreferencesFileExisted $true
        if ($null -ne $data.IncludeVersionInTitle) {
            try { $defaults.IncludeVersionInTitle = [bool]$data.IncludeVersionInTitle } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
        }
        if ($null -ne $data.RemoveVersionFromTitle) {
            try { $defaults.RemoveVersionFromTitle = [bool]$data.RemoveVersionFromTitle } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
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
                try { $defaults.AppFlow.ForceOnLaunch = [bool]$af.ForceOnLaunch } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
            }

            if ($null -ne $af.Destinations) {
                $destProps = [ordered]@{}
                foreach ($prop in $af.Destinations.PSObject.Properties) {
                    if ($prop.Name -notmatch '^package-' -or $null -eq $prop.Value) { continue }
                    $destProps[$prop.Name] = [pscustomobject]@{
                        ConfigMgr = [bool]($prop.Value.PSObject.Properties['ConfigMgr'] -and $prop.Value.ConfigMgr)
                        WSUS      = [bool]($prop.Value.PSObject.Properties['WSUS'] -and $prop.Value.WSUS)
                        Intune    = [bool]($prop.Value.PSObject.Properties['Intune'] -and $prop.Value.Intune)
                    }
                }
                $defaults.AppFlow.Destinations = [pscustomobject]$destProps
            }
            if ($null -ne $af.DefaultDestinations) {
                foreach ($name in @('ConfigMgr', 'WSUS', 'Intune')) {
                    if ($null -ne $af.DefaultDestinations.PSObject.Properties[$name]) { try { $defaults.AppFlow.DefaultDestinations.$name = [bool]$af.DefaultDestinations.$name } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
                }
            }
            elseif ($null -ne $data.Intune -and $null -ne $data.Intune.DeploymentTarget) {
                # A preferences file from before the per-row destinations
                # carries one saved target.
                $target = [string]$data.Intune.DeploymentTarget
                $defaults.AppFlow.DefaultDestinations.ConfigMgr = ($target -notin @('IntuneOnly', 'WSUSOnly'))
                $defaults.AppFlow.DefaultDestinations.WSUS      = ($target -in @('MECMAndWSUS', 'WSUSOnly'))
                $defaults.AppFlow.DefaultDestinations.Intune    = ($target -in @('MECMAndIntune', 'IntuneOnly'))
            }
            if ([string]$af.OnExisting -in @('Skip', 'Overwrite')) { $defaults.AppFlow.OnExisting = [string]$af.OnExisting }
        }

        # ContentDistribution: auto-distribute-to-DP-group settings.
        if ($null -ne $data.ContentDistribution) {
            $cd = $data.ContentDistribution
            if ($null -ne $cd.AutoDistribute) {
                try { $defaults.ContentDistribution.AutoDistribute = [bool]$cd.AutoDistribute } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
            }
            if ($null -ne $cd.DPGroupName) {
                $defaults.ContentDistribution.DPGroupName = [string]$cd.DPGroupName
            }
            if ($null -ne $cd.DeployToTestCollection) {
                try { $defaults.ContentDistribution.DeployToTestCollection = [bool]$cd.DeployToTestCollection } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
            }
            if ($null -ne $cd.TestCollectionName) {
                $defaults.ContentDistribution.TestCollectionName = [string]$cd.TestCollectionName
            }
            if ($null -ne $cd.CreateTestCollectionIfMissing) {
                try { $defaults.ContentDistribution.CreateTestCollectionIfMissing = [bool]$cd.CreateTestCollectionIfMissing } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ }
            }
        }

        if ($null -ne $data.GitHub -and $null -ne $data.GitHub.TokenProtected) { $defaults.GitHub.TokenProtected = [string]$data.GitHub.TokenProtected }

        # Intune: .intunewin production and Graph publishing during Package.
        if ($null -ne $data.Intune) {
            if ($null -ne $data.Intune.PublishToIntune) { try { $defaults.Intune.PublishToIntune = [bool]$data.Intune.PublishToIntune } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            if ([string]$data.Intune.DeploymentTarget -in @('MECM', 'MECMAndIntune', 'IntuneOnly', 'MECMAndWSUS', 'WSUSOnly')) {
                $defaults.Intune.DeploymentTarget = [string]$data.Intune.DeploymentTarget
            }
            elseif ($defaults.Intune.PublishToIntune) {
                # Pre-1.4.0.14 prefs expressed publishing as a bare toggle.
                $defaults.Intune.DeploymentTarget = 'MECMAndIntune'
            }
            if ($null -ne $data.Intune.TenantId)              { $defaults.Intune.TenantId              = [string]$data.Intune.TenantId }
            if ($null -ne $data.Intune.ClientId)              { $defaults.Intune.ClientId              = [string]$data.Intune.ClientId }
            if ($null -ne $data.Intune.ClientSecretProtected) { $defaults.Intune.ClientSecretProtected = [string]$data.Intune.ClientSecretProtected }
            # A stale Intune toggle must not publish to Intune from a WSUS target.
            if ([string]$defaults.Intune.DeploymentTarget -in @('MECMAndWSUS', 'WSUSOnly')) {
                $defaults.Intune.PublishToIntune = $false
            }
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
                if ($null -ne $sign.$flag) { try { $defaults.ScriptSigning.$flag = [bool]$sign.$flag } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
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
            if ($null -ne $cm.Found)           { try { $stored.Found = [bool]$cm.Found } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
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
            if ($null -ne $sz.Found)           { try { $stored.Found = [bool]$sz.Found } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
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
            if ($null -ne $iw.Found)          { try { $stored.Found = [bool]$iw.Found } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            if ($null -ne $iw.DisplayVersion) { $stored.DisplayVersion = [string]$iw.DisplayVersion }
            if ($null -ne $iw.ExePath)        { $stored.ExePath        = [string]$iw.ExePath }
            if ($null -ne $iw.DetectedAt)     { $stored.DetectedAt     = [string]$iw.DetectedAt }
            $defaults.DetectedTools.IntuneWinAppUtil = $stored
        }

        # Systems: a preferences file from before this key infers the
        # systems from the settings it holds.
        if ($null -ne $data.Systems) {
            foreach ($system in @('ConfigMgr', 'Intune', 'Wsus')) {
                if ($null -ne $data.Systems.$system) { try { $defaults.Systems.$system = [bool]$data.Systems.$system } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            }
        } else {
            $defaults.Systems.Intune = -not [string]::IsNullOrWhiteSpace([string]$defaults.Intune.TenantId)
            $defaults.Systems.Wsus   = ($null -ne $data.Wsus -and -not [string]::IsNullOrWhiteSpace([string]$data.Wsus.ServerName))
        }

        # Wsus: local publishing to a WSUS server during Package. Unknown or
        # out-of-range values keep the defaults above. Parsed last: a failure
        # here must not discard the sections before it.
        if ($null -ne $data.Wsus) {
            $ws = $data.Wsus
            if ($null -ne $ws.ServerName) { $defaults.Wsus.ServerName = ([string]$ws.ServerName).Trim() }
            $port = 0
            if ([int]::TryParse([string]$ws.PortNumber, [ref]$port) -and $port -ge 1 -and $port -le 65535) { $defaults.Wsus.PortNumber = $port }
            if ($null -ne $ws.UseSsl) { try { $defaults.Wsus.UseSsl = [bool]$ws.UseSsl } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
            $classNames = @('Updates')
            if (Get-Command -Name Get-WsusClassificationNames -ErrorAction SilentlyContinue) { $classNames = @(Get-WsusClassificationNames) }
            if ([string]$ws.Classification -in $classNames) { $defaults.Wsus.Classification = [string]$ws.Classification }
            if ($null -ne $ws.ApprovalGroup) { $defaults.Wsus.ApprovalGroup = ([string]$ws.ApprovalGroup).Trim() }
            if ($null -ne $ws.DeclineSuperseded) { try { $defaults.Wsus.DeclineSuperseded = [bool]$ws.DeclineSuperseded } catch { <# Reading a boolean or numeric setting from the preferences file; a bad value keeps the default. #> $null = $_ } }
        }
    }
    catch { }

    return $defaults
}

function Set-DeploymentTypeContentEnvironment {
    # Packager child processes and the background runspace inherit this
    # process environment; the module reads these two values when it adds a
    # ConfigMgr deployment type.
    param([Parameter(Mandatory)]$Prefs)
    $env:APP_PACKAGER_DT_CONTENT_FALLBACK  = $(if ([string]$Prefs.ContentFallback -eq 'Deny') { 'Deny' } else { 'Allow' })
    $env:APP_PACKAGER_DT_SLOW_NETWORK_MODE = $(if ([string]$Prefs.SlowNetworkDeploymentMode -eq 'DoNothing') { 'DoNothing' } else { 'Download' })
}

function Save-Preferences {
    param([Parameter(Mandatory)][pscustomobject]$Prefs)

    $path = Get-PreferencesPath
    $json = $Prefs | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
    Set-DeploymentTypeContentEnvironment -Prefs $Prefs

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
        $pkgPrefs["AdobeReaderInstallOptions"] = $Prefs.AdobeReaderInstallOptions
        $pkgPrefs["BeyondCompareKeyFile"] = [string]$Prefs.BeyondCompareKeyFile
        $pkgPrefs["BeyondCompareOptions"] = $Prefs.BeyondCompareOptions
        $pkgPrefs["M365Options"] = $Prefs.M365Options
        # The packagers read the imported file from this path, and only while the box is selected and the file exists.
        $customPath = Join-Path (Join-Path $PSScriptRoot "Packagers") "m365-configuration.user.xml"
        $useCustom = ([bool]$Prefs.M365CustomConfig.Enabled -and (Test-Path -LiteralPath $customPath -PathType Leaf))
        $pkgPrefs["M365UseCustomConfig"] = $useCustom
        $pkgPrefs["M365CustomConfigPath"] = $(if ($useCustom) { $customPath } else { "" })
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
                try { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop } catch { <# Reading one uninstall registry key; an unreadable key is skipped. #> $null = $_ }
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
                try { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop } catch { <# Reading one uninstall registry key; an unreadable key is skipped. #> $null = $_ }
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
    # token: GITHUB_TOKEN, GH_TOKEN, the token saved in Options, Tools, then the GitHub CLI's stored login.
    # Anonymous calls are limited to 60 per hour per address, which a version
    # sweep over the catalog exceeds. The rate-limit probe is one request with
    # a short timeout; when it cannot complete the source is still reported.
    $source = ''
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)) { $source = 'GITHUB_TOKEN environment variable' }
    elseif (-not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)) { $source = 'GH_TOKEN environment variable' }
    elseif (-not [string]::IsNullOrWhiteSpace((Get-StoredGitHubToken))) { $source = 'token saved in Options, Tools' }
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
                catch { <# Reading the GitHub account name is optional display text. #> $null = $_ }
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
    catch { <# Reading the GitHub rate limit is optional display information. #> $null = $_ }

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
            } catch { <# Reading the file version of a tool is optional detail. #> $null = $_ }
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
} catch { <# Refreshing tool detection at launch is best-effort; the earlier detection stays. #> $null = $_ }

if ([string]::IsNullOrWhiteSpace($script:Prefs.CompanyName)) {
    $pkgPrefsPath = Join-Path (Join-Path $PSScriptRoot "Packagers") "packager-preferences.json"
    if (Test-Path -LiteralPath $pkgPrefsPath) {
        try {
            $pkgData = Get-Content -LiteralPath $pkgPrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($pkgData.CompanyName) { $script:Prefs.CompanyName = [string]$pkgData.CompanyName }
        }
        catch { <# Reading the company name from the optional packager preferences file is a best-effort fallback. #> $null = $_ }
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
Set-DeploymentTypeContentEnvironment -Prefs $script:Prefs

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
        SecurityAdvisoryUrl = $null
        Description       = $null
        UpdateCadenceDays = $null
        SupportsVariants  = @()
        SupportsInstallModes = @()
        LocalSource       = $false
        LocalSourceRequired = $false
        WsusSupport       = $null
        WsusSupportReason = $null
        ArpName           = $null
        ArpVersion        = $null
        ArpArch           = $null
    }

    $lines = @([System.IO.File]::ReadLines($Path) | Select-Object -First 200)

    # Every key test below starts with the same prefix; one match per line
    # rules out the lines that carry no key, which are nearly all of them.
    $keyLine = New-Object System.Text.RegularExpressions.Regex('^\s*(?:#\s*)?(?:Vendor|App|CMName|VendorUrl|CPE|ReleaseNotesUrl|DownloadPageUrl|SecurityAdvisoryUrl|SupportsVariants|SupportsInstallModes|LocalSource|WsusSupport|ArpName|ArpVersion|ArpArch|UpdateCadenceDays|Application)\s*:', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $inSynopsis = $false
    foreach ($line in $lines) {
        $l = $line.TrimStart([char]0xFEFF)

        if ($keyLine.IsMatch($l)) {
        if (-not $meta.Vendor    -and $l -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$')    { $meta.Vendor    = $Matches[1].Trim(); continue }
        if (-not $meta.App       -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')       { $meta.App       = $Matches[1].Trim(); continue }
        if (-not $meta.CMName    -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$')    { $meta.CMName    = $Matches[1].Trim(); continue }
        if (-not $meta.VendorUrl       -and $l -match '^\s*(?:#\s*)?VendorUrl\s*:\s*(.+?)\s*$')       { $meta.VendorUrl       = $Matches[1].Trim(); continue }
        if (-not $meta.CPE             -and $l -match '^\s*(?:#\s*)?CPE\s*:\s*(.+?)\s*$')             { $meta.CPE             = $Matches[1].Trim(); continue }
        if (-not $meta.ReleaseNotesUrl -and $l -match '^\s*(?:#\s*)?ReleaseNotesUrl\s*:\s*(.+?)\s*$') { $meta.ReleaseNotesUrl = $Matches[1].Trim(); continue }
        if (-not $meta.DownloadPageUrl -and $l -match '^\s*(?:#\s*)?DownloadPageUrl\s*:\s*(.+?)\s*$') { $meta.DownloadPageUrl = $Matches[1].Trim(); continue }
        if (-not $meta.SecurityAdvisoryUrl -and $l -match '^\s*(?:#\s*)?SecurityAdvisoryUrl\s*:\s*(.+?)\s*$') { $meta.SecurityAdvisoryUrl = $Matches[1].Trim(); continue }
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
        if ($null -eq $meta.WsusSupport -and $l -match '^\s*(?:#\s*)?WsusSupport\s*:\s*(Yes|No)\s*(?:\(\s*([A-Za-z|]+)\s*\))?\s*$') {
            $meta.WsusSupport = $Matches[1]
            if ($Matches[2]) { $meta.WsusSupportReason = $Matches[2] }
            continue
        }
        if ($null -eq $meta.ArpName -and $l -match '^\s*(?:#\s*)?ArpName\s*:\s*(.+?)\s*$') { $meta.ArpName = $Matches[1].Trim(); continue }
        if ($null -eq $meta.ArpVersion -and $l -match '^\s*(?:#\s*)?ArpVersion\s*:\s*(.+?)\s*$') { $meta.ArpVersion = $Matches[1].Trim(); continue }
        if ($null -eq $meta.ArpArch -and $l -match '^\s*(?:#\s*)?ArpArch\s*:\s*(x64|x86|arm64)\s*$') { $meta.ArpArch = $Matches[1].ToLowerInvariant(); continue }
        if ($null -eq $meta.UpdateCadenceDays -and $l -match '^\s*(?:#\s*)?UpdateCadenceDays\s*:\s*(\d+)\s*$') {
            $days = [int]$Matches[1]
            if ($days -ge 1) { $meta.UpdateCadenceDays = $days }
            continue
        }
        if (-not $meta.App             -and $l -match '^\s*(?:#\s*)?Application\s*:\s*(.+?)\s*$')     { $meta.App             = $Matches[1].Trim(); continue }
        }

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
        SecurityAdvisoryUrl = $meta.SecurityAdvisoryUrl
        Description       = $meta.Description
        UpdateCadenceDays = $meta.UpdateCadenceDays
        SupportsVariants  = @($meta.SupportsVariants)
        SupportsInstallModes = @($meta.SupportsInstallModes)
        LocalSource       = [bool]$meta.LocalSource
        LocalSourceRequired = [bool]$meta.LocalSourceRequired
        WsusSupport       = $meta.WsusSupport
        WsusSupportReason = $meta.WsusSupportReason
        ArpName           = $meta.ArpName
        ArpVersion        = $meta.ArpVersion
        ArpArch           = $meta.ArpArch
        Script            = (Split-Path -Leaf $Path)
        FullPath          = $Path
    }
}

function Get-Packagers {
    <#
        Reading and parsing every packager header is the cost of a call, and the
        Options window makes several. The parsed rows are kept per folder and
        reused while no packager file has been added, removed or changed;
        every caller gets its own copies, since the grid edits its rows.
    #>
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) { return @() }

    $files = Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop |
        Where-Object { $_.Name -match '^package-.*\.(?:ps1|notps1)$' } |
        Sort-Object Name

    $signature = (@($files | ForEach-Object { '{0}|{1}|{2}' -f $_.Name, $_.Length, $_.LastWriteTimeUtc.Ticks }) -join ';')
    if (-not $script:PackagerCatalogCache) { $script:PackagerCatalogCache = @{} }
    $cacheKey = $Root.TrimEnd('\').ToLowerInvariant()
    $cached = $script:PackagerCatalogCache[$cacheKey]
    if ($cached -and $cached.Signature -ceq $signature) {
        return @($cached.Rows | ForEach-Object { $_.PSObject.Copy() })
    }

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
                CPE               = [string]$m.CPE
                SecurityAdvisoryUrl = [string]$m.SecurityAdvisoryUrl
                ArpName           = [string]$m.ArpName
                ArpVersion        = [string]$m.ArpVersion
                ArpArch           = [string]$m.ArpArch
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
                CPE               = ""
                SecurityAdvisoryUrl = ""
                ArpName           = ""
                ArpVersion        = ""
                ArpArch           = ""
                CurrentVersion    = ""
                LatestVersion     = ""
                Status            = ("Read error: " + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            })
        }
    }
    $script:PackagerCatalogCache[$cacheKey] = @{ Signature = $signature; Rows = @($items | ForEach-Object { $_.PSObject.Copy() }) }
    return $items
}

# =============================================================================
# Discover unmanaged software
# =============================================================================
$script:DiscoveryScanTimeoutSeconds = 600
$script:DiscoveryJob = $null
$script:DiscoveryTimer = $null

function Get-DiscoveryPackagerInput {
    param([object[]]$Rows = @())
    foreach ($row in @($Rows)) {
        if (-not $row) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$row.Vendor) -or [string]::IsNullOrWhiteSpace([string]$row.Application)) { continue }
        [pscustomobject]@{
            Script  = [string]$row.Script
            Vendor  = [string]$row.Vendor
            App     = [string]$row.Application
            CMName  = [string]$row.CMName
            ArpName    = $(if ($row.PSObject.Properties['ArpName']) { [string]$row.ArpName } else { '' })
            ArpVersion = $(if ($row.PSObject.Properties['ArpVersion']) { [string]$row.ArpVersion } else { '' })
            ArpArch    = $(if ($row.PSObject.Properties['ArpArch']) { [string]$row.ArpArch } else { '' })
        }
    }
}

function Invoke-DiscoveryScan {
    # Never throws: a denied login or an unreachable server comes back as
    # Error text so the caller keeps running.
    param(
        [string]$SiteCode,
        [string]$ProviderMachineName,
        [string]$ServerOverride,
        [string]$DatabaseOverride,
        [string]$InstanceOverride = '',
        [int]$PortOverride = 0,
        [object[]]$Packager = @(),
        [string[]]$HiddenScript = @(),
        [string[]]$DismissedScript = @(),
        [hashtable]$OperationState
    )

    # The scan function travels to a runspace as text, so its helper sits inside it. A message that the tool throws
    # itself passes. Any other error is raw text: it goes to the log file and the result gets a fixed sentence.
    function Get-ScanErrorText {
        param($ErrorRecord)
        $ex = $ErrorRecord.Exception
        if ($null -ne $ex -and $null -ne $ex.Data -and $ex.Data.Contains('UserMessage')) { return [string]$ex.Message }
        if ($null -ne $ex -and $ErrorRecord.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::OperationStopped -and $null -eq $ex.InnerException) { return [string]$ex.Message }
        # The runspace loads no SuiteCommon, so Write-Log does not exist here. The shell passes the log path in an environment
        # variable, and the line goes to the file with Add-Content.
        if ($null -ne $ex -and $env:APPPACKAGER_LOG_PATH) { try { Add-Content -LiteralPath $env:APPPACKAGER_LOG_PATH -Value ('[{0}] [ERROR] Discovery scan failed: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $ex.Message) -Encoding UTF8 } catch { $null = $_ } }
        return 'the query failed. See the log file.'
    }
    $result = [pscustomobject]@{ Ok = $false; Error = ''; Server = ''; Database = ''; StatusText = ''; Candidates = @() }
    try {
        if (-not (Get-Command -Name 'Get-CMInventorySnapshot' -ErrorAction SilentlyContinue)) {
            $result.Error = 'The discovery module AppPackagerDiscovery is not loaded. Reinstall AppPackager.'
            return $result
        }

        $targetArgs = @{}
        if (-not [string]::IsNullOrWhiteSpace($SiteCode)) { $targetArgs.SiteCode = $SiteCode.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($ServerOverride))   { $targetArgs.ServerOverride   = $ServerOverride.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($DatabaseOverride)) { $targetArgs.DatabaseOverride = $DatabaseOverride.Trim() }
        $resolver = Get-Command -Name 'Resolve-CMDatabaseTarget' -ErrorAction SilentlyContinue
        if ($resolver -and $resolver.Parameters.ContainsKey('InstanceOverride')) {
            if (-not [string]::IsNullOrWhiteSpace($InstanceOverride)) { $targetArgs.InstanceOverride = $InstanceOverride.Trim() }
            if ($PortOverride -gt 0) { $targetArgs.PortOverride = $PortOverride }
        }
        if ($resolver -and $resolver.Parameters.ContainsKey('ProviderMachineName') -and -not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
            $targetArgs.ProviderMachineName = $ProviderMachineName.Trim()
        }
        $target = Resolve-CMDatabaseTarget @targetArgs
        if ($target -is [string]) { $result.Error = $target; return $result }
        if (-not $target -or [string]::IsNullOrWhiteSpace([string]$target.Server) -or [string]::IsNullOrWhiteSpace([string]$target.Database)) {
            $result.Error = 'The SQL server or database of the ConfigMgr site is not known. Set the SQL server and database in Options, Connections, ConfigMgr.'
            return $result
        }
        $result.Server = [string]$target.Server
        $result.Database = [string]$target.Database

        $snapshotArgs = @{ Server = $result.Server; Database = $result.Database }
        if ($OperationState) { $snapshotArgs.OperationState = $OperationState }
        $snapshot = Get-CMInventorySnapshot @snapshotArgs
        $candidates = @(Find-DiscoveryCandidate -Snapshot $snapshot -Packager @($Packager) -HiddenScript @($HiddenScript) -DismissedScript @($DismissedScript) | Where-Object { $_ })
        $result.Candidates = $candidates
        $result.StatusText = [string](Get-DiscoveryStatusText -Snapshot $snapshot -CandidateCount $candidates.Count)
        $result.Ok = $true
    }
    catch {
        # The scan runs in its own runspace, where Get-UserErrorText does not exist. Get-ScanErrorText is the same rule.
        $result.Error = Get-ScanErrorText -ErrorRecord $_
        $result.Candidates = @()
    }
    return $result
}

function Resolve-DiscoveryProviderName {
    # Runs on the UI thread: the discovery runspace has no CMSite provider,
    # so a drive the console or Check ConfigMgr opened is visible only here.
    param(
        [string]$Preference,
        [string]$SiteCode
    )
    if (-not [string]::IsNullOrWhiteSpace($Preference)) { return $Preference.Trim() }
    if (Get-Command -Name 'Resolve-CMProviderMachineName' -ErrorAction SilentlyContinue) {
        $resolved = [string](Resolve-CMProviderMachineName -ProviderMachineName $Preference)
        if (-not [string]::IsNullOrWhiteSpace($resolved)) { return $resolved.Trim() }
    }
    $drives = @()
    try { $drives = @(Get-PSDrive -PSProvider CMSite -ErrorAction Stop | Where-Object { $_ }) } catch { return '' }
    if ($drives.Count -eq 0) { return '' }
    $drive = $drives | Where-Object { -not [string]::IsNullOrWhiteSpace($SiteCode) -and [string]$_.Name -eq $SiteCode.Trim() } | Select-Object -First 1
    if (-not $drive) { $drive = $drives[0] }
    return ([string]$drive.Root).Trim()
}

function Start-DiscoveryRunspace {
    # Returns the job object that Get-DiscoveryRunspaceState and
    # Stop-DiscoveryRunspace take. Disposes what it created when setup fails.
    param(
        [Parameter(Mandatory)][string]$ModulePath,
        [Parameter(Mandatory)][string]$ScanText,
        [Parameter(Mandatory)][hashtable]$ScanArgs
    )
    $job = [pscustomobject]@{
        State    = [hashtable]::Synchronized(@{ Done = $false; Result = $null; Command = $null; Cancellation = (New-Object System.Threading.CancellationTokenSource) })
        Runspace = $null
        Shell    = $null
        Handle   = $null
        Started  = [DateTime]::UtcNow
    }
    try {
        $ScanArgs.OperationState = $job.State
        $job.Runspace = [runspacefactory]::CreateRunspace()
        $job.Runspace.ApartmentState = 'MTA'
        $job.Runspace.Open()
        $job.Shell = [powershell]::Create()
        $job.Shell.Runspace = $job.Runspace
        [void]$job.Shell.AddScript({
            param($State, $ModulePath, $ScanText, $ScanArgs)
            try {
                if (Test-Path -LiteralPath $ModulePath) { Import-Module -Name $ModulePath -Force -DisableNameChecking -ErrorAction Stop }
                Set-Item -Path 'function:Invoke-DiscoveryScan' -Value ([scriptblock]::Create($ScanText))
                $State.Result = Invoke-DiscoveryScan @ScanArgs
            }
            catch {
                if ($env:APPPACKAGER_LOG_PATH) { try { Add-Content -LiteralPath $env:APPPACKAGER_LOG_PATH -Value ('[{0}] [ERROR] Discovery scan failed: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $_.Exception.Message) -Encoding UTF8 } catch { $null = $_ } }
                $State.Result = [pscustomobject]@{ Ok = $false; Error = 'the query failed. See the log file.'; Server = ''; Database = ''; StatusText = ''; Candidates = @() }
            }
            finally { $State.Done = $true }
        }).AddArgument($job.State).AddArgument($ModulePath).AddArgument($ScanText).AddArgument($ScanArgs)
        $job.Handle = $job.Shell.BeginInvoke()
        return $job
    }
    catch {
        Stop-DiscoveryRunspace -Job $job
        throw
    }
}

function Get-DiscoveryRunspaceState {
    # Running, Done or TimedOut.
    param($Job, [int]$TimeoutSeconds = 600, [datetime]$UtcNow = [DateTime]::UtcNow)
    if (-not $Job) { return 'Done' }
    if ($Job.State.Done) { return 'Done' }
    if (($UtcNow - $Job.Started).TotalSeconds -ge $TimeoutSeconds) { return 'TimedOut' }
    return 'Running'
}

function Stop-DiscoveryRunspace {
    # Never blocks the caller. Dispose on a running PowerShell object waits
    # for the pipeline to stop, and a SQL call can ignore the stop request,
    # so a running pipeline only gets BeginStop and the async runspace close.
    param($Job)
    if (-not $Job) { return }
    if ($Job.State.Cancellation) {
        try { $Job.State.Cancellation.Cancel() } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ }
    }
    $command = $Job.State.Command
    if ($command) { try { $command.Cancel() } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ } }
    if ($Job.Shell) {
        if ($Job.Handle -and -not $Job.State.Done) {
            try { [void]$Job.Shell.BeginStop($null, $null) } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ }
        }
        else {
            if ($Job.Handle) { try { [void]$Job.Shell.EndInvoke($Job.Handle) } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ } }
            try { $Job.Shell.Dispose() } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ }
        }
        $Job.Shell = $null
    }
    if ($Job.Runspace) {
        try { $Job.Runspace.CloseAsync() } catch { try { $Job.Runspace.Dispose() } catch { <# Stopping or disposing a discovery runspace object is best-effort cleanup. #> $null = $_ } }
        $Job.Runspace = $null
    }
}

function Enable-DiscoveredApplication {
    # Same effect as checking the apps in the Product Filter: the scripts
    # leave HiddenApplications. They are also recorded in DiscoveryEnabled
    # so the first Check ConfigMgr auto-hide does not hide them again. The
    # caller saves and refreshes the grid.
    param([Parameter(Mandatory)]$Prefs, [string[]]$Script = @())

    $enable = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Script | Where-Object { $_ }), [System.StringComparer]::OrdinalIgnoreCase)
    $kept = New-Object System.Collections.Generic.List[string]
    $removed = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($Prefs.HiddenApplications)) {
        if ($enable.Contains([string]$name)) { $removed.Add([string]$name) } else { $kept.Add([string]$name) }
    }
    $Prefs.HiddenApplications = $kept.ToArray()
    if ($Prefs.PSObject.Properties['DiscoveryDismissed']) {
        $Prefs.DiscoveryDismissed = [string[]]@($Prefs.DiscoveryDismissed | Where-Object { $_ -and -not $enable.Contains([string]$_) })
    }

    $enabled = New-Object System.Collections.Generic.List[string]
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $existing = if ($Prefs.PSObject.Properties['DiscoveryEnabled']) { @($Prefs.DiscoveryEnabled) } else { @() }
    foreach ($name in @($existing) + @($removed)) { if ($name -and $seen.Add([string]$name)) { $enabled.Add([string]$name) } }
    if ($Prefs.PSObject.Properties['DiscoveryEnabled']) { $Prefs.DiscoveryEnabled = $enabled.ToArray() }
    else { $Prefs | Add-Member -NotePropertyName DiscoveryEnabled -NotePropertyValue $enabled.ToArray() }
    return $removed.Count
}

function Get-AutoHideOffer {
    # Scripts that the first Check ConfigMgr offers to hide; empty means no
    # prompt. Discover can empty HiddenApplications, so the apps it enabled
    # are never offered for hiding again.
    param([Parameter(Mandatory)]$Prefs, [object[]]$Rows = @())

    if (@($Prefs.HiddenApplications | Where-Object { $_ }).Count -gt 0) { return @() }

    $notFound = @(foreach ($item in @($Rows)) {
            if ($item -and [string]$item.Status -eq 'Not found in ConfigMgr') { [string]$item.Script }
        })
    if ($notFound.Count -eq 0 -or $notFound.Count -ge @($Rows).Count) { return @() }

    $enabled = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($Prefs.PSObject.Properties['DiscoveryEnabled']) {
        foreach ($name in @($Prefs.DiscoveryEnabled)) { if ($name) { [void]$enabled.Add([string]$name) } }
    }
    return @($notFound | Where-Object { -not $enabled.Contains($_) })
}

function Add-DiscoveryDismissed {
    param([Parameter(Mandatory)]$Prefs, [string[]]$Script = @())

    $list = New-Object System.Collections.Generic.List[string]
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @($Prefs.DiscoveryDismissed)) { if ($name -and $seen.Add([string]$name)) { $list.Add([string]$name) } }
    $added = 0
    foreach ($name in @($Script)) { if ($name -and $seen.Add([string]$name)) { $list.Add([string]$name); $added++ } }
    $Prefs.DiscoveryDismissed = $list.ToArray()
    if ($Prefs.PSObject.Properties['DiscoveryEnabled']) {
        $Prefs.DiscoveryEnabled = [string[]]@($Prefs.DiscoveryEnabled | Where-Object { $_ -and -not $seen.Contains([string]$_) })
    }
    return $added
}

function Resolve-DiscoveryButton {
    # A non-empty Note means the prompt stays open with no change.
    param(
        [Parameter(Mandatory)][ValidateSet('EnableSelected', 'EnableAll', 'Dismiss')][string]$Button,
        [object[]]$Rows = @()
    )

    $all = @($Rows | Where-Object { $_ } | ForEach-Object { [string]$_.Script })
    if ($Button -eq 'EnableAll') { return [pscustomobject]@{ Action = 'EnableAll'; Scripts = $all; Note = '' } }

    $picked = @($Rows | Where-Object { $_ -and $_.Selected -eq $true } | ForEach-Object { [string]$_.Script })
    if ($picked.Count -eq 0) {
        $note = $(if ($Button -eq 'Dismiss') { 'Check the apps to leave out of later discoveries. Nothing changed.' } else { 'Check at least one app, or click Enable all.' })
        return [pscustomobject]@{ Action = ''; Scripts = @(); Note = $note }
    }
    return [pscustomobject]@{ Action = $Button; Scripts = $picked; Note = '' }
}

function Get-DiscoveryPromptText {
    param([Parameter(Mandatory)][int]$Count)
    return ('Discovered {0} {1}. Enable console view?' -f $Count, $(if ($Count -eq 1) { 'app' } else { 'apps' }))
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
        try { $StartInfo.ArgumentList.Clear() } catch { <# Clearing the ArgumentList is best-effort; the next step sets the arguments. #> $null = $_ }
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
    catch { <# Scanning the packager for optional download folder hints; missing hints keep the default. #> $null = $_ }
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

    if (Test-DeploymentTargetSkipsSite -DeploymentTarget ([string]$PackageArgs.DeploymentTarget)) {
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
            # The WSUS destination does not depend on the ConfigMgr application,
            # so a skipped application still publishes the manifest the probe read.
            if (Test-DeploymentTargetPublishesToWsus -DeploymentTarget ([string]$PackageArgs.DeploymentTarget)) {
                $wsusNote = Invoke-PackagerWsusPostStep -Result $res -PackagerPath ([string]$PackageArgs.PackagerPath) -DownloadRoot ([string]$PackageArgs.DownloadRoot) -Config $PackageArgs.WsusPublishConfig
                $res | Add-Member -NotePropertyName WsusPublish -NotePropertyValue $wsusNote -Force
            }
            return $res
        }
    }
}

function Invoke-PackagerIntuneWinPostStep {
    # Produces <AppFolder>-<Version>.intunewin from the staged content and
    # copies it beside the network content version folder. The artifact
    # lands in the parent of both version folders, never inside them:
    # stage hash verification fails on any file added to verified content.
    # An .intunewin written after the stage manifest is reused, so a package
    # run copies the file its stage built. Failures never fail the run - the
    # ConfigMgr application already exists when this executes - so the
    # returned note carries Ok/Message for the caller to surface.
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [AllowEmptyString()][string]$FileServerPath = '',
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
            $note.Message = 'IntuneWinAppUtil.exe not available; skipped. Configure it in Options, Tools.'
            return $note
        }

        $manifestPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Read stage manifest'
        if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
            $stagedFolder = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Stage complete'
            if (-not [string]::IsNullOrWhiteSpace($stagedFolder)) { $manifestPath = Join-Path $stagedFolder 'stage-manifest.json' }
        }
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

        $existingPath = Join-Path (Split-Path -Path $contentFolder -Parent) $outputName
        $reused = $false
        if ((Test-Path -LiteralPath $existingPath) -and
            (Get-Item -LiteralPath $existingPath).LastWriteTimeUtc -ge (Get-Item -LiteralPath $manifestPath).LastWriteTimeUtc) {
            $existing = Get-Item -LiteralPath $existingPath
            $pkg = [pscustomobject]@{
                IntuneWinPath = $existing.FullName
                SizeBytes     = $existing.Length
                Sha256        = (Get-FileHash -LiteralPath $existing.FullName -Algorithm SHA256).Hash
            }
            $reused = $true
        }
        else {
            $pkg = New-IntuneWinPackage `
                -ToolPath $ToolPath `
                -ContentFolder $contentFolder `
                -SetupFile 'install.bat' `
                -OutputFolder (Split-Path -Path $contentFolder -Parent) `
                -OutputName $outputName
        }
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
        $note.Message = ('{0} {1} ({2:N1} MB, SHA256 {3})' -f $(if ($reused) { 'reused' } else { 'created' }), $outputName, ($pkg.SizeBytes / 1MB), $pkg.Sha256.Substring(0, 12))
        return $note
    }
    catch {
        $note.Message = ('creation failed: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
        return $note
    }
}

function Compare-SemVer {
    # A is the version ConfigMgr or the history holds, B the version the
    # vendor reports. Every part B carries is compared; a part A lacks counts
    # as zero, so "8.9.8.1" after "8.9.8" and "31.0.1" after "31" are
    # updates. Parts A carries beyond B are not compared: an MSI
    # ProductVersion adds a build number ("26.2.2.2") the vendor does not
    # publish ("26.2.2"). -Strict compares all four parts of both sides, for
    # two versions from the same vendor feed.
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B,
        [switch]$Strict
    )
    try {
        # Each side is reduced the way [version] needs: no build suffix
        # ("11.0.30+7"), no channel suffix ("140.17.0esr"), at most four
        # numbers, a single number ("31") padded to "31.0". The parts are
        # compared as numbers without [version], which holds only Int32: a
        # build stamp such as 20250101000000 is above it.
        $reduce = {
            param([string]$v)
            $v = ((($v -replace '[+-].*$', '') -replace '[A-Za-z]+$', '') -replace '^((\d+\.){3}\d+)\..*$', '$1') -replace '^(\d+)$', '$1.0'
            if ($v -notmatch '^\d+(\.\d+){1,3}$') { throw 'not a version' }
            $parts = @($v -split '\.' | ForEach-Object { [decimal]$_ })
            ,$parts
        }
        $pa = & $reduce $A
        $pb = & $reduce $B

        # Parts a version does not carry count as zero.
        $count = 4
        if (-not $Strict) { $count = [Math]::Max(2, $pb.Count) }
        for ($i = 0; $i -lt $count; $i++) {
            $x = $(if ($i -lt $pa.Count) { $pa[$i] } else { [decimal]0 })
            $y = $(if ($i -lt $pb.Count) { $pb[$i] } else { [decimal]0 })
            if ($x -lt $y) { return -1 }
            if ($x -gt $y) { return  1 }
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
            try { $p.Kill() } catch { <# Killing a timed-out process is best-effort; the code throws a timeout error right after. #> $null = $_ }
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

        # A module the packager loads can log to stdout before the version
        # line (the installer analysis writes "Analyzing file: ..."), so the
        # version is the first line that has the shape of one.
        $pattern = '^\d+(\.\d+){0,3}([+-]\d+)?$'
        $version = @($lines | Where-Object { $_ -match $pattern } | Select-Object -First 1)
        if ($version.Count -eq 0) {
            throw ("Unexpected version string: '{0}'" -f ([string]$lines[0]).Trim())
        }
        return ([string]$version[0]).Trim()
    }
    finally {
        if ($p) { try { $p.Dispose() } catch { <# Disposing the process object is cleanup. #> $null = $_ } }
    }
}

function Test-MecmDriveRoot {
    # True when the root of a site drive is the provider that was asked for.
    # A short name and its fully qualified name are the same machine. With no
    # provider asked for, any root passes.
    param([AllowNull()][AllowEmptyString()][string]$Root, [AllowNull()][AllowEmptyString()][string]$Provider)
    if ([string]::IsNullOrWhiteSpace($Provider)) { return $true }
    $a = ([string]$Root).Trim().TrimEnd('\').ToLowerInvariant()
    $b = $Provider.Trim().TrimEnd('\').ToLowerInvariant()
    return ($a -eq $b -or $a.StartsWith($b + '.') -or $b.StartsWith($a + '.'))
}

function Get-OneClickObservedSkipReason {
    # The reason a One Click run skips a destination because the site holds a
    # newer version than the vendor reports, or ''. The publish record says
    # what this tool published; the site says what is there now. A forced run
    # publishes anyway.
    param([string]$Destination, [AllowNull()][AllowEmptyString()][string]$Observed, [AllowNull()][AllowEmptyString()][string]$Latest, [switch]$Force)
    if ($Force -or [string]::IsNullOrWhiteSpace($Observed) -or [string]::IsNullOrWhiteSpace($Latest)) { return '' }
    if ((Compare-SemVer -A $Observed -B $Latest) -gt 0) { return ('{0}: site holds {1}, newer than {2}' -f $Destination, $Observed, $Latest) }
    return ''
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
        } catch { <# Importing the ConfigMgr module is optional; the code throws later if the cmdlets are missing. #> $null = $_ }
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

    # A drive that opens but points at another provider would attribute the
    # answer to a provider that did not give it. It is replaced.
    if ($existingDrive -and -not (Test-MecmDriveRoot -Root ([string]$existingDrive.Root) -Provider $ProviderMachineName)) {
        Remove-PSDrive -Name $SiteCode -Force -ErrorAction SilentlyContinue
        $existingDrive = $null
    }

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
            throw ("Failed to connect to CM site PSDrive '{0}:'. Open the ConfigMgr console once on this machine, or set Site server / SMS Provider in Options, Connections, ConfigMgr (the ProviderMachineName value from the AdminUI connect script)." -f $SiteCode)
        }

        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $providerRoot -ErrorAction Stop | Out-Null
            Set-Location "${SiteCode}:" -ErrorAction Stop
        }
        catch {
            Set-Location $savedLocation -ErrorAction SilentlyContinue
            throw ("Failed to connect to CM site PSDrive '{0}:' via provider '{1}': {2}" -f $SiteCode, $providerRoot, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
        # A wildcard between the words of CMName also returns titles with the
        # version in the middle, and longer product names ("Git" and
        # "GitHub Desktop"); the title filter keeps release details only.
        if (-not $apps -or $apps.Count -eq 0) {
            $apps = @(Get-CMApplication -Name (Get-MecmApplicationNamePattern -CMName $CMName) -ErrorAction SilentlyContinue | Where-Object {
                    $title = if ($_.LocalizedDisplayName) { [string]$_.LocalizedDisplayName } else { [string]$_.Name }
                    Test-MecmApplicationTitle -CMName $CMName -Title $title
                })
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

            # WPF dispatcher pump. Invoke returns a value on every pass; left in
            # the output, it turns the result into an array, and post-step notes
            # (WsusPublish, IntunePublish) cannot attach to it.
            [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )
            Start-Sleep -Milliseconds 50
        }

        if ($timedOut) {
            try { $p.Kill() } catch { <# Killing a timed-out process is best-effort; the code logs the timeout. #> $null = $_ }
            [void]$p.WaitForExit(5000)
            $outLines.Add("[ERROR] Packager idle for $IdleTimeoutSeconds seconds; killed by Invoke-ProcessWithStreaming.")
        }
        elseif (-not $p.WaitForExit(15000)) {
            try { $p.Kill() } catch { <# Killing a timed-out process is best-effort; the code logs the timeout. #> $null = $_ }
            [void]$p.WaitForExit(5000)
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
        if ($p) { try { $p.Dispose() } catch { <# Disposing the process object is cleanup. #> $null = $_ } }
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
        [hashtable]$WsusPublishConfig = $null,
        [ValidateSet('MECM', 'MECMAndIntune', 'IntuneOnly', 'MECMAndWSUS', 'WSUSOnly')][string]$DeploymentTarget = 'MECM',
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
    # Intune-only and WSUS-only runs stop at Stage: no site connection, no
    # share copy, no ConfigMgr application. The .intunewin build, the Graph
    # publish and the WSUS publish below work entirely from the local staged
    # content.
    $skipsSite = Test-DeploymentTargetSkipsSite -DeploymentTarget $DeploymentTarget
    $phaseSwitch = if ($skipsSite) { '-StageOnly' } else { '-PackageOnly' }
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
    if (-not $skipsSite) {
        # Network-copy verification; an Intune-only or WSUS-only run never
        # copies to the share (stage integrity is verified when the manifest
        # is written).
        Assert-PackagerPackageIntegrity -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ContentLayout $ContentLayout
    }

    # Optional post-step: produce a .intunewin beside the network content.
    # Runs only after integrity passes; failures ride on the result for the
    # caller to surface, never thrown - the ConfigMgr application already exists
    # by this point. A WSUS-only run has no network content to place it beside.
    $intuneOnly = ($DeploymentTarget -eq 'IntuneOnly')
    if ((($CreateIntuneWin -and -not $skipsSite) -or $intuneOnly) -and $result.ExitCode -eq 0) {
        $intuneNote = Invoke-PackagerIntuneWinPostStep -Result $result -PackagerPath $PackagerPath -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -ToolPath $IntuneWinToolPath -ContentLayout $ContentLayout -SkipNetworkCopy:$intuneOnly
        if ($intuneOnly -and -not $IntunePublishConfig -and $intuneNote.Ok) {
            $result | Add-Member -NotePropertyName IntunePublish -NotePropertyValue ([pscustomobject]@{ Ok = $false; Message = 'Intune credentials not configured; set Tenant ID, Client ID and Client secret in Options, Connections, Intune.' }) -Force
        }
        if ($IntunePublishConfig -and $intuneNote.Ok) {
            $pubNote = Invoke-PackagerIntunePublish -IntuneNote $intuneNote -Config $IntunePublishConfig -Comment $Comment
            $result | Add-Member -NotePropertyName IntunePublish -NotePropertyValue $pubNote -Force
        }
        if ($intuneNote) {
            $result | Add-Member -NotePropertyName IntuneWin -NotePropertyValue $intuneNote -Force
        }
    }

    # WSUS post-step: publish the staged installer as a locally published
    # update. Same contract as the Intune publish: the note carries the
    # outcome and nothing here throws past the ConfigMgr application.
    if ((Test-DeploymentTargetPublishesToWsus -DeploymentTarget $DeploymentTarget) -and $result.ExitCode -eq 0) {
        $wsusNote = Invoke-PackagerWsusPostStep -Result $result -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot -Config $WsusPublishConfig -SelectedManifestPath $selectedManifestPath -SkipsSite:$skipsSite
        $result | Add-Member -NotePropertyName WsusPublish -NotePropertyValue $wsusNote -Force
    }
    return $result
}

function Invoke-PackagerIntunePublish {
    # Publishes a built .intunewin through Graph. The note carries the
    # outcome and the app id; nothing here throws.
    param(
        [Parameter(Mandatory)]$IntuneNote,
        [Parameter(Mandatory)][hashtable]$Config,
        [AllowEmptyString()][string]$Comment = ''
    )
    $pubNote = [pscustomobject]@{ Ok = $false; Message = ''; AppId = '' }
    try {
        $manifest = Read-StageManifest -Path $IntuneNote.ManifestPath
        if ($manifest.PSObject.Properties['DeploymentTypes'] -and $manifest.DeploymentTypes) {
            $pubNote.Message = 'variant-split app; Intune publish skipped (single-payload apps only for now).'
        }
        else {
            # The staged icon lives in the content version folder next
            # to the manifest, not beside the .intunewin.
            $iconArg = @{}
            if ($manifest.PSObject.Properties['Icon'] -and $manifest.Icon) {
                $iconFile = Join-Path (Split-Path -Parent $IntuneNote.ManifestPath) ([string]$manifest.Icon)
                if (Test-Path -LiteralPath $iconFile) { $iconArg.IconPath = $iconFile }
            }
            $pubId = Publish-IntuneWin32App -TenantId $Config.TenantId -ClientId $Config.ClientId -ClientSecret $Config.ClientSecret -IntuneWinPath $IntuneNote.LocalPath -Manifest $manifest -Description $Comment @iconArg
            $pubNote.Ok = $true
            $pubNote.AppId = [string]$pubId
            $pubNote.Message = ('published (app id {0})' -f $pubId)
        }
    }
    catch {
        $pubNote.Message = ('publish failed: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
    }
    return $pubNote
}

function Get-PackagerConfigMgrApplicationName {
    # The application name the package child logged: on creation, or on a
    # skip of an existing application.
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $m = [regex]::Match($Text, '(?m)^.*Created ConfigMgr application\s*:\s*(.+?)\s*$')
    if (-not $m.Success) { $m = [regex]::Match($Text, '(?m)^.*Application already exists\s*:\s*(.+?)\s+\(v') }
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ''
}

function Invoke-PackagerWsusPostStep {
    # Resolves the manifest this run staged or packaged and publishes its
    # installer to WSUS. A stage-only run names its folder in the "Stage
    # complete" line; a package run names the manifest it read.
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$PackagerPath,
        [string]$DownloadRoot = $null,
        [hashtable]$Config = $null,
        [string]$SelectedManifestPath = '',
        [switch]$SkipsSite
    )

    $note = [pscustomobject]@{
        Ok        = $false
        Refused   = $false
        Message   = ''
        PackageId = ''
        Outcome   = ''
        Warnings  = @()
    }
    try {
        if (-not $Config -or [string]::IsNullOrWhiteSpace([string]$Config.ServerName)) {
            $note.Message = 'WSUS server not configured; set it in Options, Connections, WSUS.'
            return $note
        }
        if (-not (Get-Command -Name Publish-WsusSoftwareUpdate -ErrorAction SilentlyContinue)) {
            $note.Message = 'the WSUS publishing module is not loaded; reinstall AppPackager.'
            return $note
        }

        $manifestPath = ''
        if ($SkipsSite) {
            $stagePath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Stage complete'
            if (-not [string]::IsNullOrWhiteSpace($stagePath)) { $manifestPath = Join-Path $stagePath 'stage-manifest.json' }
        }
        else {
            $manifestPath = Get-PackagerLoggedPath -Text $Result.StdOut -Label 'Read stage manifest'
            if ([string]::IsNullOrWhiteSpace($manifestPath) -and -not [string]::IsNullOrWhiteSpace($SelectedManifestPath)) {
                $manifestPath = $SelectedManifestPath
            }
        }
        if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
            $manifestPath = Find-NewestStageManifestForPackager -PackagerPath $PackagerPath -DownloadRoot $DownloadRoot
        }
        if ([string]::IsNullOrWhiteSpace($manifestPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
            $note.Message = 'stage-manifest.json not found; skipped.'
            return $note
        }

        $manifest = Read-StageManifest -Path $manifestPath
        $publish = Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder (Split-Path -Parent $manifestPath) -Settings $Config
        $note.Ok = $true
        $note.PackageId = [string]$publish.PackageId
        $note.Outcome = [string]$publish.Outcome
        $note.Message = [string]$publish.Message
        if ($publish.PSObject.Properties['Warnings']) { $note.Warnings = @($publish.Warnings) }
        return $note
    }
    catch {
        if ($_.Exception.Data -and $_.Exception.Data.Contains('WsusRefusal')) {
            $note.Refused = $true
            $note.Message = ('not supported: {0}' -f ((Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -replace '^This application cannot be published to WSUS: ', ''))
            return $note
        }
        $note.Message = ('publish failed: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
        return $note
    }
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
    param($Prefs = $script:Prefs)
    if (-not $Prefs) { return '' }
    if ([bool]$Prefs.IncludeVersionInTitle) { return 'IncludeVersion' }
    if ($Prefs.PSObject.Properties['RemoveVersionFromTitle'] -and [bool]$Prefs.RemoveVersionFromTitle) { return 'NoVersion' }
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
    } catch { <# Resolving the tool path from preferences is a safe lookup; an empty result means not detected. #> $null = $_ }
    return ''
}

function Get-IntunePublishConfigForContext {
    # Decrypts the stored client secret (DPAPI, current Windows user)
    # just-in-time. Returns $null when publishing is off or incomplete,
    # so callers can pass the result straight through. Target is the run's
    # destination; empty means the One Click destination.
    param($Prefs = $script:Prefs, [AllowEmptyString()][string]$Target = '')
    try {
        $i = $Prefs.Intune
        if (-not $i) { return $null }
        $target = if ($Target) { $Target } else { [string]$i.DeploymentTarget }
        if (Test-DeploymentTargetPublishesToWsus -DeploymentTarget $target) { return $null }
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

function Get-WsusPublishConfigForContext {
    # Plain values only: the background runspace receives a copy and cannot
    # read the preferences object. Returns $null when the run's destination
    # does not publish to WSUS or no server is set. Target is the run's
    # destination; empty means the One Click destination.
    param($Prefs = $script:Prefs, [AllowEmptyString()][string]$Target = '')
    try {
        $destination = if ($Target) { $Target } else { [string]$Prefs.Intune.DeploymentTarget }
        if (-not (Test-DeploymentTargetPublishesToWsus -DeploymentTarget $destination)) { return $null }
        $w = $Prefs.Wsus
        if (-not $w -or [string]::IsNullOrWhiteSpace([string]$w.ServerName)) { return $null }
        return @{
            ServerName        = [string]$w.ServerName
            PortNumber        = [int]$w.PortNumber
            UseSsl            = [bool]$w.UseSsl
            Classification    = [string]$w.Classification
            ApprovalGroup     = [string]$w.ApprovalGroup
            DeclineSuperseded = [bool]$w.DeclineSuperseded
        }
    } catch { return $null }
}

function Get-IntuneWinToolPathForContext {
    # Same prefs-safe resolution as Get-SevenZipPathForContext, for the
    # Win32 Content Prep Tool.
    try {
        if ($script:Prefs -and $script:Prefs.DetectedTools -and $script:Prefs.DetectedTools.IntuneWinAppUtil -and $script:Prefs.DetectedTools.IntuneWinAppUtil.Found) {
            return [string]$script:Prefs.DetectedTools.IntuneWinAppUtil.ExePath
        }
    } catch { <# Resolving the tool path from preferences is a safe lookup; an empty result means not detected. #> $null = $_ }
    return ''
}

function Select-OnlyUpdateAvailable {
    # Clear every row first, then select within the visible (filtered) set
    # only: a row hidden by the grid filter must never keep or gain
    # Selected=true, or a later Stage/Package would run on rows the user
    # cannot see.
    foreach ($item in $script:PackagerData) { $item.Selected = $false }
    # Check CVEs appends the count: "Update available (3 CVEs, 1 in KEV)".
    foreach ($item in @($dataGrid.ItemsSource)) {
        $item.Selected = ([string]$item.Status -like "Update available*")
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
        $null = $_
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
    # -Suite: the update is a suite release (this copy runs inside a suite
    # install); Update now runs the suite installer instead of install.ps1.
    param(
        [AllowNull()][string]$LatestVersion,
        [AllowNull()][string]$ReleaseUrl,
        [switch]$Suite
    )

    if (-not $pnlUpdate) { return }

    if ([string]::IsNullOrWhiteSpace($LatestVersion)) {
        $pnlUpdate.Visibility = 'Collapsed'
        return
    }

    $script:UpdateLatestVersion = $LatestVersion
    $script:UpdateIsSuite       = [bool]$Suite
    $script:UpdateReleaseUrl    = if ([string]::IsNullOrWhiteSpace($ReleaseUrl)) { Get-UpdateReleaseUrl -Version $LatestVersion } else { $ReleaseUrl }
    $runUpdateText.Text         = $(if ($Suite) { "Suite update available: v{0}" -f $LatestVersion } else { "Update available: v{0}" -f $LatestVersion })
    $lnkUpdateAvailable.ToolTip = ("Open the v{0} release page on GitHub" -f $LatestVersion)
    $pnlUpdate.Visibility       = 'Visible'
}

function Start-UpdateCheck {
    # Never let a failed or slow check touch launch: everything below is
    # best-effort and the caller keeps going regardless.
    try {
        # Inside a suite install the suite release is the update: its
        # installer replaces every tool at once. The lookup and its daily
        # cache are shared with the launcher and the other tools.
        $suite = $null
        if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            try { $suite = Get-SuiteInstallInfo -Path $PSScriptRoot } catch { $suite = $null }
        }
        if ($suite -and $suite.Installed) {
            $null = Start-SuiteUpdateNotice -Path $PSScriptRoot -OnResult {
                param($u)
                if ($u.Available) {
                    Add-LogLine -Message ("Suite update available: v{0}." -f $u.Latest)
                    Show-UpdateIndicator -LatestVersion $u.Latest -ReleaseUrl $u.ReleaseUrl -Suite
                }
            }
            return
        }

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
            } catch { <# Update check is optional and its cleanup calls are best-effort. #> $null = $_ } finally {
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
                Add-LogLine -Message ("Update check failed: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            } finally {
                try { $script:UpdateCheckPS.EndInvoke($script:UpdateCheckHandle) } catch { <# Update check is optional and its cleanup calls are best-effort. #> $null = $_ }
                try { $script:UpdateCheckPS.Dispose() } catch { <# Update check is optional and its cleanup calls are best-effort. #> $null = $_ }
                try { $script:UpdateCheckRunspace.Close() } catch { <# Update check is optional and its cleanup calls are best-effort. #> $null = $_ }
                $script:UpdateCheckPS = $null
                $script:UpdateCheckRunspace = $null
            }
        })
        $script:UpdateCheckTimer.Start()
    } catch {
        try { Add-LogLine -Message ("Update check skipped: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) } catch { <# Update check is optional and its cleanup calls are best-effort. #> $null = $_ }
    }
}

function Invoke-SelfUpdate {
    param([Parameter(Mandatory)]$Owner)

    $latest = $script:UpdateLatestVersion
    if ([string]::IsNullOrWhiteSpace($latest)) { return }

    # Inside a suite install the launcher is the one place that updates;
    # install.ps1 over this folder would break the suite's file list.
    if ($script:UpdateIsSuite) {
        $launcher = Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'app-packager-suite') 'start-suite.ps1'
        if (Test-Path -LiteralPath $launcher) {
            Add-LogLine -Message ("Suite update v{0}: opening the AppPackager Suite Launcher. Close this and every other suite tool, then click Update there." -f $latest)
            try { Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $launcher) -WorkingDirectory (Split-Path -Parent $launcher) } catch { Add-LogLine -Message ('Could not start the launcher: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
        }
        else {
            [void](Show-ThemedMessage -Owner $Owner -Title 'Suite Update' -Message ("Suite v{0} is available. Update from the AppPackager Suite Launcher in the Start Menu." -f $latest) -Buttons OK -Icon Information)
        }
        return
    }

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
Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','"$($relaunch.Replace("'","''"))"'
"@
        $bootstrap = Join-Path ([IO.Path]::GetTempPath()) ('apupdate-' + [Guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $bootstrap -Value $script -Encoding UTF8

        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $bootstrap)

        Add-LogLine -Message ("Updating to v{0}; closing." -f $latest)
        $Owner.Close()
    } catch {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Update Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
    }
}

function Open-UpdateReleasePage {
    try {
        $url = $script:UpdateReleaseUrl
        if ([string]::IsNullOrWhiteSpace($url)) { $url = Get-UpdateReleaseUrl -Version $script:UpdateLatestVersion }
        Start-Process $url
    } catch {
        Add-LogLine -Message ("Could not open the release page: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -eq 403 -or $status -eq 429) {
                $result.Message = 'GitHub rate limit reached; try again later.'
            } elseif ($status -eq 404) {
                $result.Message = 'No icon pack release is published yet.'
            } else {
                $result.Message = ("Icon pack lookup failed: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
        $result.Message = ("Icon pack install failed: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
        $result.Message = ("Icon pack install failed: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
            ForceInstall            = $false
            BidirContentRedirection = $false
            ClientHostedAppsUrl     = $false
            ForceLaa                = $false
            LegacyFtaIcons          = $false
            DisableSetting          = ''
            ClientName              = ''
        }
        Plugins = [pscustomobject]@{
            MSTeamsPlugin        = $true
            ZoomPlugin           = $true
            WebExPlugin          = $false
            UberAgent            = $false
            UberAgentSkipUpgrade = $false
            EPAClient            = $true
            SessionRecording     = $false
            SkipSRAgentUpgrade   = $false
        }
        UpdateAndTelemetry = [pscustomobject]@{
            AutoUpdateCheck  = "disabled"
            AutoUpdateStream = ""
            AURolloutPriority = ""
            EnableCEIP       = $false
            EnableTracing    = $false
        }
        StorePolicy = [pscustomobject]@{
            AllowAddStore = ""
            AllowSavePwd  = ""
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
            CitrixEnterpriseBrowser = $false
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
        foreach ($prop in @('CleanInstall','IncludeSSON','EnableSSON','AppProtection','SessionPreLaunch','SelfServiceMode','ForceInstall','BidirContentRedirection','ClientHostedAppsUrl','ForceLaa','LegacyFtaIcons')) {
            if ($null -ne $data.Installation.$prop) { $defaults.Installation.$prop = [bool]$data.Installation.$prop }
        }
        foreach ($prop in @('DisableSetting','ClientName')) {
            if ($null -ne $data.Installation.$prop) { $defaults.Installation.$prop = [string]$data.Installation.$prop }
        }
        foreach ($prop in @('MSTeamsPlugin','ZoomPlugin','WebExPlugin','UberAgent','UberAgentSkipUpgrade','EPAClient','SessionRecording','SkipSRAgentUpgrade')) {
            if ($null -ne $data.Plugins.$prop) { $defaults.Plugins.$prop = [bool]$data.Plugins.$prop }
        }
        if ($null -ne $data.UpdateAndTelemetry) {
            if ($null -ne $data.UpdateAndTelemetry.AutoUpdateCheck) { $defaults.UpdateAndTelemetry.AutoUpdateCheck = [string]$data.UpdateAndTelemetry.AutoUpdateCheck }
            if ($null -ne $data.UpdateAndTelemetry.AutoUpdateStream) { $defaults.UpdateAndTelemetry.AutoUpdateStream = [string]$data.UpdateAndTelemetry.AutoUpdateStream }
            if ($null -ne $data.UpdateAndTelemetry.AURolloutPriority) { $defaults.UpdateAndTelemetry.AURolloutPriority = [string]$data.UpdateAndTelemetry.AURolloutPriority }
            if ($null -ne $data.UpdateAndTelemetry.EnableCEIP)      { $defaults.UpdateAndTelemetry.EnableCEIP      = [bool]$data.UpdateAndTelemetry.EnableCEIP }
            if ($null -ne $data.UpdateAndTelemetry.EnableTracing)   { $defaults.UpdateAndTelemetry.EnableTracing   = [bool]$data.UpdateAndTelemetry.EnableTracing }
        }
        if ($null -ne $data.StorePolicy) {
            if ($null -ne $data.StorePolicy.AllowAddStore) { $defaults.StorePolicy.AllowAddStore = [string]$data.StorePolicy.AllowAddStore }
            if ($null -ne $data.StorePolicy.AllowSavePwd)  { $defaults.StorePolicy.AllowSavePwd  = [string]$data.StorePolicy.AllowSavePwd }
        }
        foreach ($prop in @('Customize','ReceiverInside','ICA_Client','AM','SelfService','DesktopViewer','WebHelper','BCR_Client','USB','SSON','CitrixEnterpriseBrowser')) {
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
    # Build records seal the Get-SigningPolicy digest and the freshness check
    # compares against it; any other digest marks every build stale.
    param($Prefs = $script:Prefs)
    $json = Get-WorkbenchSigningPolicyJson -Prefs $Prefs
    if (-not $json) { return '' }
    if (Get-Command -Name 'Get-SigningPolicy' -ErrorAction SilentlyContinue) {
        return [string](Get-SigningPolicy -Json $json).PolicyDigest
    }
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

        # The profile decides the staging root, so it is resolved before
        # anything compares or launches against that root.
        $batchAppId = ''
        $batchProfileId = 'default'
        $batchDownloadRoot = $DownloadRoot
        if ($batchWorkbenchRoot -and (Get-Command -Name 'Get-WorkbenchProfileDownloadRoot' -ErrorAction SilentlyContinue)) {
            try {
                $batchAppId = [string](New-ApplicationId -Kind Catalog -ScriptPath $scriptPath)
                $batchDefinition = Get-ApplicationDefinition -ApplicationId $batchAppId
                if ([string]$batchDefinition.ActiveProfileId) { $batchProfileId = [string]$batchDefinition.ActiveProfileId }
                $batchDownloadRoot = Get-WorkbenchProfileDownloadRoot -DownloadRoot $DownloadRoot -ProfileId $batchProfileId
            }
            catch {
                Write-Log ("[batch] {0}: profile not resolved, using the base staging root: {1}" -f $baseName, $_.Exception.Message) -Level WARN
                $batchAppId = ''
                $batchProfileId = 'default'
                $batchDownloadRoot = $DownloadRoot
            }
        }
        $stageScope   = [string]$batchDownloadRoot
        $packageScope = ('{0}|{1}|{2}' -f $SiteCode, $FileServerPath, $batchProfileId)

        $history = Read-PackagerHistory
        $lastKnown    = $null
        $lastChecked  = $null
        $lastStaged   = $null
        $lastPackaged = $null
        $entry = $null
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
        # The version this kind of run has already completed for this
        # destination. A check records the newest version seen, which says
        # nothing about what was built.
        $actedVersion = switch ($OnUpdateFound) {
            'Stage'           { Get-PackagerHistoryActedVersion -Entry $entry -Event Staged -Scope $stageScope }
            'StageAndPackage' { Get-PackagerHistoryActedVersion -Entry $entry -Event Packaged -Scope $packageScope }
            default           { Get-PackagerHistoryActedVersion -Entry $entry -Event Checked }
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
                } catch { <# Reading the cadence from packager metadata is optional; the default cadence applies. #> $null = $_ }
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
            } catch { <# A failed cadence date check fails open, so the app is checked instead of skipped. #> $null = $_ }
        }

        # 1. Discover vendor's current version. On failure, don't update
        # LastChecked - next run should retry, not wait for cadence.
        $latest = $null
        try {
            $latest = Invoke-PackagerGetLatestVersion -PackagerPath $scriptPath -SiteCode $SiteCode -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot
        } catch {
            Write-Log ("[batch] {0}: check failed: {1}" -f $baseName, $_.Exception.Message) -Level WARN
            $results += [pscustomobject]@{ Name = $baseName; Action = 'CheckFailed'; OldVersion = $lastKnown; NewVersion = $null; Reason = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
            continue
        }

        # A vendor feed can step back to an older release (Dell's catalog
        # listed 5.7.0 after 5.7.1). Acting would publish the older build, so
        # the run skips it and keeps the newer version on record.
        if (-not $Force -and $lastKnown -and (Compare-SemVer -A ([string]$lastKnown) -B ([string]$latest) -Strict) -gt 0) {
            Write-Log ("[batch] [VendorOlder] {0}: vendor reports {1}, older than {2} seen before; skipping" -f $baseName, $latest, $lastKnown) -Level WARN
            $results += [pscustomobject]@{ Name = $baseName; Action = 'Skipped'; OldVersion = $lastKnown; NewVersion = $latest; Reason = ("vendor reports {0}, older than {1}" -f $latest, $lastKnown) }
            continue
        }

        # 2. Decide whether to act. NoChange short-circuit applies only
        # when there's nothing new to do: same version AND (for Stage /
        # StageAndPackage) we've already staged/packaged that version
        # at least once. First-time stages and Force=on always run.
        $versionChanged = (-not $actedVersion) -or ($actedVersion -ne $latest)

        # A matching version is not freshness: a changed profile revision or
        # signing policy makes the last build stale, as in a Full Run.
        $buildStale = $false
        if ($batchAppId -and $OnUpdateFound -ne 'Report' -and (Get-Command -Name 'Get-LatestBuildRecord' -ErrorAction SilentlyContinue) -and
            (Get-LatestBuildRecord -ApplicationId $batchAppId -ProfileId $batchProfileId -ErrorAction SilentlyContinue)) {
            # Only an application with a build record can be compared; one without
            # has nothing to be stale against.
            try {
                $batchRevision = [int](Get-Profile -ApplicationId $batchAppId -ProfileId $batchProfileId).Revision
                $batchDigest = ''
                if (-not [string]::IsNullOrWhiteSpace($SigningJson) -and (Get-Command -Name 'Get-SigningPolicy' -ErrorAction SilentlyContinue)) {
                    $batchDigest = [string](Get-SigningPolicy -Json $SigningJson).PolicyDigest
                }
                $current = Test-WorkbenchBuildIsCurrent -ApplicationId $batchAppId -ProfileId $batchProfileId `
                    -Version ([string]$latest) -ProfileRevision $batchRevision -PolicyDigest $batchDigest
                if ($null -ne $current -and -not $current) { $buildStale = $true }
            }
            catch { $buildStale = $false }
        }
        $shouldAct = $versionChanged -or $Force -or $buildStale

        if (-not $shouldAct) {
            Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result NoChange
            Write-Log ("[batch] [NoChange] {0}: {1}" -f $baseName, $latest) -Level INFO
            $results += [pscustomobject]@{ Name = $baseName; Action = 'NoChange'; OldVersion = $lastKnown; NewVersion = $latest }
            continue
        }

        # 3. Taking action
        $oldDisplay = if ($actedVersion) { $actedVersion } else { '(none)' }
        $label = if ($versionChanged) { 'Updated' } elseif ($buildStale) { 'Rebuilt' } else { 'Forced' }
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
        if ($batchDownloadRoot)    { $pkgArgs += @('-DownloadRoot',         $batchDownloadRoot)    }
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
            if ($batchWorkbenchRoot -and (Get-Command -Name 'New-RunSnapshot' -ErrorAction SilentlyContinue)) {
                # The snapshot carries the signing policy so the certificate
                # gate runs here, before the child writes any content.
                $batchSigningPolicy = $null
                if (-not [string]::IsNullOrWhiteSpace($SigningJson)) {
                    try { $batchSigningPolicy = $SigningJson | ConvertFrom-Json } catch { $batchSigningPolicy = $null }
                }
                try {
                    if (-not $batchAppId) { $batchAppId = [string](New-ApplicationId -Kind Catalog -ScriptPath $scriptPath) }
                    $batchSnapshot = New-RunSnapshot -ApplicationId $batchAppId -ProfileId $batchProfileId `
                        -Target 'MECM' -SigningPolicy $batchSigningPolicy -PackagerScriptPath $scriptPath -DownloadRoot $batchDownloadRoot
                    $batchRunSnapshot = [string]$batchSnapshot.Path
                }
                catch {
                    if ($_.Exception.Message -match 'SigningCertificateUnavailable') {
                        Write-Log ("[batch] {0}: skipped, the signing policy cannot be met: {1}" -f $baseName, $_.Exception.Message) -Level ERROR
                        $results += [pscustomobject]@{ Name = $baseName; Action = 'Failed'; OldVersion = $lastKnown; NewVersion = $null; Reason = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
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
                    try { Pop-Location } catch { <# Restoring the previous location with Pop-Location is cleanup. #> $null = $_ }
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

            Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $latest -Result Updated -Scope $stageScope
            if ($OnUpdateFound -eq 'StageAndPackage') {
                Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $latest -Result Updated -Scope $packageScope
            }
            $results += [pscustomobject]@{ Name = $baseName; Action = $OnUpdateFound; OldVersion = $lastKnown; NewVersion = $latest }
        } catch {
            # Don't update LastChecked on failure - next run should retry
            # immediately, not wait for cadence to expire.
            Write-Log ("[batch] {0}: action failed: {1}" -f $baseName, $_.Exception.Message) -Level ERROR
            $results += [pscustomobject]@{ Name = $baseName; Action = 'Failed'; OldVersion = $lastKnown; NewVersion = $latest; Reason = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
        }
    }

    return ,$results
}

# =============================================================================
# Batch-mode entry point (exits before WPF)
# =============================================================================
# =============================================================================
# CVE helpers shared by the grid, the CVE window, the report button and the
# headless monitor. Defined before the batch dispatch so -CveReport can use
# them without the WPF shell.
# =============================================================================

function Get-CveNvdApiKey {
    # The protected key from the preferences, else NVD_API_KEY, else ''.
    $key = ''
    try {
        $protected = [string]$script:Prefs.Vulnerability.NvdApiKeyProtected
        if (-not [string]::IsNullOrWhiteSpace($protected)) {
            $sec = ConvertTo-SecureString -String $protected -ErrorAction Stop
            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($sec)
            try { $key = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
        }
    } catch { $key = '' }
    if ([string]::IsNullOrWhiteSpace($key) -and $env:NVD_API_KEY) { $key = [string]$env:NVD_API_KEY }
    return ([string]$key).Trim()
}

function Get-CveCacheTtlMinutes {
    $ttl = 360
    try { if ($null -ne $script:Prefs.Vulnerability.CacheTtlMinutes) { $ttl = [int]$script:Prefs.Vulnerability.CacheTtlMinutes } } catch { <# Reading the CVE cache lifetime is optional; the default value stays. #> $null = $_ }
    return $ttl
}

function Get-CveRunContext {
    # The background context of one Check CVEs run. KevCatalog is filled by
    # the first row that needs it and shared by the rest of the run.
    # CveSession keeps the results of every Check CVEs run of this session, so a
    # report uses them when a cache write failed.
    if ($null -eq $script:CveSessionResults) { $script:CveSessionResults = [hashtable]::Synchronized(@{}) }
    @{
        CveSession      = $script:CveSessionResults
        NvdApiKey       = Get-CveNvdApiKey
        CacheTtlMinutes = Get-CveCacheTtlMinutes
        KevCatalog      = $null
        ForceRefresh    = $false
    }
}

function ConvertTo-CveVersionKey {
    # The same reduction the vulnerability module applies to a version before
    # it becomes a cache key (ConvertTo-VulnerabilityVersion), inline so the
    # grid needs no module.
    param([AllowNull()][AllowEmptyString()][string]$Version)
    if ([string]::IsNullOrWhiteSpace($Version)) { return '' }
    $v = ($Version.Trim() -replace '\s.*$', '') -replace '\+.*$', ''
    if ($v -notmatch '\d') { return '' }
    return $v
}

function ConvertTo-CvePairKey {
    # "from|to" for the versions of a row. A lookup needs one of the two,
    # so either side may be empty; '' when neither has a number.
    param([AllowNull()][AllowEmptyString()][string]$Current, [AllowNull()][AllowEmptyString()][string]$Latest)
    $from = ConvertTo-CveVersionKey -Version $Current
    $to = ConvertTo-CveVersionKey -Version $Latest
    if (-not $from -and -not $to) { return '' }
    return ('{0}|{1}' -f $from, $to)
}

function Set-CveCellsPair {
    # A row remembers the version pair its CVE cells describe. Without -Pair
    # the pair is the one the row holds now.
    param([Parameter(Mandatory)]$Row, [AllowEmptyString()][string]$Pair)
    if (-not $PSBoundParameters.ContainsKey('Pair')) { $Pair = ConvertTo-CvePairKey -Current ([string]$Row.CurrentVersion) -Latest ([string]$Row.LatestVersion) }
    if ($Row.PSObject.Properties['CvePair']) { $Row.CvePair = $Pair }
    else { Add-Member -InputObject $Row -NotePropertyName CvePair -NotePropertyValue $Pair }
}

function Clear-StaleCveCells {
    # Check Latest, Check ConfigMgr and One Click change the versions of a
    # row. The CVE cells and the CVE count in the status describe the pair
    # the row held when they were filled; for any other pair they are wrong,
    # so they are cleared.
    param([Parameter(Mandatory)]$Row)
    $cellNames = @('CveCount', 'CveCvss', 'CveKev', 'CveEpss', 'CveChecked', 'CveFixed', 'CveNewest', 'CveUpdated')
    $filled = $false
    foreach ($n in $cellNames) { if ($Row.PSObject.Properties[$n] -and [string]$Row.$n) { $filled = $true } }
    if (-not $filled) { return }
    $stored = $(if ($Row.PSObject.Properties['CvePair']) { [string]$Row.CvePair } else { '' })
    if ($stored -eq (ConvertTo-CvePairKey -Current ([string]$Row.CurrentVersion) -Latest ([string]$Row.LatestVersion))) { return }
    foreach ($n in $cellNames) { if ($Row.PSObject.Properties[$n]) { $Row.$n = '' } }
    if ($Row.PSObject.Properties['CveCount'] -and $Row.PSObject.Properties['Cpe'] -and [string]::IsNullOrWhiteSpace([string]$Row.Cpe)) { $Row.CveCount = 'CPE missing' }
    if ($Row.PSObject.Properties['Status'] -and [string]$Row.Status) { $Row.Status = ([string]$Row.Status -replace '^(Update available) \(.*\)$', '$1') }
    Set-CveCellsPair -Row $Row -Pair ''
}

function Get-CveSkipReason {
    # Why Check CVEs skips a row, or '' when the row can be looked up. A
    # lookup needs a CPE header and one version: the ConfigMgr version, or
    # the latest version when ConfigMgr holds none.
    param([AllowNull()][AllowEmptyString()][string]$Cpe, [AllowNull()][AllowEmptyString()][string]$Current, [AllowNull()][AllowEmptyString()][string]$Latest)
    if ([string]::IsNullOrWhiteSpace($Cpe)) { return 'CPE missing' }
    if (-not (ConvertTo-CvePairKey -Current $Current -Latest $Latest)) { return 'No version' }
    return ''
}

function Remove-ExpiredCveReports {
    # Deletes the report files older than KeepDays. Only the names the report
    # writer produces (Export-VulnerabilityReport with its default name) are
    # deleted; any other file in the folder stays.
    param([Parameter(Mandatory)][string]$Folder, [int]$KeepDays)
    if ($KeepDays -le 0 -or -not (Test-Path -LiteralPath $Folder)) { return }
    $cutoff = (Get-Date).AddDays(-$KeepDays)
    Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^cve-report-\d{8}-\d{6}(-applications\.csv|-cves\.csv|\.json|\.html|\.xlsx|\.md)$' -and $_.LastWriteTime -lt $cutoff } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

function Get-CveGridCells {
    # The CVE cells of a grid row from a history summary (hashtable or the
    # PSCustomObject ConvertFrom-Json returns), or the placeholder for a
    # row that has no CPE header. No module call: the grid renders before
    # and without the vulnerability module.
    # Count is the number of CVEs that affect the examined version: the
    # ConfigMgr version, or the latest version when ConfigMgr holds none.
    # Fixed is "5 of 15" when the latest version is newer: the update
    # removes 5 of the 15. Text is the same in words, for the status.
    # With -CurrentVersion and -LatestVersion the summary must be for that
    # version pair; a count left from older versions is not shown.
    param($Summary, [AllowNull()][AllowEmptyString()][string]$Cpe, [AllowNull()][AllowEmptyString()][string]$CurrentVersion, [AllowNull()][AllowEmptyString()][string]$LatestVersion)
    $get = { param($name) if ($null -eq $Summary) { $null } elseif ($Summary -is [System.Collections.IDictionary]) { $Summary[$name] } elseif ($Summary.PSObject.Properties[$name]) { $Summary.PSObject.Properties[$name].Value } else { $null } }
    $cells = [pscustomobject]@{ Cpe = [string]$Cpe; Count = ''; Cvss = ''; Kev = ''; Epss = ''; Checked = ''; Text = ''; Fixed = ''; Newest = ''; Updated = '' }
    if ([string]::IsNullOrWhiteSpace($Cpe)) { $cells.Count = 'CPE missing'; return $cells }
    if ($null -eq $Summary) { return $cells }
    # A summary of another layout counts other things under the same names
    # (the layout number is the SchemaVersion of the vulnerability module).
    if ("$(& $get 'Schema')" -ne '2') { return $cells }
    # A summary records the CPE it was looked up for. After the CPE of the
    # packager changed, the count of the old product is not shown under the
    # new one.
    if ([string](& $get 'Cpe') -ne [string]$Cpe) { return $cells }
    if ($PSBoundParameters.ContainsKey('CurrentVersion') -or $PSBoundParameters.ContainsKey('LatestVersion')) {
        $from = ConvertTo-CveVersionKey -Version $CurrentVersion
        $to = ConvertTo-CveVersionKey -Version $LatestVersion
        if ((-not $from -and -not $to) -or [string](& $get 'From') -ne $from -or [string](& $get 'To') -ne $to) { return $cells }
    }
    # PowerShell 7 reads the time text of the history file back as a
    # DateTime; the cell keeps the text the lookup writes, which the report
    # compares with the time of the cached result.
    $checked = & $get 'QueriedAt'
    if ($checked -is [datetime]) {
        $stamp = [datetime]$checked
        if ($stamp.Kind -eq [System.DateTimeKind]::Unspecified) { $stamp = [datetime]::SpecifyKind($stamp, [System.DateTimeKind]::Utc) }
        $checked = $stamp.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [cultureinfo]::InvariantCulture)
    }
    $cells.Checked = [string]$checked
    if (& $get 'Error') { $cells.Count = 'Error'; return $cells }
    $count = [int](& $get 'Count')
    $kevCount = [int](& $get 'KevCount')
    $update = ([string](& $get 'Mode') -eq 'Update')
    # NVD answered and holds no record under the product name of the CPE
    # header: the product has no known CVE, or NVD files it under another
    # name. The 0 is not a finding about the version, so the cell says so
    # (Test-VulnerabilityProductUnknown of the vulnerability module).
    $records = & $get 'ProductRecords'
    $unknown = ($null -ne $records -and "$records" -ne '' -and [int]$records -eq 0)
    $cells.Count = $(if ($unknown) { 'Not in NVD' } else { [string]$count })
    if ($unknown) { $cells.Text = 'product not in NVD' }
    elseif ($count -eq 0) { $cells.Text = 'no CVE' }
    else {
        $cells.Text = '{0} CVE{1}' -f $count, $(if ($count -eq 1) { '' } else { 's' })
        if ($update) { $cells.Text = 'fixes {0} of {1}' -f [int](& $get 'FixedCount'), $cells.Text }
        if ($kevCount -gt 0) { $cells.Text += ', {0} in KEV' -f $kevCount }
    }
    $score = & $get 'MaxCvss'
    if ($null -ne $score -and "$score" -ne '') {
        $sev = [string](& $get 'MaxSeverity')
        $cells.Cvss = ('{0:N1} {1}' -f [double]$score, $sev).Trim()
    }
    if ($kevCount -gt 0) { $cells.Kev = 'Yes' }
    $epss = & $get 'MaxEpss'
    if ($null -ne $epss -and "$epss" -ne '') { $cells.Epss = ('{0:P1}' -f [double]$epss) }
    if ($update -and $count -gt 0) { $cells.Fixed = '{0} of {1}' -f [int](& $get 'FixedCount'), $count }
    $cells.Newest = [string](& $get 'NewestPublished')
    $cells.Updated = [string](& $get 'LastNvdUpdate')
    return $cells
}

function Find-CveSessionResult {
    # The result of a lookup this session made for the row: same script, same
    # CPE, same version pair, and no error. $null when there is none.
    param([Parameter(Mandatory)]$Row, $Results)
    if (-not $Results -or -not $Results.ContainsKey([string]$Row.Script)) { return $null }
    $candidate = $Results[[string]$Row.Script]
    if ($candidate -and -not $candidate.Error -and [string]$candidate.Cpe -eq [string]$Row.Cpe -and
        [string]$candidate.From -eq (ConvertTo-VulnerabilityVersion -Version ([string]$Row.CurrentVersion)) -and
        [string]$candidate.To -eq (ConvertTo-VulnerabilityVersion -Version ([string]$Row.LatestVersion))) { return $candidate }
    return $null
}

function Get-CveCachedResult {
    # The saved lookup for a grid row, or $null when none was run for the
    # row's CPE and version pair.
    param([Parameter(Mandatory)]$Row)
    $cpe = [string]$Row.Cpe
    if ([string]::IsNullOrWhiteSpace($cpe)) { return $null }
    # A lookup of this session comes first: its cache write can have failed.
    $session = Find-CveSessionResult -Row $Row -Results $script:CveSessionResults
    if ($session) { return (ConvertTo-VulnerabilityResultObject -Result $session) }
    return (Read-VulnerabilityCachedResult -Cpe $cpe -CurrentVersion ([string]$Row.CurrentVersion) -LatestVersion ([string]$Row.LatestVersion))
}

function Get-CveReportRows {
    # Always an array, also for zero rows (the comma keeps an empty array
    # from unrolling into $null).
    # -Results holds the lookups of this run by script name. A result of the
    # run is exported as it is, also when the cache write failed; the cache
    # serves only a row the run did not look up. -MarkCached says so in the
    # note of such a row (the monitor; the grid keeps the history of earlier
    # sessions on purpose).
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Rows, [hashtable]$Results = $null, [switch]$MarkCached)
    $out = foreach ($row in $Rows) {
        $res = Find-CveSessionResult -Row $row -Results $Results
        $own = [bool]$res
        if (-not $res) { $res = Get-CveCachedResult -Row $row }
        $note = ''
        $count = [string]$row.CveCount
        # A row whose latest lookup failed keeps the failure note, also when
        # an earlier lookup of this session is the result shown.
        if ($own -and $count -ne 'Error') { $note = '' }
        elseif (-not $res -and $count -and $count -notmatch '^\d+$' -and $count -ne 'Not in NVD') { $note = $count }
        elseif (-not $res) { $note = 'Not checked' }
        # The cache holds the last good lookup. A row whose latest lookup
        # failed, or whose lookup is newer than the cache (its cache write
        # failed), says so instead of passing the old result off as current.
        elseif ($count -eq 'Error') { $note = ('Last lookup failed; this is the result from {0}' -f $res.QueriedAt) }
        elseif ([string]$row.CveChecked -and [string]$row.CveChecked -gt [string]$res.QueriedAt) { $note = ('Result from {0}; the lookup of {1} was not cached' -f $res.QueriedAt, [string]$row.CveChecked) }
        elseif ($MarkCached) { $note = 'From the cache; not looked up in this run' }
        [pscustomobject]@{
            Application    = [string]$row.Application
            Vendor         = [string]$row.Vendor
            Script         = [string]$row.Script
            CurrentVersion = [string]$row.CurrentVersion
            LatestVersion  = [string]$row.LatestVersion
            Status         = [string]$row.Status
            Cpe            = [string]$row.Cpe
            Note           = $note
            Result         = $res
            # True when the result is a lookup of this run, whatever the note says.
            FromRun        = $own
        }
    }
    return ,@($out)
}

function Get-CveReportFolder {
    Join-Path (Join-Path $PSScriptRoot 'Logs') 'Reports'
}

function Invoke-CveMonitor {
    <#
    .SYNOPSIS
        Headless CVE monitor: the Check Latest, Check ConfigMgr and Check
        CVEs flow for every packager (or -Apps), then the report and the
        notification. Shares app-history.json and the CVE cache with the
        GUI, so a scheduled run fills the grid for the next session.
    .OUTPUTS
        [pscustomobject] Total, Updates, LookedUp, Cached, Cves, Fixed, Open,
        Kev, Failed, Paths, Started, Ended, Computer. Cves counts the CVEs
        that affect the ConfigMgr versions; Fixed of them are absent from
        the latest versions and Open are not.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagersRoot,
        [AllowEmptyCollection()][string[]]$Apps = @(),
        [Parameter(Mandatory)][string]$SiteCode,
        [AllowEmptyString()][string]$ProviderMachineName = '',
        [AllowEmptyString()][string]$DownloadRoot = '',
        [AllowEmptyString()][string]$FileServerPath = '',
        [AllowEmptyString()][string]$M365Channel = '',
        [AllowEmptyString()][string]$M365DeployMode = '',
        [switch]$SkipLatest,
        [switch]$SkipMecm,
        [switch]$SkipCve,
        [AllowEmptyString()][string]$NvdApiKey = '',
        [int]$CacheTtlMinutes = 360,
        [Parameter(Mandatory)][string]$ReportFolder,
        [string[]]$Formats = @('Html', 'Csv', 'Json', 'Xlsx', 'Markdown'),
        [string]$ReportName = '',
        [AllowEmptyString()][string]$WebhookUrl = '',
        [AllowEmptyString()][string]$DropFolder = '',
        [switch]$NotifyAlways,
        [int]$KeepReportDays = 0
    )
    $started = Get-Date
    Write-Log ('[cve] Monitor started: packagers {0}, site {1}' -f $PackagersRoot, $SiteCode) -Level INFO

    $items = @(Get-Packagers -Root $PackagersRoot | Where-Object { [string]$_.Script -like '*.ps1' -and -not ([string]$_.Status).StartsWith('Read error') })
    $Apps = @($Apps | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($Apps.Count -gt 0) {
        $wanted = @{}
        foreach ($a in $Apps) { $wanted[[System.IO.Path]::GetFileNameWithoutExtension([string]$a).ToLowerInvariant()] = $true }
        $items = @($items | Where-Object { $wanted.ContainsKey([System.IO.Path]::GetFileNameWithoutExtension([string]$_.Script).ToLowerInvariant()) })
    }
    Write-Log ('[cve] {0} application(s) in scope' -f $items.Count) -Level INFO

    $history = @{}
    try { $history = Read-PackagerHistory } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
    $field = { param($entry, $name) if ($null -eq $entry) { $null } elseif ($entry -is [hashtable]) { $entry[$name] } elseif ($entry.PSObject.Properties[$name]) { $entry.PSObject.Properties[$name].Value } else { $null } }

    $rows = @()
    foreach ($m in $items) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$m.Script)
        $entry = $(if ($history.ContainsKey($base)) { $history[$base] } else { $null })
        $cm = ''
        if ((& $field $entry 'LastCmVersion') -and [string](& $field $entry 'LastCmSite') -eq $SiteCode) { $cm = [string](& $field $entry 'LastCmVersion') }
        $cpe = $(if ($m.PSObject.Properties['CPE']) { [string]$m.CPE } else { '' })
        $cells = Get-CveGridCells -Summary (& $field $entry 'VulnSummary') -Cpe $cpe -CurrentVersion $cm -LatestVersion ([string](& $field $entry 'LastKnownVersion'))
        $rows += [pscustomobject]@{
            Application    = [string]$m.Application
            Vendor         = [string]$m.Vendor
            Script         = [string]$m.Script
            FullPath       = [string]$m.FullPath
            CMName         = [string]$m.CMName
            Cpe            = $cpe
            AdvisoryUrl    = $(if ($m.PSObject.Properties['SecurityAdvisoryUrl']) { [string]$m.SecurityAdvisoryUrl } else { '' })
            CurrentVersion = $cm
            LatestVersion  = [string](& $field $entry 'LastKnownVersion')
            Status         = 'Ready'
            CveCount       = $cells.Count
            CvePair        = $(if ($cells.Checked) { ConvertTo-CvePairKey -Current $cm -Latest ([string](& $field $entry 'LastKnownVersion')) } else { '' })
            Failed         = $false
        }
    }

    $failed = 0
    if (-not $SkipLatest) {
        foreach ($row in $rows) {
            Write-Log ('[cve] Latest: {0}' -f $row.Application) -Level INFO
            try {
                $prior = [string]$row.LatestVersion
                $latest = Invoke-PackagerGetLatestVersion -PackagerPath $row.FullPath -SiteCode $SiteCode -FileServerPath $FileServerPath -DownloadRoot $DownloadRoot -M365Channel $M365Channel -M365DeployMode $M365DeployMode
                $row.LatestVersion = $latest
                Clear-StaleCveCells -Row $row
                $result = $(if ($prior -and $prior -eq $latest) { 'NoChange' } else { 'Updated' })
                try { Update-PackagerHistory -PackagerName ([System.IO.Path]::GetFileNameWithoutExtension($row.Script)) -Event Checked -Version $latest -Result $result } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
            }
            catch {
                Write-Log ('[cve] Latest failed for {0}: {1}' -f $row.Application, $_.Exception.Message) -Level WARN
                $row.Status = 'Error'
                $row.Failed = $true
                $failed++
                try { Update-PackagerHistory -PackagerName ([System.IO.Path]::GetFileNameWithoutExtension($row.Script)) -Event Checked -Result Failed } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
            }
        }
    }

    if (-not $SkipMecm) {
        $cmFound = 0; $cmMissing = 0; $cmFailed = 0
        foreach ($row in $rows) {
            if ($row.Failed) { continue }
            $base = [System.IO.Path]::GetFileNameWithoutExtension($row.Script)
            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName -CMName $row.CMName
                if ($res.Found) {
                    $cmFound++
                    $row.CurrentVersion = [string]$res.SoftwareVersion
                    Clear-StaleCveCells -Row $row
                    try { Update-PackagerHistory -PackagerName $base -Event CmChecked -Version ([string]$res.SoftwareVersion) -Data @{ SiteCode = $SiteCode } } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
                }
                else {
                    $cmMissing++
                    $row.CurrentVersion = ''
                    Clear-StaleCveCells -Row $row
                    $row.Status = 'Not found in ConfigMgr'
                    try { Update-PackagerHistory -PackagerName $base -Event CmChecked -Version '' -Data @{ SiteCode = $SiteCode } } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
                }
            }
            catch {
                Write-Log ('[cve] ConfigMgr failed for {0}: {1}' -f $row.Application, $_.Exception.Message) -Level WARN
                $cmFailed++
                $row.Status = 'Error'
                $row.Failed = $true
                $failed++
            }
        }
        Write-Log ('[cve] ConfigMgr query: site {0}, provider {1}: {2} found, {3} not found, {4} failed' -f $SiteCode, $(if ($ProviderMachineName) { $ProviderMachineName } else { '(default)' }), $cmFound, $cmMissing, $cmFailed) -Level INFO
    }

    foreach ($row in $rows) {
        if ($row.Failed -or $row.Status -eq 'Not found in ConfigMgr') { continue }
        if ([string]$row.CurrentVersion -and [string]$row.LatestVersion) {
            $cmp = Compare-SemVer -A ([string]$row.CurrentVersion) -B ([string]$row.LatestVersion)
            $row.Status = $(if ($cmp -lt 0) { 'Update available' } elseif ($cmp -eq 0) { 'Up to date' } else { 'Current newer' })
        }
        elseif ([string]$row.LatestVersion) { $row.Status = 'Latest retrieved' }
    }
    $updates = @($rows | Where-Object { $_.Status -like 'Update available*' })
    Write-Log ('[cve] {0} application(s) with an update' -f $updates.Count) -Level INFO

    # A lookup is for a version ConfigMgr holds: with an update it says what
    # the update fixes, without one it lists what stays open. An application
    # that ConfigMgr does not hold is not looked up.
    $deployed = @($rows | Where-Object { -not $_.Failed -and (ConvertTo-CveVersionKey -Version ([string]$_.CurrentVersion)) })
    $lookedUp = 0; $cveTotal = 0; $fixedTotal = 0; $openTotal = 0; $kevTotal = 0
    $ownResults = @{}
    if (-not $SkipCve -and $deployed.Count -gt 0) {
        $kev = $null
        foreach ($row in $deployed) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($row.Script)
            if (-not [string]$row.Cpe) { $row.CveCount = 'CPE missing'; Write-Log ('[cve] {0}: CPE missing' -f $row.Application) -Level WARN; continue }
            try {
                if ($null -eq $kev) { $kev = Get-KevCatalog }
                $res = Get-ProductVulnerabilities -Cpe $row.Cpe -CurrentVersion $row.CurrentVersion -LatestVersion $row.LatestVersion -ApiKey $NvdApiKey -CacheTtlMinutes $CacheTtlMinutes -KevCatalog $kev -VendorAdvisoryUrl $row.AdvisoryUrl
                $summary = Get-VulnerabilitySummary -Result $res
                $cells = Get-CveGridCells -Summary $summary -Cpe $row.Cpe
                $row.CveCount = $cells.Count
                Set-CveCellsPair -Row $row
                if ($res.Error) { Write-Log ('[cve] {0}: {1}' -f $row.Application, $res.Error) -Level WARN; $failed++; continue }
                try { Update-PackagerHistory -PackagerName $base -Event VulnChecked -Data $summary } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
                $ownResults[[string]$row.Script] = $res
                $lookedUp++
                $cveTotal += [int]$res.Count
                $fixedTotal += [int]$res.FixedCount
                $openTotal += [int]$res.OpenCount
                $kevTotal += [int]$res.KevCount
                if ([string]$row.Status -eq 'Update available') { $row.Status = 'Update available (' + $cells.Text + ')' }
                Write-Log ('[cve] {0}: {1} ({2})' -f $row.Application, (ConvertTo-VulnerabilityResultText -Result $res), $res.Source) -Level INFO
            }
            catch {
                Write-Log ('[cve] CVE lookup failed for {0}: {1}' -f $row.Application, $_.Exception.Message) -Level WARN
                $row.CveCount = 'Error'
                $failed++
            }
        }
    }

    $reportRows = Get-CveReportRows -Rows $rows -Results $ownResults -MarkCached
    # A result that this run did not look up comes from the cache; the count
    # keeps it apart from the lookups of the run.
    $cachedCount = @($reportRows | Where-Object { $_.Result -and -not $_.FromRun }).Count
    $paths = Export-VulnerabilityReport -Rows $reportRows -Folder $ReportFolder -Formats $Formats -Name $ReportName -Title 'AppPackager CVE report'
    foreach ($p in @($paths.PSObject.Properties | ForEach-Object { [string]$_.Value })) { if ($p) { Write-Log ('[cve] Report: {0}' -f $p) -Level INFO } }

    $ended = Get-Date
    $summaryOut = [pscustomobject]@{
        Total    = $rows.Count
        Updates  = $updates.Count
        LookedUp = $lookedUp
        Cached   = $cachedCount
        Cves     = $cveTotal
        Fixed    = $fixedTotal
        Open     = $openTotal
        Kev      = $kevTotal
        Failed   = $failed
        Started  = $started.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Ended    = $ended.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Computer = $env:COMPUTERNAME
        Paths    = $paths
    }
    if (($WebhookUrl -or $DropFolder) -and ($NotifyAlways -or $updates.Count -gt 0)) {
        $null = Send-VulnerabilityNotification -Paths $paths -Summary $summaryOut -WebhookUrl $WebhookUrl -DropFolder $DropFolder
    }
    Remove-ExpiredCveReports -Folder $ReportFolder -KeepDays $KeepReportDays

    Write-Log ('[cve] Monitor finished: {0} application(s), {1} with an update, {2} looked up, {3} from the cache, {4} CVE(s) in the ConfigMgr versions, {5} fixed by updating, {6} open, {7} in KEV, {8} failed' -f $rows.Count, $updates.Count, $lookedUp, $cachedCount, $cveTotal, $fixedTotal, $openTotal, $kevTotal, $failed) -Level INFO
    return $summaryOut
}

# User-facing error text. These functions sit above the -BatchMode block, because the batch catch blocks call them.
function Get-HttpStatusCode {
    # The HTTP status code of a web error, or $null. A number is no exception text.
    param([Parameter(Mandatory)]$ErrorRecord)
    $response = $ErrorRecord.Exception.Response
    if ($null -eq $response) { return $null }
    return [int]$response.StatusCode
}

function Get-UserErrorText {
    # The text of an error that the window may show. A message that this tool throws itself (throw 'text') is
    # written for the user and passes through. Any other error (a ConfigMgr, WMI, file or .NET exception) is raw
    # exception text: it goes to the log file and the caller gets the fallback sentence.
    # Only a throw statement has the category OperationStopped. An engine error has another category.
    # An exception that the tool builds itself carries Data['UserMessage'] (or Data['WsusRefusal']): its text is for the
    # user too. PowerShell 7 gives a rethrown exception object no OperationStopped category, so the marker is needed.
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [Parameter(Mandatory)][string]$Fallback
    )
    $ex = $ErrorRecord.Exception
    if ($null -ne $ex -and $null -ne $ex.Data -and ($ex.Data.Contains('UserMessage') -or $ex.Data.Contains('WsusRefusal'))) { return [string]$ex.Message }
    if ($null -ne $ex -and $ErrorRecord.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::OperationStopped -and
        $ex.GetType().FullName -eq 'System.Management.Automation.RuntimeException' -and $null -eq $ex.InnerException) {
        return [string]$ex.Message
    }
    if ($null -ne $ex) { try { Write-Log ('Error detail kept out of the window: {0}' -f $ex.Message) -Level ERROR } catch { $null = $_ } }
    return $Fallback
}
if ($BatchMode) {
    if ($LogPath) { Initialize-Logging -LogPath $LogPath }

    if ($Apps -and $Apps.Count -eq 1 -and $Apps[0] -match ',') {
        $Apps = @($Apps[0] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if ($CveReport) {
        $prefs = if (Test-Path (Get-PreferencesPath)) { Read-Preferences } else { $null }
        $siteForMonitor = $SiteCode
        if ($prefs -and -not $PSBoundParameters.ContainsKey('SiteCode') -and $prefs.SiteCode) { $siteForMonitor = [string]$prefs.SiteCode }
        $providerForMonitor = $ProviderMachineName
        if ([string]::IsNullOrWhiteSpace($providerForMonitor) -and $prefs -and $prefs.ProviderMachineName) { $providerForMonitor = [string]$prefs.ProviderMachineName }
        $folderForMonitor = $ReportFolder
        if ([string]::IsNullOrWhiteSpace($folderForMonitor)) { $folderForMonitor = Get-CveReportFolder }
        $summary = Invoke-CveMonitor `
            -PackagersRoot $PackagersRoot `
            -Apps @($Apps) `
            -SiteCode $siteForMonitor `
            -ProviderMachineName $providerForMonitor `
            -DownloadRoot $(if ($prefs -and $prefs.DownloadRoot) { [string]$prefs.DownloadRoot } else { '' }) `
            -FileServerPath $(if ($prefs -and $prefs.FileShareRoot) { [string]$prefs.FileShareRoot } else { '' }) `
            -M365Channel $(if ($prefs) { [string]$prefs.M365Channel } else { '' }) `
            -M365DeployMode $(if ($prefs) { [string]$prefs.M365DeployMode } else { '' }) `
            -SkipLatest:$SkipLatest -SkipMecm:$SkipMecm -SkipCve:$SkipCve `
            -NvdApiKey (Get-CveNvdApiKey) -CacheTtlMinutes (Get-CveCacheTtlMinutes) `
            -ReportFolder $folderForMonitor -Formats $ReportFormats `
            -WebhookUrl $WebhookUrl -DropFolder $DropFolder -NotifyAlways:$NotifyAlways `
            -KeepReportDays $KeepReportDays
        if ($summary.Failed -gt 0) { exit 1 } else { exit 0 }
    }

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
        } catch { <# Reading or writing packager history is bookkeeping; the main result does not depend on it. #> $null = $_ }
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
        -DefaultTitleMode  (Get-DefaultTitleModeForContext -Prefs $prefs) `
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
# Log file of the window
# =============================================================================
# The window logs to a daily file under Logs. Get-UserErrorText and the background runspaces write the raw text of an
# error there, so that the window can say "See the log file." and the file holds the reason. -Attach appends.
$script:GuiLogPath = Join-Path (Join-Path $PSScriptRoot 'Logs') ('AppPackager-{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
try {
    Initialize-Logging -LogPath $script:GuiLogPath -Attach
    $env:APPPACKAGER_LOG_PATH = $script:GuiLogPath
}
catch { $script:GuiLogPath = $null }
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
$btnDiscover     = $window.FindName('btnDiscover')
$btnCheckCve     = $window.FindName('btnCheckCve')
$btnCveReport    = $window.FindName('btnCveReport')
$btnStage        = $window.FindName('btnStage')
$btnPackage      = $window.FindName('btnPackage')
$btnPublishIntune = $window.FindName('btnPublishIntune')
$btnPublishWsus  = $window.FindName('btnPublishWsus')
$btnWsusUpdates  = $window.FindName('btnWsusUpdates')
$btnAddInstaller = $window.FindName('btnAddInstaller')
$btnConvertCm    = $window.FindName('btnConvertCm')
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
$colCves         = $window.FindName('colCves')
$colCveChecked   = $window.FindName('colCveChecked')
$colCveEpss      = $window.FindName('colCveEpss')
$colCveUpdated   = $window.FindName('colCveUpdated')
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
$script:WorkflowButtons = @($btnFullRun, $btnCheckLatest, $btnCheckMECM, $btnDiscover, $btnConvertCm, $btnCheckCve, $btnStage, $btnPackage, $btnPublishIntune, $btnPublishWsus, $btnWsusUpdates, $btnCveReport, $btnWorkbench, $btnAddInstaller)
$script:OptionsButtons  = @($btnOptions)

Initialize-SuiteTheme -Window $window `
    -IsDarkGetter { $toggleTheme.IsOn -eq $true } `
    -WorkflowButtons $script:WorkflowButtons `
    -OptionsButtons $script:OptionsButtons `
    -LogLabel $lblLogOutput

# The sidebar styles hold no color. SuiteCommon sets the colors of the sidebar buttons for the start theme here and at each toggle.
Set-ButtonTheme -IsDark ($toggleTheme.IsOn -eq $true)


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
    if (-not [object]::ReferenceEquals($Dialog.Owner, $Owner)) { $Dialog.Owner = $Owner }
    try {
        $Dialog.WindowTitleBrush          = $Owner.WindowTitleBrush
        $Dialog.NonActiveWindowTitleBrush = $Owner.NonActiveWindowTitleBrush
        $Dialog.GlowBrush                 = $Owner.GlowBrush
        $Dialog.NonActiveGlowBrush        = $Owner.NonActiveGlowBrush
    } catch { <# Copying window brush colors to the dialog is cosmetic. #> $null = $_ }
}


# The styles that every dialog and the Application Workbench take from MainWindow.xaml, so that each style exists once.
# A dialog marks the place with <!--SHARED_STYLES--> in its resource dictionary. Expand-SharedStyle replaces the marker
# with the styles of the key list below and with the implicit item styles (ListBoxItem, ComboBoxItem, TreeViewItem, CheckBox, TextBox, ComboBox, DataGridColumnHeader).
# The functions hold no script variable but the cache, so that a window probe can publish them as they are.
function Get-SharedStyleXaml {
    if ($null -ne $script:SharedStyleText) { return $script:SharedStyleText }
    $keys = @('ThemedButtonTemplate', 'ThemedButtonStyle', 'DialogAccentButton', 'DialogButton', 'InlineButtonStyle', 'LinkButtonStyle', 'GridLinkButtonStyle', 'ThemedListBoxItemStyle', 'ThemedComboBoxItemStyle', 'ThemedTreeViewItemStyle', 'GridRowStyle', 'GridCellStyle', 'PaddedGridCellStyle')
    $types = @('ListBoxItem', 'ComboBoxItem', 'TreeViewItem', 'CheckBox', 'TextBox', 'ComboBox', 'DataGridColumnHeader')
    $base = if ($script:SharedStyleRootOverride) { $script:SharedStyleRootOverride } else { $PSScriptRoot }
    [xml]$main = Get-Content -LiteralPath (Join-Path $base 'MainWindow.xaml') -Raw
    $nodes = @($main.SelectNodes("//*[local-name()='ResourceDictionary']/*[local-name()='ControlTemplate' or local-name()='Style']") | Where-Object {
            $key = $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml')
            ($key -and ($keys -contains $key)) -or (-not $key -and ($types -contains $_.GetAttribute('TargetType')))
        })
    if ($nodes.Count -lt $keys.Count) { throw 'MainWindow.xaml does not hold every shared style.' }
    $script:SharedStyleText = (($nodes | ForEach-Object { $_.OuterXml }) -join [Environment]::NewLine)
    return $script:SharedStyleText
}

function Expand-SharedStyle {
    param([Parameter(Mandatory)][string]$Xaml)
    $marker = '<!--SHARED_STYLES-->'
    if (-not $Xaml.Contains($marker)) { return $Xaml }
    return $Xaml.Replace($marker, (Get-SharedStyleXaml))
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
    # A modeless dialog (the CVE detail window) stays open across the toggle:
    # re-apply the theme and the title, glow and inactive brushes of the owner.
    foreach ($owned in @($window.OwnedWindows)) {
        if ($owned -is [MahApps.Metro.Controls.MetroWindow]) { Set-DialogChromeFromOwner -Dialog $owned -Owner $window }
    }
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

# The CVE source submenu is rebuilt for the selected row each time the menu opens.
$menuOpenCveSources = New-Object System.Windows.Controls.MenuItem
$menuOpenCveSources.Header = "Open CVE Sources"

$menuCopyCveIds = New-Object System.Windows.Controls.MenuItem
$menuCopyCveIds.Header = "Copy CVE IDs"
$menuSep0 = New-Object System.Windows.Controls.Separator

$contextMenu.Items.Add($menuEditApplication) | Out-Null
$contextMenu.Items.Add($menuSep0) | Out-Null
$contextMenu.Items.Add($menuOpenLogFolder) | Out-Null
$contextMenu.Items.Add($menuOpenStagedFolder) | Out-Null
$contextMenu.Items.Add($menuOpenNetworkShare) | Out-Null
$contextMenu.Items.Add($menuSep1) | Out-Null
$contextMenu.Items.Add($menuCopyLatestVersion) | Out-Null
$contextMenu.Items.Add($menuOpenCveSources) | Out-Null
$contextMenu.Items.Add($menuCopyCveIds) | Out-Null

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
    Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $logFolder)
})

$menuOpenStagedFolder.Add_Click({
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    $dlRoot = $script:Prefs.DownloadRoot

    if ([string]::IsNullOrWhiteSpace($dlRoot)) {
        Add-LogLine -Message "Download root is not set. Set it in Options, Packaging, Folders."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row.FullPath)
    if ($info.DownloadSubfolder) {
        $targetPath = Join-Path $dlRoot $info.DownloadSubfolder

        $version = [string]$row.LatestVersion
        if (-not [string]::IsNullOrWhiteSpace($version)) {
            $versionPath = Join-Path $targetPath $version
            if (Test-Path -LiteralPath $versionPath) {
                Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $versionPath)
                return
            }
        }

        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $targetPath)
            return
        }
    }

    if (Test-Path -LiteralPath $dlRoot) {
        Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $dlRoot)
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
        Add-LogLine -Message "Content share is not set. Set it in Options, Connections, ConfigMgr."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row.FullPath)
    if ($info.VendorFolder -and $info.AppFolder) {
        $targetPath = Join-Path (Join-Path (Join-Path $fsPath "Applications") $info.VendorFolder) $info.AppFolder
        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $targetPath)
            return
        }
    }

    $appsRoot = Join-Path $fsPath "Applications"
    if (Test-Path -LiteralPath $appsRoot) {
        Start-Process "explorer.exe" -ArgumentList ('"{0}"' -f $appsRoot)
    }
    else {
        Add-LogLine -Message ("Network path not accessible: {0}" -f $appsRoot)
    }
})

$contextMenu.Add_Opened({
    $menuOpenCveSources.Items.Clear()
    $row = $dataGrid.SelectedItem
    $cpe = $(if ($row -and $row.PSObject.Properties['Cpe']) { [string]$row.Cpe } else { '' })
    if (-not $cpe) {
        $menuOpenCveSources.IsEnabled = $false
        $menuCopyCveIds.IsEnabled = $false
        return
    }
    $menuOpenCveSources.IsEnabled = $true
    $menuCopyCveIds.IsEnabled = $true
    $advisory = $(if ($row.PSObject.Properties['AdvisoryUrl']) { [string]$row.AdvisoryUrl } else { '' })
    $set = Get-VulnerabilityLinks -Cpe $cpe -VersionStart $(if ([string]$row.CurrentVersion) { [string]$row.CurrentVersion } else { [string]$row.LatestVersion }) -VendorAdvisoryUrl $advisory
    foreach ($name in $set.Product.Keys) {
        $item = New-Object System.Windows.Controls.MenuItem
        $item.Header = [string]$name
        $item.Tag = [string]$set.Product[$name]
        $item.ToolTip = [string]$set.Product[$name]
        $item.Add_Click({ param($s, $e) Open-ExternalUrl -Url ([string]$s.Tag) })
        [void]$menuOpenCveSources.Items.Add($item)
    }
})

$menuCopyCveIds.Add_Click({
    $row = $dataGrid.SelectedItem
    if (-not $row) { return }
    $res = Get-CveCachedResult -Row $row
    if (-not $res) { Add-LogLine -Message "No CVE lookup for this row yet. Run Check CVEs first."; return }
    $ids = @(foreach ($c in @($res.Cves)) { [string]$c.Id })
    try { [System.Windows.Clipboard]::SetText(($ids -join "`r`n")) } catch { <# Copying text to the clipboard is a best-effort convenience. #> $null = $_ }
    Add-LogLine -Message ('{0} CVE id(s) copied.' -f $ids.Count)
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
$script:ActionButtonsEnabled = $true
function Set-ActionButtonsEnabled {
    param([bool]$Enabled)
    $script:ActionButtonsEnabled = $Enabled
    $btnCheckLatest.IsEnabled   = $Enabled
    $btnCheckMECM.IsEnabled     = $Enabled
    $btnDiscover.IsEnabled      = $Enabled
    $btnConvertCm.IsEnabled     = $Enabled
    $btnCheckCve.IsEnabled      = $Enabled
    $btnCveReport.IsEnabled     = $Enabled
    $btnStage.IsEnabled         = $Enabled
    $btnPackage.IsEnabled       = $Enabled
    $btnPublishIntune.IsEnabled = $Enabled
    $btnPublishWsus.IsEnabled   = $Enabled
    $btnWsusUpdates.IsEnabled   = $Enabled
    $btnFullRun.IsEnabled       = $Enabled
    $btnOptions.IsEnabled       = $Enabled
    Update-SidebarForSystems
}

# =============================================================================
# Sidebar state for the configured systems
# -----------------------------------------------------------------------------
# Get-SidebarSystemState holds the decision and touches no WPF types, so it is
# exercised headlessly. Update-SidebarForSystems applies it and is the only
# writer of the publish buttons' tooltips; it runs at launch and after every
# path that can change the ConfigMgr, Intune or WSUS settings.
# =============================================================================
function Get-SidebarSystemState {
    param($Prefs)

    $systems = $Prefs.Systems
    $mecmInUse   = [bool](-not $systems -or $systems.ConfigMgr)
    $intuneInUse = [bool]($systems -and $systems.Intune)
    $wsusInUse   = [bool]($systems -and $systems.Wsus)
    $notInUse = 'not in use. Select it under Systems in Use in Options, Connections.'

    $consoleFound = [bool]($Prefs.DetectedTools -and $Prefs.DetectedTools.ConfigMgrConsole -and $Prefs.DetectedTools.ConfigMgrConsole.Found)
    $siteCode = [string]$Prefs.SiteCode
    $mecmReason = ''
    if (-not $mecmInUse) { $mecmReason = 'ConfigMgr is ' + $notInUse }
    elseif (-not $consoleFound) { $mecmReason = 'Needs the ConfigMgr console on this computer.' }
    elseif ([string]::IsNullOrWhiteSpace($siteCode)) { $mecmReason = 'Set the site code in Options, Connections, ConfigMgr.' }

    $intune = $Prefs.Intune
    $intuneReady = [bool]($intuneInUse -and $intune -and -not [string]::IsNullOrWhiteSpace([string]$intune.TenantId) -and
        -not [string]::IsNullOrWhiteSpace([string]$intune.ClientId) -and -not [string]::IsNullOrWhiteSpace([string]$intune.ClientSecretProtected))
    $intuneReason = $(if ($intuneInUse) { 'set Tenant ID, Client ID and Client secret in Options, Connections, Intune.' } else { 'Intune is ' + $notInUse })
    $wsusReady = [bool]($wsusInUse -and $Prefs.Wsus -and -not [string]::IsNullOrWhiteSpace([string]$Prefs.Wsus.ServerName))
    $wsusReason = $(if ($wsusInUse) { 'Set the WSUS server in Options, Connections, WSUS.' } else { 'WSUS is ' + $notInUse })

    # Discovery reads the site database over SQL, so the console is not needed.
    $sqlOverride = [bool](-not [string]::IsNullOrWhiteSpace([string]$Prefs.DiscoverySqlServer) -and -not [string]::IsNullOrWhiteSpace([string]$Prefs.DiscoverySqlDatabase))
    $discoverReason = ''
    if (-not $mecmInUse) { $discoverReason = 'ConfigMgr is ' + $notInUse }
    elseif ([string]::IsNullOrWhiteSpace($siteCode) -and -not $sqlOverride) { $discoverReason = 'Set the site code, or the SQL server and database, in Options, Connections, ConfigMgr.' }

    # The conversion reads the application with the console cmdlets. The dialog
    # asks for the site code, so a blank saved site code does not block it.
    $convertReason = ''
    if (-not $mecmInUse) { $convertReason = 'ConfigMgr is ' + $notInUse }
    elseif (-not $consoleFound) { $convertReason = 'Needs the ConfigMgr console on this computer.' }

    return @{
        ConvertCmEnabled     = (-not $convertReason)
        ConvertCmToolTip     = $(if ($convertReason) { 'Convert ConfigMgr App to Intune: ' + $convertReason } else { 'Convert an existing ConfigMgr application to Intune Win32 apps, one app for each deployment type. It reads ConfigMgr and changes nothing there.' })
        DiscoverEnabled      = (-not $discoverReason)
        DiscoverToolTip      = $(if ($discoverReason) { 'Discover unmanaged software: ' + $discoverReason } else { 'Find hidden apps that hardware inventory shows installed on clients and that have no ConfigMgr application' })
        CheckMecmEnabled     = (-not $mecmReason)
        CheckMecmToolTip     = $(if ($mecmReason) { 'Check ConfigMgr: ' + $mecmReason } else { 'Query ConfigMgr for the currently deployed version of each checked app' })
        MecmEnabled          = (-not $mecmReason)
        MecmToolTip          = $(if ($mecmReason) { 'Publish to ConfigMgr: ' + $mecmReason } else { 'Build a ConfigMgr Application and Deployment Type for every checked app' })
        IntuneEnabled        = $intuneReady
        IntuneToolTip        = $(if ($intuneReady) { 'Stage every checked app, build its .intunewin, and publish it to Intune' } else { 'Publish to Intune: ' + $intuneReason })
        WsusEnabled          = $wsusReady
        WsusToolTip          = $(if ($wsusReady) { 'Stage every checked app and publish its installer to WSUS as an update for computers that have an older version' } else { 'Publish to WSUS: ' + $wsusReason })
        WsusUpdatesEnabled   = $wsusReady
        WsusUpdatesToolTip   = $(if ($wsusReady) { 'List the updates AppPackager published to WSUS; approve, decline, expire or remove them, or import from the Microsoft Update Catalog' } else { 'WSUS Updates: ' + $wsusReason })
    }
}

# The footer shows the saved site code and the site server (SMS Provider) as read-only text. It stays hidden
# while ConfigMgr is not in the systems in use. The site text comes from SuiteCommon.
function Update-ApFooter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window text only.')]
    param()
    if (-not $window) { return }
    try {
        Update-SuiteFooter -Window $window -SiteCode ([string]$script:Prefs.SiteCode) -SiteServer ([string]$script:Prefs.ProviderMachineName)
        $systems = $script:Prefs.Systems
        $site = $window.FindName('txtFooterSite')
        if ($site) { $site.Visibility = $(if (-not $systems -or $systems.ConfigMgr) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }) }
    }
    catch { Add-LogLine -Message 'The footer was not updated.' }
}

function Update-SidebarForSystems {
    $state = Get-SidebarSystemState -Prefs $script:Prefs

    # Disabled rather than hidden: the button stays discoverable and its
    # tooltip states what to configure. A button disabled at launch for a
    # missing setting stays disabled after Setup or Options saves that
    # setting unless this assignment also enables it; a running pipeline
    # keeps every button disabled through $script:ActionButtonsEnabled.
    $pairs = @(
        @($btnCheckMECM, 'CheckMecm'), @($btnDiscover, 'Discover'), @($btnConvertCm, 'ConvertCm'), @($btnPackage, 'Mecm'), @($btnPublishIntune, 'Intune'),
        @($btnPublishWsus, 'Wsus'), @($btnWsusUpdates, 'WsusUpdates')
    )
    foreach ($pair in $pairs) {
        $button = $pair[0]
        $button.IsEnabled = ([bool]$script:ActionButtonsEnabled -and [bool]$state[$pair[1] + 'Enabled'])
        $button.ToolTip = $state[$pair[1] + 'ToolTip']
        # A disabled WPF button shows no tooltip unless told to.
        [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($button, $true)
    }
}

# --- One busy display for every long action -----------------------------------
# A run over a list of apps (Invoke-MultiAppPipeline) shows the overlay with
# "Verb N/M: name". An action on the UI thread uses the same overlay through
# these functions: Show-BusyIndicator, then Set-BusyStep per item, then
# Hide-BusyIndicator in a finally block. The wait cursor is part of the display
# and is never set alone. A dialog that the overlay cannot cover reports in its
# own status line through Set-DialogBusyStatus.
function Update-UiNow {
    # Lets WPF draw what was just changed before a long call blocks the thread.
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Render,
        [Action]{ }
    )
}

function Format-BusyStep {
    param([Parameter(Mandatory)][string]$Verb, [Parameter(Mandatory)][int]$Index, [Parameter(Mandatory)][int]$Total, [AllowEmptyString()][string]$Name = '')
    '{0} {1}/{2}: {3}' -f $Verb, $Index, $Total, $Name
}

function Show-BusyIndicator {
    param(
        [Parameter(Mandatory)][string]$Title,
        [string]$Step = 'Starting...',
        [switch]$AllowCancel
    )
    $txtProgressTitle.Text = $Title
    $txtProgressStep.Text = $Step
    # These actions cannot pause. Cancel stays only where the action can stop.
    $btnPausePipeline.Visibility = [System.Windows.Visibility]::Collapsed
    if ($AllowCancel) {
        $btnCancelPipeline.Visibility = [System.Windows.Visibility]::Visible
        $btnCancelPipeline.IsEnabled = $true
    }
    else {
        $btnCancelPipeline.Visibility = [System.Windows.Visibility]::Collapsed
    }
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    Update-UiNow
}

function Set-BusyStep {
    param([Parameter(Mandatory)][string]$Text)
    $txtProgressStep.Text = $Text
    Update-UiNow
}

function Hide-BusyIndicator {
    $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
    $window.Cursor = $null
    $btnPausePipeline.Visibility = [System.Windows.Visibility]::Visible
    $btnPausePipeline.Content = 'Pause'
    $btnPausePipeline.IsEnabled = $true
    $btnCancelPipeline.Visibility = [System.Windows.Visibility]::Visible
    $btnCancelPipeline.IsEnabled = $false
}

function Set-DialogBusyStatus {
    # For a dialog or option page: the status line says what runs, and the
    # wait cursor is an addition to it.
    param([Parameter(Mandatory)]$Dialog, [Parameter(Mandatory)]$Control, [Parameter(Mandatory)][string]$Text)
    $Dialog.Cursor = [System.Windows.Input.Cursors]::Wait
    $Control.Text = $Text
    Update-UiNow
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
    if ($script:DiscoveryJob) {
        Stop-DiscoveryScan
        return
    }
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
    $colCveChecked.Visibility  = $vis
    $colCveEpss.Visibility     = $vis
    $colCveUpdated.Visibility  = $vis
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
    Update-ApFooter
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
function New-AppFlowPanel {
    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" FontSize="12"
               Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}" Margin="0,0,0,12"
               Text="One Click opens a plan of the apps you track here, then runs Check (and optionally Stage / Publish) without further prompts. Each row publishes to its own destinations; a row with no box checked is skipped. Rows that keep the default follow it when the default changes. The sidebar publish buttons choose their own destination per run."/>
    <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Action on update" FontWeight="Bold" VerticalAlignment="Center" FontSize="12" Margin="0,0,8,0"/>
        <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}" Grid.Row="0" Grid.Column="1" x:Name="cboAction" Width="190" VerticalAlignment="Center">
            <ComboBoxItem Content="Report only"/>
            <ComboBoxItem Content="Stage"/>
            <ComboBoxItem Content="Stage and Publish"/>
        </ComboBox>
        <Controls:ToggleSwitch Grid.Row="0" Grid.Column="3" x:Name="toggleForce" IsOn="False"
                                Header="Force on launch (ignore cadence)"
                                OnContent="" OffContent="" MinWidth="0"
                                VerticalAlignment="Center"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Default destinations" VerticalAlignment="Center" FontSize="12" Margin="0,8,8,0"/>
        <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,8,0,0"
                    ToolTip="The destinations of every tracked app that has no selection of its own in the grid.">
            <CheckBox x:Name="chkOneClickMecm" Content="ConfigMgr" FontSize="12" Margin="0,0,16,0" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <CheckBox x:Name="chkOneClickWsus" Content="WSUS" FontSize="12" Margin="0,0,16,0" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <CheckBox x:Name="chkOneClickIntune" Content="Intune" FontSize="12" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
        <TextBlock Grid.Row="2" Grid.Column="0" Text="Existing ConfigMgr app" FontWeight="Bold" VerticalAlignment="Center" FontSize="12" Margin="0,8,8,0"/>
        <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}" Grid.Row="2" Grid.Column="1" x:Name="cboOneClickOnExisting" Width="190" VerticalAlignment="Center" Margin="0,8,0,0"
                  ToolTip="The action when the ConfigMgr application exists with another version.">
            <ComboBoxItem Content="Skip the application" Tag="Skip"/>
            <ComboBoxItem Content="Overwrite it" Tag="Overwrite"/>
        </ComboBox>
    </Grid>
    <DataGrid RowStyle="{DynamicResource GridRowStyle}" CellStyle="{DynamicResource GridCellStyle}" x:Name="dgApps" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
              GridLinesVisibility="Horizontal" HeadersVisibility="Column" RowHeaderWidth="0" BorderThickness="0"
              IsTextSearchEnabled="True" TextSearch.TextPath="Application">
        <DataGrid.Columns>
            <DataGridTemplateColumn MinWidth="82" Header="Track" Width="82" CanUserSort="True" SortMemberPath="Tracked">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <CheckBox IsChecked="{Binding Tracked, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}"
                                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn MinWidth="125" Header="Application" Width="*" Binding="{Binding Application}" IsReadOnly="True"/>
            <DataGridTextColumn MinWidth="94" Header="Vendor" Width="150" Binding="{Binding Vendor}" IsReadOnly="True"/>
            <DataGridTemplateColumn MinWidth="119" Header="ConfigMgr" Width="119" CanUserSort="True" SortMemberPath="ConfigMgr">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <CheckBox IsChecked="{Binding ConfigMgr, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTemplateColumn MinWidth="78" Header="WSUS" Width="78" CanUserSort="True" SortMemberPath="WSUS">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <CheckBox IsChecked="{Binding WSUS, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTemplateColumn MinWidth="89" Header="Intune" Width="89" CanUserSort="True" SortMemberPath="Intune">
                <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                        <CheckBox IsChecked="{Binding Intune, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn MinWidth="145" Header="Cadence (Days)" Width="145" Binding="{Binding CadenceDisplay, UpdateSourceTrigger=LostFocus, Mode=TwoWay}"/>
        </DataGrid.Columns>
    </DataGrid>
</DockPanel>
'@

    [xml]$xml = Expand-SharedStyle $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $cboAction   = $element.FindName('cboAction')
    $toggleForce = $element.FindName('toggleForce')
    $dgApps      = $element.FindName('dgApps')
    $chkOneClickMecm   = $element.FindName('chkOneClickMecm')
    $chkOneClickWsus   = $element.FindName('chkOneClickWsus')
    $chkOneClickIntune = $element.FindName('chkOneClickIntune')
    $cboOneClickOnExisting = $element.FindName('cboOneClickOnExisting')

    $currentPrefs = $script:Prefs.AppFlow
    $defaultSet = ConvertTo-OneClickDestinationSet -Value $currentPrefs.DefaultDestinations
    $chkOneClickMecm.IsChecked   = [bool]$defaultSet.ConfigMgr
    $chkOneClickWsus.IsChecked   = [bool]$defaultSet.WSUS
    $chkOneClickIntune.IsChecked = [bool]$defaultSet.Intune
    $cboOneClickOnExisting.SelectedIndex = $(if ([string]$currentPrefs.OnExisting -eq 'Overwrite') { 1 } else { 0 })
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

        $rowSet = Get-OneClickDestinations -Prefs $script:Prefs -PackagerName $base
        $rows.Add([pscustomobject]@{
            Packager       = $base
            Application    = $p.Application
            Vendor         = $p.Vendor
            Tracked        = $trackedSet.Contains($base)
            ConfigMgr      = [bool]$rowSet.ConfigMgr
            WSUS           = [bool]$rowSet.WSUS
            Intune         = [bool]$rowSet.Intune
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

        # A row that matches the default stores nothing, so it follows the
        # default when that changes later.
        $newDefault = [pscustomobject]@{ ConfigMgr = [bool]$chkOneClickMecm.IsChecked; WSUS = [bool]$chkOneClickWsus.IsChecked; Intune = [bool]$chkOneClickIntune.IsChecked }
        $destProps = [ordered]@{}
        foreach ($row in $rows) {
            if ([bool]$row.ConfigMgr -eq $newDefault.ConfigMgr -and [bool]$row.WSUS -eq $newDefault.WSUS -and [bool]$row.Intune -eq $newDefault.Intune) { continue }
            $destProps[$row.Packager] = [pscustomobject]@{ ConfigMgr = [bool]$row.ConfigMgr; WSUS = [bool]$row.WSUS; Intune = [bool]$row.Intune }
        }

        $prefsRef.AppFlow.Tracked             = $newTracked
        $prefsRef.AppFlow.Action              = $newAction
        $prefsRef.AppFlow.CadenceOverrides    = [pscustomobject]$overrideProps
        $prefsRef.AppFlow.ForceOnLaunch       = [bool]$toggleForce.IsOn
        $prefsRef.AppFlow.Destinations        = [pscustomobject]$destProps
        $prefsRef.AppFlow.DefaultDestinations = $newDefault
        $prefsRef.AppFlow.OnExisting          = [string]$cboOneClickOnExisting.SelectedItem.Tag
        $prefsRef.Intune.DeploymentTarget     = ConvertTo-DeploymentTargetName -Destinations $newDefault
        $prefsRef.Intune.PublishToIntune      = [bool]$newDefault.Intune
    }.GetNewClosure()

    return @{ Name = 'One Click'; Element = $element; Commit = $commit }
}

function Show-PreviewDialog {
    # Themed read-only preview window (monospaced, scrollable, Copy / Close).
    # Used by the previews of the Citrix Workspace App and Microsoft 365 Apps sections.
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <DockPanel Margin="12">
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnCopy"  Content="Copy" Margin="0,0,8,0" Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnClose" Content="Close" IsDefault="True" IsCancel="True" Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
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
    [xml]$xmlDoc = Expand-SharedStyle $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $pvWin = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $pvWin
    Set-DialogChromeFromOwner -Dialog $pvWin -Owner $Owner

    $txt   = $pvWin.FindName('txtContent')
    $copy  = $pvWin.FindName('btnCopy')
    $close = $pvWin.FindName('btnClose')
    $txt.Text = $Content

    $copy.Add_Click({
        try { [System.Windows.Clipboard]::SetText($txt.Text) } catch { <# Copying text to the clipboard is a best-effort convenience. #> $null = $_ }
    }.GetNewClosure())
    $close.Add_Click({ $pvWin.Close() }.GetNewClosure())

    [void]$pvWin.ShowDialog()
}

function New-ProductFilterPanel {
    $xaml = @'
<DockPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <Grid DockPanel.Dock="Top" Margin="0,0,0,12">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Not offered by Discover" FontSize="13" FontWeight="SemiBold" VerticalAlignment="Center"
                   ToolTip="Apps that Discover unmanaged software leaves out after Don't ask again for selected."/>
        <TextBlock Grid.Column="1" x:Name="txtDiscoveryDismissed" FontSize="13" VerticalAlignment="Center" Margin="0,0,12,0"/>
        <Button Grid.Column="2" x:Name="btnClearDiscoveryDismissed" Content="Clear List"
                Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                ToolTip="Offers each hidden app again in the next discovery."/>
    </Grid>
    <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Select which applications appear in the main grid. Uncheck to hide."
                   FontSize="12" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}" VerticalAlignment="Center" TextWrapping="Wrap"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Button x:Name="btnSelAll"  Content="Select All" Margin="6,0,0,0" Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnSelNone" Content="Select None" Margin="6,0,0,0" Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </Grid>
    <TreeView ItemContainerStyle="{DynamicResource ThemedTreeViewItemStyle}" x:Name="treeApps" />
</DockPanel>
'@

    [xml]$xml = Expand-SharedStyle $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    $treeApps  = $element.FindName('treeApps')
    $btnSelAll = $element.FindName('btnSelAll')
    $btnSelNone= $element.FindName('btnSelNone')
    $txtDiscoveryDismissed = $element.FindName('txtDiscoveryDismissed')
    $btnClearDiscoveryDismissed = $element.FindName('btnClearDiscoveryDismissed')
    $dismissedCount = @($script:Prefs.DiscoveryDismissed | Where-Object { $_ }).Count
    $txtDiscoveryDismissed.Text = ('{0} app(s)' -f $dismissedCount)
    $btnClearDiscoveryDismissed.IsEnabled = ($dismissedCount -gt 0)
    $discoveryClear = @{ Pending = $false }
    $btnClearDiscoveryDismissed.Add_Click({
        $discoveryClear.Pending = $true
        $txtDiscoveryDismissed.Text = '0 app(s) after OK'
        $btnClearDiscoveryDismissed.IsEnabled = $false
    }.GetNewClosure())

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
        if ($discoveryClear.Pending) { $prefsRef.DiscoveryDismissed = @() }
    }.GetNewClosure()

    return @{ Name = 'Product Filter'; Element = $element; Commit = $commit }
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
            <!--SHARED_STYLES-->
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
        <TextBlock Grid.Row="1" Text="Install Command:" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <TextBox   x:Name="txtInstall" Grid.Row="2" FontSize="12" FontFamily="Consolas" Margin="0,0,0,10"
                   Controls:TextBoxHelper.Watermark="install.bat (packager default)"/>
        <TextBlock Grid.Row="3" Text="Uninstall Command:" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <TextBox   x:Name="txtUninstall" Grid.Row="4" FontSize="12" FontFamily="Consolas" Margin="0,0,0,16"
                   Controls:TextBoxHelper.Watermark="uninstall.bat (packager default)"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnRevert" Content="Revert to Default" Style="{DynamicResource DialogButton}"/>
            <Button x:Name="btnSave"   Content="Save"              Style="{DynamicResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel"            Style="{DynamicResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = Expand-SharedStyle $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $Owner
    $dlg.Title = "Command Overrides - $AppLabel"
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
            <!--SHARED_STYLES-->
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
            <Button x:Name="btnSkip"      Content="Skip"          Style="{DynamicResource DialogButton}"/>
            <Button x:Name="btnOverwrite" Content="Overwrite"     Style="{DynamicResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel"    Content="Cancel Run"    Style="{DynamicResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = Expand-SharedStyle $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    $dlg.Title = 'Application Already Exists'
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
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
      VerticalScrollBarVisibility="Auto">
    <StackPanel>
        <TextBlock Text="About" FontSize="20" Margin="0,0,0,2"/>
        <TextBlock x:Name="txtAboutVersion" FontSize="13" Margin="0,0,0,14"
                   Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
        <TextBlock TextWrapping="Wrap" FontSize="12" Margin="0,0,0,14"
                   Text="Automated application packaging for ConfigMgr, Intune and WSUS: vendor version checks, silent install wrappers, detection rules, and a publish button per destination. Built entirely in in-box PowerShell 5.1."/>
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
                <Hyperlink x:Name="lnkAboutRepo" ToolTip="Open the project on GitHub.">
                    <Run x:Name="runAboutRepo" Text="github.com"/>
                </Hyperlink>
            </TextBlock>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="Last update check" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="txtAboutLastCheck" FontSize="12" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="Latest release" FontSize="12"/>
            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="txtAboutLatest" FontSize="12"/>
        </Grid>
        <StackPanel Orientation="Horizontal">
            <Button x:Name="btnAboutUpdate" Content="Update Now" Margin="0,0,8,0"
                    Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                    ToolTip="Install the latest release and restart AppPackager."/>
            <Button x:Name="btnAboutReleases" Content="Release Notes"
                    Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                    ToolTip="Open the releases page on GitHub."/>
        </StackPanel>
    </StackPanel>
</ScrollViewer>
'@

    [xml]$xml = Expand-SharedStyle $xaml
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
    $lnkAboutRepo.Add_Click({ try { Start-Process $repoUrl } catch { <# Opening a link or file in the shell is a best-effort convenience. #> $null = $_ } }.GetNewClosure())

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
    $btnAboutReleases.Add_Click({ try { Start-Process $releasesUrl } catch { <# Opening a link or file in the shell is a best-effort convenience. #> $null = $_ } }.GetNewClosure())

    return @{ Name = 'About'; Element = $element; Commit = { } }
}

# =============================================================================
# Options window - a tree of sections on the left, the section on the right.
# Connections, Packaging, Applications, Updates and CVEs, Security, Tools and
# About. The window and the Setup wizard share the field builders below.
# =============================================================================
# =============================================================================
# Options window: preference access, field editors, section builder
# -----------------------------------------------------------------------------
# A section is a list of field definitions (Get-OptionsSectionDefinition). One
# builder turns a field definition into a title, an editor and a tooltip, so
# the Options window and the Setup wizard share the same Connections fields
# and every field has a title in bold, a regular value and a tooltip.
# =============================================================================
function Get-OptionsPrefTarget {
    # Path 'Wsus.PortNumber' is a preference. 'cwa:Store.Name' and 'tv:ApiToken' are
    # the Citrix switches file and the TeamViewer Host file.
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Path)
    $root = $Context.Prefs
    $rest = $Path
    $aux = ''
    if ($Path -match '^(?<aux>cwa|tv):(?<rest>.+)$') {
        $aux = $Matches['aux']
        $rest = $Matches['rest']
        $root = $Context.Aux[$aux]
    }
    $names = @($rest -split '\.')
    $obj = $root
    for ($i = 0; $i -lt $names.Count - 1; $i++) {
        if ($null -eq $obj -or -not $obj.PSObject.Properties[$names[$i]]) { return $null }
        $obj = $obj.($names[$i])
    }
    if ($null -eq $obj) { return $null }
    return @{ Object = $obj; Name = $names[$names.Count - 1]; Aux = $aux }
}

function Get-OptionsPrefValue {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Path)
    $target = Get-OptionsPrefTarget -Context $Context -Path $Path
    if (-not $target -or -not $target.Object.PSObject.Properties[$target.Name]) { return $null }
    return $target.Object.($target.Name)
}

function Set-OptionsPrefValue {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Path, [AllowNull()]$Value)
    $target = Get-OptionsPrefTarget -Context $Context -Path $Path
    if (-not $target) { return }
    if ($target.Object.PSObject.Properties[$target.Name]) { $target.Object.($target.Name) = $Value }
    else { $target.Object | Add-Member -NotePropertyName $target.Name -NotePropertyValue $Value -Force }
    if ($target.Aux) { $Context.Touched[$target.Aux] = $true }
}

function New-OptionsContext {
    param([Parameter(Mandatory)]$Prefs)
    return @{
        Prefs    = $Prefs
        Aux      = @{ cwa = (Read-CwaSwitches); tv = (Read-TvHostConfig) }
        Touched  = @{}
        Owner    = $null
        Sections = @{}
        Packagers = $null
    }
}

function Get-OptionsPackagerList {
    # The catalog rows, read once for the window.
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($null -eq $Context.Packagers) { $Context.Packagers = @(Get-Packagers -Root $PackagersRoot) }
    return @($Context.Packagers)
}

function New-OptionsTextBlock {
    param([string]$Text, [double]$Size = 13, [bool]$Bold = $false, [bool]$Wrap = $true)
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.FontSize = $Size
    if ($Bold) { $tb.FontWeight = [System.Windows.FontWeights]::Bold }
    if ($Wrap) { $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap }
    return $tb
}

function New-OptionsButton {
    param([Parameter(Mandatory)][string]$Label, [string]$Tip = '')
    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $Label
    if ($Tip) { $btn.ToolTip = $Tip }
    $btn.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, 'InlineButtonStyle')
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btn, [System.Windows.Controls.CharacterCasing]::Normal)
    return $btn
}

function Select-OptionsFolder {
    param([string]$Current = '')
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($Current -and (Test-Path -LiteralPath $Current -PathType Container)) { $dlg.SelectedPath = $Current }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return [string]$dlg.SelectedPath }
    return $null
}

function ConvertTo-OptionsItemList {
    # A combo box item is a string, or a hashtable with Label and Value.
    param($Items)
    foreach ($item in @($Items | Where-Object { $null -ne $_ })) {
        if ($item -is [hashtable]) { [pscustomobject]@{ Label = [string]$item.Label; Value = [string]$item.Value } }
        else { [pscustomobject]@{ Label = [string]$item; Value = [string]$item } }
    }
}

# One place for the size of every input: a height that is the same everywhere, and a width that has a floor.
# A text field is Short (a code or a port), Medium (a name) or Long (a path, a key, a URL; the default).
function Get-OptionsInputHeight { 28 }
function Get-OptionsInputSizeName {
    # A text field is Short (a code or a port, set with Size = 'Short') or Long (a name, a path, a key, a URL). A declared Width never narrows it.
    param([Parameter(Mandatory)][hashtable]$Field)
    if ([string]$Field.Size -eq 'Short') { return 'Short' }
    return 'Long'
}

function Get-OptionsInputFloor {
    # The narrowest width that an input of this field may have.
    param([Parameter(Mandatory)][hashtable]$Field)
    switch ([string]$Field.Kind) {
        'Number' { return 80 }
        { $_ -in 'Combo', 'ComboText' } { return 200 }
    }
    switch (Get-OptionsInputSizeName -Field $Field) { 'Short' { return 80 } default { return 360 } }
}

function Get-OptionsInputWidth {
    param([Parameter(Mandatory)][hashtable]$Field)
    switch ([string]$Field.Kind) {
        'Number' { return 90 }
        { $_ -in 'Combo', 'ComboText' } {
            $wanted = $(if ($Field.Width) { [double]$Field.Width } else { 360 })
            return [Math]::Max(200, $wanted)
        }
    }
    switch (Get-OptionsInputSizeName -Field $Field) { 'Short' { return 80 } default { return 420 } }
}

function Get-OptionsInputHardMinimum {
    # The width that an input keeps when the window is at its smallest size. The floor of Get-OptionsInputFloor holds at the native size.
    param([Parameter(Mandatory)][hashtable]$Field)
    switch ([string]$Field.Kind) { 'Number' { return 80 } { $_ -in 'Combo', 'ComboText' } { return 120 } }
    if ((Get-OptionsInputSizeName -Field $Field) -eq 'Short') { return 80 }
    return 160
}

function Set-OptionsInputSize {
    # The input has a maximum width and shrinks when its row is narrower; the field row gives it a cell that is at most that wide.
    param([Parameter(Mandatory)]$Control, [Parameter(Mandatory)][hashtable]$Field, [switch]$Stretch)
    $Control.Height = (Get-OptionsInputHeight)
    $Control.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Center
    $Control.MinWidth = Get-OptionsInputHardMinimum -Field $Field
    $Control.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
    if (-not $Stretch) { $Control.MaxWidth = Get-OptionsInputWidth -Field $Field }
}

function Set-OptionsPanelInputSize {
    # The sections that a XAML panel builds take the same height and a floor.
    param([Parameter(Mandatory)]$Element)
    foreach ($child in @([System.Windows.LogicalTreeHelper]::GetChildren($Element))) {
        if ($child -is [System.Windows.Controls.DataGrid] -or $child -is [System.Windows.Controls.TreeView]) { continue }
        if ($child -is [System.Windows.Controls.TextBox] -or $child -is [System.Windows.Controls.PasswordBox] -or $child -is [System.Windows.Controls.ComboBox]) {
            $child.Height = (Get-OptionsInputHeight)
            $child.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Center
            $child.MinWidth = $(if ($child -is [System.Windows.Controls.ComboBox]) { 180 } else { 80 })
        }
        if ($child -is [System.Windows.DependencyObject]) { Set-OptionsPanelInputSize -Element $child }
    }
}

function New-OptionsFieldEditor {
    # Returns @{ Element; Focus; Get; Set } for one field definition. Get returns the
    # value as the field shows it; ConvertTo-OptionsPrefValue turns it into the stored type.
    param([Parameter(Mandatory)][hashtable]$Field, [Parameter(Mandatory)][hashtable]$Context)

    $tip = [string]$Field.Tip
    $kind = [string]$Field.Kind
    $editor = @{ Id = [string]$Field.Id; Field = $Field; Kind = $kind; Element = $null; Focus = $null; Get = $null; Set = $null; Buttons = @{} }

    switch ($kind) {
        'Text' {
            $tb = New-Object System.Windows.Controls.TextBox
            $tb.FontSize = 13
            $tb.ToolTip = $tip
            if ($Field.MaxLength) { $tb.MaxLength = [int]$Field.MaxLength }
            Set-OptionsInputSize -Control $tb -Field $Field
            if ($Field.Watermark) { [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($tb, [string]$Field.Watermark) }
            $editor.Element = $tb
            $editor.Focus = $tb
            $editor.Get = { $tb.Text.Trim() }.GetNewClosure()
            $editor.Set = { param($v) $tb.Text = [string]$v }.GetNewClosure()
        }
        'Number' {
            $tb = New-Object System.Windows.Controls.TextBox
            $tb.FontSize = 13
            $tb.ToolTip = $tip
            Set-OptionsInputSize -Control $tb -Field $Field
            $tb.MaxLength = $(if ($Field.MaxLength) { [int]$Field.MaxLength } else { 5 })
            $panel = New-Object System.Windows.Controls.StackPanel
            $panel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
            [void]$panel.Children.Add($tb)
            if ($Field.Unit) {
                $unit = New-OptionsTextBlock -Text (' ' + [string]$Field.Unit) -Wrap $false
                $unit.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                [void]$panel.Children.Add($unit)
            }
            $zeroIsEmpty = [bool]$Field.ZeroIsEmpty
            $editor.Element = $panel
            $editor.Focus = $tb
            $editor.Get = { $tb.Text.Trim() }.GetNewClosure()
            $editor.Set = { param($v) $n = 0; if ($zeroIsEmpty -and [int]::TryParse([string]$v, [ref]$n) -and $n -eq 0) { $tb.Text = '' } else { $tb.Text = [string]$v } }.GetNewClosure()
        }
        'Check' {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.FontSize = 13
            $cb.ToolTip = $tip
            $label = New-OptionsTextBlock -Text ([string]$Field.Label)
            $cb.Content = $label
            $editor.Element = $cb
            $editor.Focus = $cb
            $editor.Get = { $cb.IsChecked -eq $true }.GetNewClosure()
            $editor.Set = { param($v) $cb.IsChecked = [bool]$v }.GetNewClosure()
        }
        'Combo' {
            $cmb = New-Object System.Windows.Controls.ComboBox
            $cmb.FontSize = 13
            $cmb.ToolTip = $tip
            Set-OptionsInputSize -Control $cmb -Field $Field
            $cmb.SetResourceReference([System.Windows.Controls.ItemsControl]::ItemContainerStyleProperty, 'ThemedComboBoxItemStyle')
            $source = $Field.Items
            if ($Field.ItemsFrom) { $source = & $Field.ItemsFrom }
            foreach ($item in @(ConvertTo-OptionsItemList -Items $source)) {
                $ci = New-Object System.Windows.Controls.ComboBoxItem
                $ci.Content = $item.Label
                $ci.Tag = $item.Value
                [void]$cmb.Items.Add($ci)
            }
            $defaultValue = [string]$Field.Default
            $editor.Element = $cmb
            $editor.Focus = $cmb
            $editor.Get = { $sel = $cmb.SelectedItem; if ($sel) { [string]$sel.Tag } else { $defaultValue } }.GetNewClosure()
            $editor.Set = {
                param($v)
                $picked = $null
                foreach ($ci in $cmb.Items) { if ([string]$ci.Tag -ceq [string]$v) { $picked = $ci; break } }
                if (-not $picked) { foreach ($ci in $cmb.Items) { if ([string]$ci.Tag -ceq $defaultValue) { $picked = $ci; break } } }
                if (-not $picked -and $cmb.Items.Count -gt 0) { $picked = $cmb.Items[0] }
                $cmb.SelectedItem = $picked
            }.GetNewClosure()
        }
        'ComboText' {
            $cmb = New-Object System.Windows.Controls.ComboBox
            $cmb.FontSize = 13
            $cmb.ToolTip = $tip
            $cmb.IsEditable = $true
            Set-OptionsInputSize -Control $cmb -Field $Field
            $cmb.SetResourceReference([System.Windows.Controls.ItemsControl]::ItemContainerStyleProperty, 'ThemedComboBoxItemStyle')
            $editor.Element = $cmb
            $editor.Focus = $cmb
            $editor.Get = { ([string]$cmb.Text).Trim() }.GetNewClosure()
            $editor.Set = { param($v) $cmb.Text = [string]$v }.GetNewClosure()
            $editor.Combo = $cmb
        }
        'Secret' {
            $pwd = New-Object System.Windows.Controls.PasswordBox
            $pwd.FontSize = 13
            $pwd.ToolTip = $tip
            Set-OptionsInputSize -Control $pwd -Field $Field
            $pwdMax = $pwd.MaxWidth
            $pwd.MaxWidth = [double]::PositiveInfinity
            $hint = New-OptionsTextBlock -Text 'A value is saved. Leave this empty to keep it.' -Size 12
            $hint.Margin = New-Object System.Windows.Thickness(0, 4, 0, 0)
            $hint.Visibility = [System.Windows.Visibility]::Collapsed
            $panel = New-Object System.Windows.Controls.StackPanel
            [void]$panel.Children.Add($pwd)
            [void]$panel.Children.Add($hint)
            $panel.MaxWidth = $pwdMax
            $editor.Element = $panel
            $editor.Focus = $pwd
            $editor.Get = { $pwd.SecurePassword }.GetNewClosure()
            $editor.Set = { param($v) $hint.Visibility = $(if ([string]::IsNullOrWhiteSpace([string]$v)) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }) }.GetNewClosure()
        }
        { $_ -in 'File', 'Folder' } {
            $tb = New-Object System.Windows.Controls.TextBox
            $tb.FontSize = 13
            $tb.ToolTip = $tip
            if ($Field.MaxLength) { $tb.MaxLength = [int]$Field.MaxLength }
            Set-OptionsInputSize -Control $tb -Field $Field -Stretch
            $browse = New-OptionsButton -Label 'Browse...' -Tip ('Choose the ' + $(if ($kind -eq 'Folder') { 'folder' } else { 'file' }) + '.')
            $browse.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
            $grid = New-Object System.Windows.Controls.Grid
            $grid.MaxWidth = (Get-OptionsInputWidth -Field $Field) + 110
            $grid.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
            $col0 = New-Object System.Windows.Controls.ColumnDefinition
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = [System.Windows.GridLength]::Auto
            [void]$grid.ColumnDefinitions.Add($col0)
            [void]$grid.ColumnDefinitions.Add($col1)
            [System.Windows.Controls.Grid]::SetColumn($browse, 1)
            [void]$grid.Children.Add($tb)
            [void]$grid.Children.Add($browse)
            $filter = [string]$Field.Filter
            $dialogTitle = [string]$Field.Title
            $isFolder = ($kind -eq 'Folder')
            $browse.Add_Click({
                if ($isFolder) {
                    $picked = Select-OptionsFolder -Current $tb.Text
                    if ($picked) { $tb.Text = $picked }
                }
                else {
                    $open = New-Object Microsoft.Win32.OpenFileDialog
                    $open.Title = $dialogTitle
                    if ($filter) { $open.Filter = $filter }
                    if ($open.ShowDialog() -eq $true) { $tb.Text = $open.FileName }
                }
            }.GetNewClosure())
            $editor.Element = $grid
            $editor.Focus = $tb
            $editor.Get = { $tb.Text.Trim() }.GetNewClosure()
            $editor.Set = { param($v) $tb.Text = [string]$v }.GetNewClosure()
        }
        'Status' {
            $text = New-OptionsTextBlock -Text '' -Size 13
            $text.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            $text.ToolTip = $tip
            $grid = New-Object System.Windows.Controls.Grid
            $grid.MaxWidth = 700
            $col0 = New-Object System.Windows.Controls.ColumnDefinition
            [void]$grid.ColumnDefinitions.Add($col0)
            [void]$grid.Children.Add($text)
            $index = 1
            foreach ($b in @($Field.Buttons | Where-Object { $_ })) {
                $btn = New-OptionsButton -Label ([string]$b.Label) -Tip ([string]$b.Tip)
                $btn.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
                $btn.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $colN = New-Object System.Windows.Controls.ColumnDefinition
                $colN.Width = [System.Windows.GridLength]::Auto
                [void]$grid.ColumnDefinitions.Add($colN)
                [System.Windows.Controls.Grid]::SetColumn($btn, $index)
                [void]$grid.Children.Add($btn)
                $editor.Buttons[[string]$b.Id] = $btn
                $index++
            }
            $editor.Element = $grid
            $editor.Focus = $text
            $editor.Get = { $null }
            $editor.Set = { param($v) $text.Text = [string]$v }.GetNewClosure()
            $editor.Text = $text
        }
        'Actions' {
            $panel = New-Object System.Windows.Controls.WrapPanel
            foreach ($b in @($Field.Buttons | Where-Object { $_ })) {
                $btn = New-OptionsButton -Label ([string]$b.Label) -Tip ([string]$b.Tip)
                $btn.Margin = New-Object System.Windows.Thickness(0, 0, 8, 4)
                [void]$panel.Children.Add($btn)
                $editor.Buttons[[string]$b.Id] = $btn
            }
            $editor.Element = $panel
            $editor.Focus = $(if ($panel.Children.Count -gt 0) { $panel.Children[0] } else { $panel })
            $editor.Get = { $null }
            $editor.Set = { param($v) }
        }
        'Multi' {
            $grid = New-Object System.Windows.Controls.Grid
            $columns = $(if ($Field.Columns) { [int]$Field.Columns } else { 2 })
            for ($c = 0; $c -lt $columns; $c++) {
                $cd = New-Object System.Windows.Controls.ColumnDefinition
                $cd.Width = [System.Windows.GridLength]::Auto
                [void]$grid.ColumnDefinitions.Add($cd)
            }
            $boxes = [ordered]@{}
            $i = 0
            foreach ($item in @($Field.Items)) {
                $cb = New-Object System.Windows.Controls.CheckBox
                $cb.FontSize = 13
                $cb.Content = New-OptionsTextBlock -Text ([string]$item.Label)
                if ($item.Tip) { $cb.ToolTip = [string]$item.Tip } else { $cb.ToolTip = $tip }
                $cb.Margin = New-Object System.Windows.Thickness(0, 2, 18, 2)
                $row = [int][math]::Floor($i / $columns)
                while ($grid.RowDefinitions.Count -le $row) {
                    $rd = New-Object System.Windows.Controls.RowDefinition
                    $rd.Height = [System.Windows.GridLength]::Auto
                    [void]$grid.RowDefinitions.Add($rd)
                }
                [System.Windows.Controls.Grid]::SetColumn($cb, ($i % $columns))
                [System.Windows.Controls.Grid]::SetRow($cb, $row)
                [void]$grid.Children.Add($cb)
                $boxes[[string]$item.Id] = $cb
                $i++
            }
            $editor.Element = $grid
            $editor.Focus = $(if ($boxes.Count -gt 0) { @($boxes.Values)[0] } else { $grid })
            $editor.Get = { @($boxes.Keys | Where-Object { $boxes[$_].IsChecked -eq $true }) }.GetNewClosure()
            $editor.Set = { param($v) $set = @($v | ForEach-Object { [string]$_ }); foreach ($k in $boxes.Keys) { $boxes[$k].IsChecked = ($set -contains $k) } }.GetNewClosure()
            $editor.Boxes = $boxes
        }
        'Custom' {
            $built = & $Field.Build $Field $Context
            foreach ($key in $built.Keys) { $editor[$key] = $built[$key] }
        }
        default { throw ('Unknown Options field kind: ' + $kind) }
    }
    # The field id names the control for UI Automation, so that a capture or a test finds it by id.
    if ($editor.Focus -is [System.Windows.DependencyObject]) { [System.Windows.Automation.AutomationProperties]::SetAutomationId($editor.Focus, [string]$Field.Id) }
    foreach ($buttonId in @($editor.Buttons.Keys)) { [System.Windows.Automation.AutomationProperties]::SetAutomationId($editor.Buttons[$buttonId], ('{0}.{1}' -f [string]$Field.Id, $buttonId)) }
    return $editor
}

function New-OptionsFieldRow {
    param([Parameter(Mandatory)][hashtable]$Field, [Parameter(Mandatory)][hashtable]$Editor)
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $titleColumn = New-Object System.Windows.Controls.ColumnDefinition
    $titleColumn.Width = New-Object System.Windows.GridLength(220)
    $valueColumn = New-Object System.Windows.Controls.ColumnDefinition
    [void]$grid.ColumnDefinitions.Add($titleColumn)
    [void]$grid.ColumnDefinitions.Add($valueColumn)

    $title = New-OptionsTextBlock -Text ([string]$Field.Title) -Size 14 -Bold $true
    $title.Margin = New-Object System.Windows.Thickness(0, 4, 12, 0)
    $title.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
    $title.ToolTip = [string]$Field.Tip
    [System.Windows.Controls.Grid]::SetColumn($title, 0)
    [void]$grid.Children.Add($title)

    $element = $Editor.Element
    if ($Field.Indent) { $element.Margin = New-Object System.Windows.Thickness([double]$Field.Indent, 0, 0, 0) }
    if ([double]::IsInfinity($element.MaxWidth)) {
        $element.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        [System.Windows.Controls.Grid]::SetColumn($element, 1)
        [void]$grid.Children.Add($element)
    }
    else {
        # A cell that is at most as wide as the input and shrinks with the window; the input fills the cell.
        $cell = New-Object System.Windows.Controls.Grid
        $cellInput = New-Object System.Windows.Controls.ColumnDefinition
        $cellInput.Width = New-Object System.Windows.GridLength(1000, [System.Windows.GridUnitType]::Star)
        $cellInput.MaxWidth = $element.MaxWidth
        $cellRest = New-Object System.Windows.Controls.ColumnDefinition
        $cellRest.Width = New-Object System.Windows.GridLength(0.001, [System.Windows.GridUnitType]::Star)
        [void]$cell.ColumnDefinitions.Add($cellInput)
        [void]$cell.ColumnDefinitions.Add($cellRest)
        $element.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
        [void]$cell.Children.Add($element)
        [System.Windows.Controls.Grid]::SetColumn($cell, 1)
        [void]$grid.Children.Add($cell)
    }
    $Editor.Row = $grid
    return $grid
}

function ConvertTo-OptionsPrefValue {
    param([Parameter(Mandatory)][hashtable]$Field, [AllowNull()]$Raw)
    switch ([string]$Field.Kind) {
        'Number' {
            $text = ([string]$Raw).Trim()
            if (-not $text) { return [int]$Field.Default }
            return [int]$text
        }
        'Check' { return [bool]$Raw }
        'Multi' {
            if ($Field.Normalize) { return ,@(& $Field.Normalize $Raw) }
            return ,@($Raw)
        }
        default {
            if ($Raw -is [System.Array]) { return ,@($Raw) }
            return [string]$Raw
        }
    }
}

function Test-OptionsFieldInput {
    # Returns an error text, or $null when the field holds a value that can be stored.
    param([Parameter(Mandatory)][hashtable]$Editor)
    $field = $Editor.Field
    if ([string]$field.Kind -ne 'Number') { return $null }
    $text = [string](& $Editor.Get)
    if (-not $text) { return $null }
    $min = $(if ($null -ne $field.Min) { [int]$field.Min } else { 0 })
    $max = $(if ($null -ne $field.Max) { [int]$field.Max } else { 65535 })
    $number = 0
    if (-not [int]::TryParse($text, [ref]$number) -or $number -lt $min -or $number -gt $max) {
        return ('{0} must be a whole number from {1} to {2}.' -f [string]$field.Title, $min, $max)
    }
    return $null
}

function Connect-OptionsFieldLink {
    # Links between the editors of one group. The same links apply wherever the group appears.
    param([Parameter(Mandatory)][hashtable]$Group, [Parameter(Mandatory)][hashtable]$Context)
    $e = $Group.Editors

    # The port follows the SSL box until the operator types a port of their own.
    if ($e.Contains('wsus.port') -and $e.Contains('wsus.ssl')) {
        $portBox = $e['wsus.port'].Focus
        $sslBox = $e['wsus.ssl'].Focus
        $follow = @{ Auto = ([string](& $e['wsus.port'].Get) -in @('', '8530', '8531')) }
        $portBox.Add_TextChanged({ if ($portBox.IsKeyboardFocusWithin) { $follow.Auto = ($portBox.Text.Trim() -in @('', '8530', '8531')) } }.GetNewClosure())
        # Checked and Unchecked also fire for a keyboard, UI Automation or code change, not only a mouse click.
        $sslChanged = { if ($follow.Auto) { $portBox.Text = $(if ($sslBox.IsChecked -eq $true) { '8531' } else { '8530' }) } }.GetNewClosure()
        $sslBox.Add_Checked($sslChanged)
        $sslBox.Add_Unchecked($sslChanged)
    }

    # A name of the form server\instance,port in the SQL server field shows split into the three fields.
    if ($e.Contains('cm.sqlServer') -and $e.Contains('cm.sqlInstance') -and $e.Contains('cm.sqlPort')) {
        $parts = Split-CMSqlServerName -Name ([string](& $e['cm.sqlServer'].Get))
        if ($parts.Instance -or $parts.Port -gt 0) {
            & $e['cm.sqlServer'].Set $parts.Server
            if ($parts.Instance -and -not [string](& $e['cm.sqlInstance'].Get)) { & $e['cm.sqlInstance'].Set $parts.Instance }
            if ($parts.Port -gt 0 -and -not [string](& $e['cm.sqlPort'].Get)) { & $e['cm.sqlPort'].Set $parts.Port }
        }
    }

    # The target line shows the name that Discover connects to.
    if ($e.Contains('cm.sqlTarget')) {
        $targetText = $e['cm.sqlTarget'].Text
        $serverBox = $e['cm.sqlServer'].Focus
        $instanceBox = $e['cm.sqlInstance'].Focus
        $portBox = $e['cm.sqlPort'].Focus
        $databaseBox = $e['cm.sqlDatabase'].Focus
        $siteBox = $(if ($e.Contains('cm.siteCode')) { $e['cm.siteCode'].Focus } else { $null })
        $refresh = {
            $port = 0
            [void][int]::TryParse($portBox.Text.Trim(), [ref]$port)
            $name = Join-CMSqlServerName -Server $serverBox.Text -Instance $instanceBox.Text -Port $port
            if (-not $name) { $name = 'the SQL server in the site definition' }
            $database = $databaseBox.Text.Trim()
            if (-not $database) { $database = $(if ($siteBox -and $siteBox.Text.Trim()) { 'CM_' + $siteBox.Text.Trim().ToUpperInvariant() } else { 'CM_<site code>' }) }
            $targetText.Text = ('{0}, database {1}' -f $name, $database)
        }.GetNewClosure()
        foreach ($box in @($serverBox, $instanceBox, $portBox, $databaseBox, $siteBox)) { if ($box) { $box.Add_TextChanged($refresh) } }
        & $refresh
    }

    # Application title: the two modes exclude each other.
    if ($e.Contains('title.include') -and $e.Contains('title.remove')) {
        $includeBox = $e['title.include'].Focus
        $removeBox = $e['title.remove'].Focus
        $exclusive = { $removeBox.IsEnabled = ($includeBox.IsChecked -ne $true); if ($includeBox.IsChecked -eq $true) { $removeBox.IsChecked = $false } }.GetNewClosure()
        $includeBox.Add_Checked($exclusive)
        $includeBox.Add_Unchecked($exclusive)
        & $exclusive
        $Group.ExtraCommit.Add({
                param($ctx)
                $ctx.Prefs.RemoveVersionFromTitle = ($removeBox.IsChecked -eq $true -and $includeBox.IsChecked -ne $true)
            }.GetNewClosure()) | Out-Null
    }

    # The test deployment needs automatic distribution and a DP group, because it runs after the content distribution.
    if ($e.Contains('dist.auto') -and $e.Contains('dist.group') -and $e.Contains('dist.test')) {
        $autoBox = $e['dist.auto'].Focus
        $groupBox = $e['dist.group'].Focus
        $testBox = $e['dist.test'].Focus
        $collectionBox = $e['dist.collection'].Focus
        $createBox = $e['dist.create'].Focus
        $refreshTest = {
            $ready = ($autoBox.IsChecked -eq $true) -and -not [string]::IsNullOrWhiteSpace($groupBox.Text)
            $testBox.IsEnabled = $ready
            $collectionBox.IsEnabled = $ready -and ($testBox.IsChecked -eq $true)
            $createBox.IsEnabled = $ready -and ($testBox.IsChecked -eq $true)
        }.GetNewClosure()
        $autoBox.Add_Click($refreshTest)
        $autoBox.Add_Unchecked($refreshTest)
        $groupBox.Add_TextChanged($refreshTest)
        $testBox.Add_Click($refreshTest)
        & $refreshTest
    }
}

function New-OptionsFieldGroup {
    # Builds the rows of the given fields. Commit writes the values to the preferences, Validate
    # returns the first input error as @{ FieldId; Message }.
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][object[]]$Field,
        [string]$Description = ''
    )
    $stack = New-Object System.Windows.Controls.StackPanel
    $group = @{ Element = $stack; Editors = [ordered]@{}; Order = New-Object System.Collections.Generic.List[string]; ExtraCommit = New-Object System.Collections.Generic.List[scriptblock]; ExtraValidate = New-Object System.Collections.Generic.List[scriptblock] }

    if ($Description) {
        $desc = New-OptionsTextBlock -Text $Description -Size 12
        $desc.Margin = New-Object System.Windows.Thickness(0, 0, 0, 14)
        $desc.MaxWidth = 760
        $desc.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        [void]$stack.Children.Add($desc)
    }

    foreach ($f in $Field) {
        switch ([string]$f.Kind) {
            'Group' {
                $heading = New-OptionsTextBlock -Text ([string]$f.Title) -Size 15
                $heading.Margin = New-Object System.Windows.Thickness(0, 14, 0, 4)
                [void]$stack.Children.Add($heading)
                $rule = New-Object System.Windows.Controls.Border
                $rule.Height = 1
                $rule.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                $rule.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'MahApps.Brushes.Control.Border')
                [void]$stack.Children.Add($rule)
                if ($f.Description) {
                    $note = New-OptionsTextBlock -Text ([string]$f.Description) -Size 12
                    $note.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                    $note.MaxWidth = 760
                    $note.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
                    [void]$stack.Children.Add($note)
                }
            }
            'Note' {
                $note = New-OptionsTextBlock -Text ([string]$f.Text) -Size 12
                $note.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                $note.MaxWidth = 760
                $note.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
                [void]$stack.Children.Add($note)
                if ($f.Id) { $group.Editors[[string]$f.Id] = @{ Id = [string]$f.Id; Field = $f; Kind = 'Note'; Element = $note; Focus = $note; Get = { $null }; Set = { param($v) }; Text = $note } }
            }
            default {
                $editor = New-OptionsFieldEditor -Field $f -Context $Context
                $row = New-OptionsFieldRow -Field $f -Editor $editor
                [void]$stack.Children.Add($row)
                $group.Editors[[string]$f.Id] = $editor
                $group.Order.Add([string]$f.Id)
                if ($f.Pref -and $editor.Set) {
                    $stored = Get-OptionsPrefValue -Context $Context -Path ([string]$f.Pref)
                    if ($f.ToText -and $null -ne $stored) { $stored = & $f.ToText $stored }
                    if ([string]$f.Kind -eq 'Secret') { & $editor.Set $stored }
                    elseif ($null -ne $stored) { & $editor.Set $stored }
                    elseif ($null -ne $f.Default) { & $editor.Set $f.Default }
                    else { & $editor.Set $(if ([string]$f.Kind -eq 'Check') { $false } else { '' }) }
                }
                elseif ($null -ne $f.Default -and $editor.Set -and -not $f.Pref) { & $editor.Set $f.Default }
            }
        }
    }

    Connect-OptionsFieldLink -Group $group -Context $Context

    $group.Validate = {
        foreach ($id in $group.Order) {
            $message = Test-OptionsFieldInput -Editor $group.Editors[$id]
            if ($message) { return @{ FieldId = $id; Message = $message } }
        }
        foreach ($check in $group.ExtraValidate) {
            $result = & $check $Context
            if ($result) { return $result }
        }
        return $null
    }.GetNewClosure()

    $group.Commit = {
        foreach ($id in $group.Order) {
            $editor = $group.Editors[$id]
            $f = $editor.Field
            if ($editor.Commit) { & $editor.Commit $Context }
            if (-not $f.Pref -or $f.ReadOnly) { continue }
            if ([string]$f.Kind -eq 'Secret') {
                $secure = & $editor.Get
                if ($secure -and $secure.Length -gt 0) { Set-OptionsPrefValue -Context $Context -Path ([string]$f.Pref) -Value ($secure | ConvertFrom-SecureString) }
                continue
            }
            if ($null -eq $editor.Get) { continue }
            $raw = & $editor.Get
            if ($f.FromText) { $raw = & $f.FromText $raw }
            Set-OptionsPrefValue -Context $Context -Path ([string]$f.Pref) -Value (ConvertTo-OptionsPrefValue -Field $f -Raw $raw)
        }
        foreach ($extra in $group.ExtraCommit) { & $extra $Context }
    }.GetNewClosure()

    return $group
}

function Get-OptionsFieldTable {
    # Every field definition of every section by id, for the wizard and the tests.
    $table = @{}
    foreach ($section in @(Get-OptionsSectionDefinition)) {
        foreach ($f in @($section.Fields)) { if ($f.Id) { $table[[string]$f.Id] = $f } }
    }
    return $table
}

function New-OptionsSectionPage {
    # One section: its title, its description and the group of its fields, in a scroll viewer.
    param([Parameter(Mandatory)][hashtable]$Definition, [Parameter(Mandatory)][hashtable]$Context)

    if ($Definition.Kind -eq 'Panel') {
        $panel = & $Definition.Factory
        Set-OptionsPanelInputSize -Element $panel.Element
        $page = @{ Id = [string]$Definition.Id; Element = $panel.Element; Editors = @{}; Commit = $panel.Commit; Validate = $null; Panel = $panel }
        if ($Definition.SearchFields) {
            foreach ($sf in @($Definition.SearchFields)) {
                $control = $panel.Element.FindName([string]$sf.Control)
                if ($control) { $page.Editors[[string]$sf.Id] = @{ Id = [string]$sf.Id; Focus = $control } }
            }
        }
        return $page
    }

    $outer = New-Object System.Windows.Controls.ScrollViewer
    $outer.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $outer.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Disabled
    $content = New-Object System.Windows.Controls.StackPanel
    $content.Margin = New-Object System.Windows.Thickness(0, 0, 14, 12)
    $outer.Content = $content

    $heading = New-OptionsTextBlock -Text ([string]$Definition.Title) -Size 20
    $heading.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
    $heading.Tag = 'section-title'
    [void]$content.Children.Add($heading)

    $group = New-OptionsFieldGroup -Context $Context -Field @($Definition.Fields) -Description ([string]$Definition.Description)
    [void]$content.Children.Add($group.Element)

    $page = @{ Id = [string]$Definition.Id; Element = $outer; Editors = $group.Editors; Commit = $group.Commit; Validate = $group.Validate; Group = $group }
    if ($Definition.Post) { & $Definition.Post $Context $page }
    return $page
}

# =============================================================================
# Options window: search
# =============================================================================
function Get-OptionsSectionPath {
    param([Parameter(Mandatory)][object[]]$Definition, [Parameter(Mandatory)][string]$Id)
    $byId = @{}
    foreach ($d in $Definition) { $byId[[string]$d.Id] = $d }
    $names = New-Object System.Collections.Generic.List[string]
    $cursor = $byId[$Id]
    while ($cursor) {
        $names.Insert(0, [string]$cursor.Title)
        $cursor = $(if ($cursor.Parent) { $byId[[string]$cursor.Parent] } else { $null })
    }
    return ($names -join ' > ')
}

function Get-OptionsSearchIndex {
    # One entry per field and per section. A field of a section whose name appears in
    # the entry text is found by the section name too.
    param([Parameter(Mandatory)][object[]]$Definition)
    $index = New-Object System.Collections.Generic.List[object]
    foreach ($section in $Definition) {
        if ($section.Kind -eq 'Group') { continue }
        $path = Get-OptionsSectionPath -Definition $Definition -Id ([string]$section.Id)
        $index.Add([pscustomobject]@{ SectionId = [string]$section.Id; FieldId = ''; Title = [string]$section.Title; Path = $path; Text = ($path + ' ' + [string]$section.Keywords).ToLowerInvariant() })
        $entries = @()
        if ($section.Kind -eq 'Panel') { $entries = @($section.SearchFields | ForEach-Object { @{ Id = [string]$_.Id; Title = [string]$_.Title; Keywords = [string]$_.Keywords } }) }
        else { $entries = @($section.Fields | Where-Object { $_.Id -and $_.Title -and $_.Kind -notin 'Group', 'Note' } | ForEach-Object { @{ Id = [string]$_.Id; Title = [string]$_.Title; Keywords = [string]$_.Keywords } }) }
        foreach ($entry in $entries) {
            $index.Add([pscustomobject]@{ SectionId = [string]$section.Id; FieldId = $entry.Id; Title = $entry.Title; Path = $path; Text = ($entry.Title + ' ' + $entry.Keywords + ' ' + $path).ToLowerInvariant() })
        }
    }
    return $index.ToArray()
}

function Search-OptionsIndex {
    # Every word of the query must appear in the text of an entry. A title that starts with the query comes first.
    param([Parameter(Mandatory)][object[]]$Index, [AllowEmptyString()][string]$Query = '', [int]$Limit = 40)
    $words = @(([string]$Query).Trim().ToLowerInvariant() -split '\s+' | Where-Object { $_ })
    if ($words.Count -eq 0) { return @() }
    $hits = foreach ($entry in $Index) {
        $all = $true
        foreach ($w in $words) { if (-not $entry.Text.Contains($w)) { $all = $false; break } }
        if (-not $all) { continue }
        $rank = 3
        $title = $entry.Title.ToLowerInvariant()
        if ($title -eq ($words -join ' ')) { $rank = 0 }
        elseif ($title.StartsWith($words[0])) { $rank = 1 }
        elseif ($title.Contains($words[0])) { $rank = 2 }
        [pscustomobject]@{ Rank = $rank; Entry = $entry }
    }
    return @($hits | Sort-Object Rank, @{ Expression = { $_.Entry.Title } } | Select-Object -First $Limit | ForEach-Object { $_.Entry })
}

# =============================================================================
# Options window: section definitions
# -----------------------------------------------------------------------------
# One entry per node of the tree. Kind is Group (a node with children and no
# page), Fields (a page built from the field list) or Panel (a page built by a
# factory). The tree text, the page title and the search use the same Title.
# =============================================================================
function Get-OptionsNonUiPreference {
    # The preference keys that no field of the Options window edits, and who writes each one. A subtree
    # has one entry. The inventory test fails for a key that is in neither this list nor a field.
    return [ordered]@{
        'FirstRunCompleted'         = 'The Setup wizard writes it. The box Do not show this again and Save set it.'
        'DiscoveryEnabled'          = 'The Discover unmanaged software window writes it when an application is enabled.'
        'DeploymentConditions.Apps' = 'The conditions dialog of the main window edits it for each application.'
        'CommandOverrides.Apps'     = 'The command override dialog of the main window edits it for each application.'
        'WorkbenchDataRoot'         = 'File only. It sets the data folder of the Workbench; APP_PACKAGER_WORKBENCH_ROOT overrides it.'
        'cwa:Installation.EnableSSON' = 'It follows Single sign-on in Options, Applications, Citrix Workspace App, Installation.'
    }
}

function Get-OptionsSectionDefinition {
    $channelItems = @(@{ Label = 'Monthly Enterprise Channel'; Value = 'MonthlyEnterprise' }, @{ Label = 'Current Channel'; Value = 'Current' })
    $deployItems = @(@{ Label = 'Managed (offline source)'; Value = 'Managed' }, @{ Label = 'Online (CDN)'; Value = 'Online' })
    $notSet = @{ Label = '(not set)'; Value = '' }
    $cwaPolicyItems = @($notSet, @{ Label = 'S - HTTPS stores only'; Value = 'S' }, @{ Label = 'A - HTTPS and HTTP stores'; Value = 'A' }, @{ Label = 'N - none'; Value = 'N' })

    $sections = New-Object System.Collections.Generic.List[hashtable]

    # ---------------------------------------------------------------- Connections
    $sections.Add(@{ Id = 'connections'; Title = 'Connections'; Kind = 'Group' })

    $sections.Add(@{
            Id = 'systems'; Title = 'Systems in Use'; Parent = 'connections'; Kind = 'Fields'
            Description = 'Select the systems that AppPackager publishes to. The sidebar disables the button of a system that is not selected.'
            Fields = @(
                @{ Id = 'sys.configmgr'; Title = 'ConfigMgr'; Kind = 'Check'; Pref = 'Systems.ConfigMgr'; Label = 'Publish to a ConfigMgr site'; Tip = 'Enables the sidebar buttons that read or change the ConfigMgr site.'; Keywords = 'systems in use' }
                @{ Id = 'sys.intune'; Title = 'Intune'; Kind = 'Check'; Pref = 'Systems.Intune'; Label = 'Publish Win32 apps through Microsoft Graph'; Tip = 'Enables Publish to Intune; it needs the tenant ID, client ID and client secret.'; Keywords = 'systems in use' }
                @{ Id = 'sys.wsus'; Title = 'WSUS'; Kind = 'Check'; Pref = 'Systems.Wsus'; Label = 'Publish updates to a WSUS server'; Tip = 'Enables Publish to WSUS and WSUS Updates; it needs the WSUS server name.'; Keywords = 'systems in use' }
            )
        })

    $sections.Add(@{
            Id = 'cm'; Title = 'ConfigMgr'; Parent = 'connections'; Kind = 'Fields'
            Description = 'The site that AppPackager publishes to and reads from.'
            Post = { param($Context, $Page) Initialize-OptionsConfigMgrSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'cm.siteCode'; Title = 'Site code'; Kind = 'Text'; Pref = 'SiteCode'; MaxLength = 5; Size = 'Short'; Tip = 'The ConfigMgr site code, for example ABC.' }
                @{ Id = 'cm.provider'; Title = 'Site server / SMS Provider'; Kind = 'Text'; Pref = 'ProviderMachineName'; MaxLength = 200; Tip = 'The computer name of the SMS Provider of the site.'; Keywords = 'smsprovider provider site server' }
                @{ Id = 'cm.dbGroup'; Kind = 'Group'; Title = 'Site database'; Description = 'Discover unmanaged software reads this database. An empty field takes its value from the site definition.' }
                @{ Id = 'cm.sqlServer'; Title = 'SQL server'; Kind = 'Text'; Pref = 'DiscoverySqlServer'; MaxLength = 200; Tip = 'The host name of the SQL Server that holds the site database.'; Keywords = 'discovery sql discover' }
                @{ Id = 'cm.sqlInstance'; Title = 'SQL instance'; Kind = 'Text'; Pref = 'DiscoverySqlInstance'; MaxLength = 128; Tip = 'The named instance of the SQL Server; empty uses the default instance.'; Keywords = 'discovery sql discover' }
                @{ Id = 'cm.sqlPort'; Title = 'SQL port'; Kind = 'Number'; Pref = 'DiscoverySqlPort'; Min = 1; Max = 65535; ZeroIsEmpty = $true; Default = 0; Tip = 'The TCP port of the instance; empty uses port 1433 or SQL Server Browser.'; Keywords = 'discovery sql discover tcp' }
                @{ Id = 'cm.sqlDatabase'; Title = 'Database name'; Kind = 'Text'; Pref = 'DiscoverySqlDatabase'; MaxLength = 128; Tip = 'The site database name; empty uses CM_ and the site code.'; Keywords = 'discovery sql discover cm_' }
                @{ Id = 'cm.sqlTarget'; Title = 'Discover reads'; Kind = 'Status'; Tip = 'The SQL Server name and database that Discover unmanaged software connects to.'; Keywords = 'discovery sql target' }
                @{ Id = 'cm.shareGroup'; Kind = 'Group'; Title = 'Content source' }
                @{ Id = 'cm.fileShare'; Title = 'Content share UNC'; Kind = 'Text'; Pref = 'FileShareRoot'; MaxLength = 200; Tip = 'The UNC path of the share that holds the application content sources.'; Keywords = 'file share root unc' }
                @{ Id = 'cm.console'; Title = 'Console status'; Kind = 'Status'; Tip = 'Shows whether the ConfigMgr console is installed on this computer.'; Keywords = 'adminui' }
            )
        })

    $sections.Add(@{
            Id = 'wsus'; Title = 'WSUS'; Parent = 'connections'; Kind = 'Fields'
            Description = 'Publish to WSUS publishes the installer of each checked app as a locally published update. This computer needs the WSUS console or the RSAT WSUS tools at the version of the server, an account in the WSUS Administrators group, and file share access to the server. A computer that is not the WSUS server must trust the signing certificate.'
            Post = { param($Context, $Page) Initialize-OptionsWsusSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'wsus.server'; Title = 'Server'; Kind = 'Text'; Pref = 'Wsus.ServerName'; MaxLength = 255; Watermark = 'wsus01.contoso.com'; Tip = 'The WSUS server name, or the top-level software update point.'; Keywords = 'wsus server' }
                @{ Id = 'wsus.port'; Title = 'Port'; Kind = 'Number'; Pref = 'Wsus.PortNumber'; Min = 1; Max = 65535; Default = 8530; Tip = 'The WSUS port; 8530 for HTTP and 8531 for HTTPS unless the site uses another port.'; Keywords = 'wsus port 8530 8531' }
                @{ Id = 'wsus.ssl'; Title = 'Use SSL'; Kind = 'Check'; Pref = 'Wsus.UseSsl'; Label = 'Connect to the WSUS administration API over HTTPS'; Tip = 'Switches the port between 8530 and 8531 until you type a custom port.'; Keywords = 'wsus https tls' }
                @{ Id = 'wsus.certGroup'; Kind = 'Group'; Title = 'Signing certificate'; Description = 'Clients install a locally published update only when the signing certificate is in their Trusted Publishers store, and in Trusted Root Certification Authorities when it is self-signed. The policy Allow signed updates from an intranet Microsoft update service location must be enabled (AcceptTrustedPublisherCerts = 1 under HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate).' }
                @{ Id = 'wsus.cert'; Title = 'Certificate'; Kind = 'Status'; Tip = 'The signing certificate of the WSUS server; Test Connection reads it.'; Keywords = 'wsus signing certificate' }
                @{ Id = 'wsus.certActions'; Title = 'Certificate actions'; Kind = 'Actions'; Tip = 'Create, import or export the signing certificate of the WSUS server.'; Keywords = 'pfx self-signed export'; Buttons = @(
                        @{ Id = 'create'; Label = 'Create self-signed...'; Tip = 'Has the WSUS server create and use a self-signed signing certificate.' }
                        @{ Id = 'import'; Label = 'Import PFX...'; Tip = 'Imports a code-signing certificate from a PFX file.' }
                        @{ Id = 'export'; Label = 'Export certificate...'; Tip = 'Saves the public signing certificate as a .cer file for Group Policy.' }) }
                @{ Id = 'wsus.pubGroup'; Kind = 'Group'; Title = 'Publishing'; Description = 'Each update installs only where an older version of the application is installed. A computer without the application does not get the update.' }
                @{ Id = 'wsus.classification'; Title = 'Classification'; Kind = 'Combo'; Pref = 'Wsus.Classification'; Width = 260; Default = 'Updates'; ItemsFrom = { @(Get-WsusClassificationNameList | ForEach-Object { @{ Label = (Get-WsusClassificationLabel -Name $_); Value = $_ } }) }; Tip = 'The WSUS classification of the published updates.' }
                @{ Id = 'wsus.group'; Title = 'Approve for group'; Kind = 'ComboText'; Pref = 'Wsus.ApprovalGroup'; Tip = 'The computer group that gets the update at once; empty approves later.'; Keywords = 'computer group approval' }
                @{ Id = 'wsus.groupLoad'; Title = 'Computer groups'; Kind = 'Actions'; Tip = 'Loads the computer groups of the WSUS server into the group list.'; Buttons = @(@{ Id = 'load'; Label = 'Load groups'; Tip = 'Reads the computer groups from the WSUS server.' }) }
                @{ Id = 'wsus.decline'; Title = 'Decline superseded'; Kind = 'Check'; Pref = 'Wsus.DeclineSuperseded'; Label = 'Decline earlier versions of the same application after publishing'; Tip = 'Declines the earlier versions of the application in WSUS.'; Keywords = 'supersedence' }
            )
        })

    $sections.Add(@{
            Id = 'intune'; Title = 'Intune'; Parent = 'connections'; Kind = 'Fields'
            Description = 'The app registration that Publish to Intune uses with Microsoft Graph. Stage builds an .intunewin beside the content when the Content Prep Tool is detected.'
            Fields = @(
                @{ Id = 'int.tenant'; Title = 'Tenant ID'; Kind = 'Text'; Pref = 'Intune.TenantId'; Tip = 'The Entra tenant ID, as a GUID or a domain name.'; Keywords = 'entra azure' }
                @{ Id = 'int.client'; Title = 'Client ID'; Kind = 'Text'; Pref = 'Intune.ClientId'; Tip = 'The client ID of an app registration with DeviceManagementApps.ReadWrite.All.'; Keywords = 'app registration graph' }
                @{ Id = 'int.secret'; Title = 'Client secret'; Kind = 'Secret'; Pref = 'Intune.ClientSecretProtected'; Tip = 'Stored with DPAPI for the current Windows user; empty keeps the saved secret.'; Keywords = 'dpapi' }
            )
        })

    $sections.Add(@{
            Id = 'test'; Title = 'Test Connection'; Parent = 'connections'; Kind = 'Fields'
            Description = 'Each button tests the values that are in the fields now, before you select OK.'
            Post = { param($Context, $Page) Initialize-OptionsTestSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'test.configmgr'; Title = 'ConfigMgr'; Kind = 'Status'; Tip = 'Reads the site from the SMS Provider and opens the site database.'; Keywords = 'test connection site provider sql'; Buttons = @(@{ Id = 'run'; Label = 'Test'; Tip = 'Tests the SMS Provider and the site database.' }) }
                @{ Id = 'test.wsus'; Title = 'WSUS'; Kind = 'Status'; Tip = 'Connects to the WSUS server and reads its version, role and signing certificate.'; Keywords = 'test connection'; Buttons = @(@{ Id = 'run'; Label = 'Test'; Tip = 'Tests the WSUS server.' }) }
                @{ Id = 'test.intune'; Title = 'Intune'; Kind = 'Status'; Tip = 'Requests a Microsoft Graph token with the tenant, client and secret.'; Keywords = 'test connection graph token'; Buttons = @(@{ Id = 'run'; Label = 'Test'; Tip = 'Tests the app registration.' }) }
            )
        })

    # ---------------------------------------------------------------- Packaging
    $sections.Add(@{ Id = 'packaging'; Title = 'Packaging'; Kind = 'Group' })

    $sections.Add(@{
            Id = 'folders'; Title = 'Folders'; Parent = 'packaging'; Kind = 'Fields'
            Post = { param($Context, $Page) Initialize-OptionsFoldersSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'pkg.downloadRoot'; Title = 'Download root'; Kind = 'Folder'; Pref = 'DownloadRoot'; MaxLength = 200; Tip = 'The local folder where Stage downloads and builds the installer content.'; Keywords = 'stage download folder' }
                @{ Id = 'pkg.packagers'; Title = 'Packagers folder'; Kind = 'Status'; Tip = 'The folder that holds the package-*.ps1 files of the catalog.'; Keywords = 'catalog'; Buttons = @(@{ Id = 'open'; Label = 'Open'; Tip = 'Opens the packagers folder in Explorer.' }) }
                @{ Id = 'pkg.layout'; Title = 'Content layout'; Kind = 'Combo'; Pref = 'ContentLayout'; Default = 'Nested'; Items = @(@{ Label = 'Nested - Applications\Vendor\App\Version'; Value = 'Nested' }, @{ Label = 'Flat - Applications\Vendor-App-Version'; Value = 'Flat' }); Tip = 'Nested keeps the versions of an app together; Flat uses one folder for each package.'; Keywords = 'share folder structure' }
            )
        })

    $sections.Add(@{
            Id = 'dtdefaults'; Title = 'Deployment Type Defaults'; Parent = 'packaging'; Kind = 'Fields'
            Description = 'The values of every deployment type that Publish to ConfigMgr creates.'
            Fields = @(
                @{ Id = 'pkg.estimated'; Title = 'Estimated runtime'; Kind = 'Number'; Pref = 'EstimatedRuntimeMins'; Min = 1; Max = 9999; Default = 15; MaxLength = 4; Unit = 'minutes'; Tip = 'The estimated installation time of a new deployment type, in minutes.' }
                @{ Id = 'pkg.maximum'; Title = 'Maximum runtime'; Kind = 'Number'; Pref = 'MaximumRuntimeMins'; Min = 1; Max = 9999; Default = 30; MaxLength = 4; Unit = 'minutes'; Tip = 'The maximum allowed run time of a new deployment type, in minutes.' }
                @{ Id = 'pkg.fallback'; Title = 'Fallback distribution points'; Kind = 'Combo'; Pref = 'ContentFallback'; Default = 'Allow'; Items = @(@{ Label = 'Allow - use the site default boundary group'; Value = 'Allow' }, @{ Label = 'Deny - current and neighbor groups only'; Value = 'Deny' }); Tip = 'Controls whether clients fall back to the default site boundary group for content.' }
                @{ Id = 'pkg.neighbor'; Title = 'Neighbor/default DP'; Kind = 'Combo'; Pref = 'SlowNetworkDeploymentMode'; Default = 'Download'; Items = @(@{ Label = 'Download content and install'; Value = 'Download' }, @{ Label = 'Do not download content'; Value = 'DoNothing' }); Tip = 'The content option for a distribution point in a neighbor or default boundary group.' }
            )
        })

    $sections.Add(@{
            Id = 'title'; Title = 'Application Title'; Parent = 'packaging'; Kind = 'Fields'
            Description = 'The naming rule of every Package run. Without a selection, a packager uses its own name.'
            Fields = @(
                @{ Id = 'title.include'; Title = 'Include version'; Kind = 'Check'; Pref = 'IncludeVersionInTitle'; Label = 'Add the version to each application name'; Tip = 'Creates a separate application for each release.' }
                @{ Id = 'title.remove'; Title = 'Remove version'; Kind = 'Check'; Pref = 'RemoveVersionFromTitle'; Label = 'Remove the release version from the vendor product name'; Tip = 'Keeps one application for all releases; it applies while Include version is off.' }
            )
        })

    $sections.Add(@{
            Id = 'postpackaging'; Title = 'Post Packaging'; Parent = 'packaging'; Kind = 'Fields'
            Description = 'The steps that follow Publish to ConfigMgr. The test deployment needs automatic distribution and a DP group.'
            Fields = @(
                @{ Id = 'dist.auto'; Title = 'Distribute content'; Kind = 'Check'; Pref = 'ContentDistribution.AutoDistribute'; Label = 'Run Start-CMContentDistribution after Package'; Tip = 'Distributes the content to the DP group when Package finishes.' }
                @{ Id = 'dist.group'; Title = 'DP group'; Kind = 'Text'; Pref = 'ContentDistribution.DPGroupName'; MaxLength = 200; Tip = 'The exact name of the distribution point group.'; Keywords = 'distribution point group' }
                @{ Id = 'dist.test'; Title = 'Test deployment'; Kind = 'Check'; Pref = 'ContentDistribution.DeployToTestCollection'; Label = 'Deploy to the test collection after distribution'; Tip = 'Creates an Available deployment for the test collection.' }
                @{ Id = 'dist.collection'; Title = 'Test collection'; Kind = 'Text'; Pref = 'ContentDistribution.TestCollectionName'; MaxLength = 255; Tip = 'The exact name of the device collection that receives the test deployment.' }
                @{ Id = 'dist.create'; Title = 'Create collection'; Kind = 'Check'; Pref = 'ContentDistribution.CreateTestCollectionIfMissing'; Label = 'Create the collection if it does not exist'; Tip = 'Creates the collection with All Systems as the limiting collection.' }
            )
        })

    # ---------------------------------------------------------------- Applications
    $sections.Add(@{ Id = 'applications'; Title = 'Applications'; Kind = 'Group' })

    $sections.Add(@{
            Id = 'adobe'; Title = 'Adobe Acrobat Reader'; Parent = 'applications'; Kind = 'Fields'
            Post = { param($Context, $Page) Initialize-OptionsAdobeSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'adobe.edition'; Title = 'Edition'; Kind = 'Combo'; Pref = 'AdobeReaderInstallOptions.Edition'; Default = 'English'; Width = 260; Items = @(@{ Label = 'English'; Value = 'English' }, @{ Label = 'Multilingual (MUI)'; Value = 'MUI' }); Tip = 'English stages the en_US installer; Multilingual stages the MUI installer.' }
                @{ Id = 'adobe.languages'; Title = 'Languages'; Kind = 'Multi'; Pref = 'AdobeReaderInstallOptions.Languages'; Columns = 6
                    Tip = 'The MUI languages to install; English is always included (LANG_LIST).'
                    Normalize = { param($Values) if (@($Values) -contains 'All') { @('All') } else { @($Values) } }
                    Items = @(@{ Id = 'All'; Label = 'All languages'; Tip = 'Installs every language of the MUI installer (LANG_LIST=ALL).' }) + @('ca_ES', 'cs_CZ', 'da_DK', 'de_DE', 'es_ES', 'eu_ES', 'fi_FI', 'fr_FR', 'hr_HR', 'hu_HU', 'it_IT', 'ja_JP', 'ko_KR', 'nb_NO', 'nl_NL', 'pl_PL', 'pt_BR', 'ro_RO', 'ru_RU', 'sk_SK', 'sl_SI', 'sv_SE', 'tr_TR', 'uk_UA', 'zh_CN', 'zh_TW' | ForEach-Object { @{ Id = $_; Label = $_ } }) }
                @{ Id = 'adobe.eula'; Title = 'License agreement'; Kind = 'Check'; Pref = 'AdobeReaderInstallOptions.AcceptEula'; Label = 'Accept the Adobe license agreement (EULA_ACCEPT=YES)'; Tip = 'Accepts the license for all users, so the dialog does not show at first start.' }
                @{ Id = 'adobe.shortcut'; Title = 'Desktop shortcut'; Kind = 'Check'; Pref = 'AdobeReaderInstallOptions.DisableDesktopShortcut'; Label = 'Do not create the desktop shortcut (DISABLEDESKTOPSHORTCUT=1)'; Tip = 'Skips the desktop shortcut at install and after every update.' }
                @{ Id = 'adobe.chrome'; Title = 'Chrome extension'; Kind = 'Check'; Pref = 'AdobeReaderInstallOptions.DisableChromeExtension'; Label = 'Do not install the Chrome extension (ENABLE_CHROMEEXT=0)'; Tip = 'Blocks the Adobe Chrome extension at install, for sites that set extensions by policy.' }
                @{ Id = 'adobe.browser'; Title = 'Browser viewer control'; Kind = 'Check'; Pref = 'AdobeReaderInstallOptions.DisableBrowserIntegration'; Label = 'Do not install the browser viewer control (DISABLE_BROWSER_INTEGRATION=YES)'; Tip = 'Skips the PDF viewer ActiveX control for Internet Explorer.' }
                @{ Id = 'adobe.update'; Title = 'Update mode'; Kind = 'Combo'; Pref = 'AdobeReaderInstallOptions.UpdateMode'; Default = ''; Width = 460; Items = @($notSet, @{ Label = '0 - check and install manually'; Value = '0' }, @{ Label = '2 - download automatically, user chooses when to install'; Value = '2' }, @{ Label = '3 - install updates automatically'; Value = '3' }, @{ Label = '4 - notify, user chooses to download and install'; Value = '4' }); Tip = 'Sets how Reader updates itself (UPDATE_MODE); use 0 when ConfigMgr patches Reader.' }
            )
        })

    $sections.Add(@{
            Id = 'bc5'; Title = 'Beyond Compare 5'; Parent = 'applications'; Kind = 'Fields'
            Fields = @(
                @{ Id = 'bc.key'; Title = 'License key file'; Kind = 'File'; Pref = 'BeyondCompareKeyFile'; Filter = 'License key file (BC5Key.txt)|BC5Key.txt|Text files (*.txt)|*.txt|All files (*.*)|*.*'; Tip = 'The BC5Key.txt file from Scooter Software; Stage and Package require it.'; Keywords = 'bc5key license' }
                @{ Id = 'bc.remote'; Title = 'Remote profiles'; Kind = 'Check'; Pref = 'BeyondCompareOptions.DisableRemoteProfiles'; Label = 'Turn off FTP, SFTP, WebDAV and cloud profiles (/DisableRemoteProfiles)'; Tip = 'Needs Beyond Compare 5.2 or later.' }
                @{ Id = 'bc.passwords'; Title = 'Saved passwords'; Kind = 'Check'; Pref = 'BeyondCompareOptions.DisableSavedPasswords'; Label = 'Do not save passwords on disk (/DisableSavedPasswords)'; Tip = 'Needs Beyond Compare 5.2 or later.' }
            )
        })

    $cwaComponents = @(
        @{ Name = 'ReceiverInside'; Label = 'ReceiverInside (core SDK)'; Tip = 'The core Workspace SDK services; required.'; Required = $true }
        @{ Name = 'ICA_Client'; Label = 'ICA_Client (HDX engine)'; Tip = 'Session launch and ICA protocol handling; required.'; Required = $true }
        @{ Name = 'AM'; Label = 'AM (authentication)'; Tip = 'The user authentication manager; required.'; Required = $true }
        @{ Name = 'SelfService'; Label = 'SelfService (self-service UI)'; Tip = 'The native application launch and self-service plug-in.' }
        @{ Name = 'DesktopViewer'; Label = 'DesktopViewer (virtual desktop)'; Tip = 'The virtual desktop UI framework.' }
        @{ Name = 'WebHelper'; Label = 'WebHelper (browser helper)'; Tip = 'Browser-to-application connectivity.' }
        @{ Name = 'BCR_Client'; Label = 'BCR_Client (browser content redirection)'; Tip = 'Redirects browser content rendering to the client device.' }
        @{ Name = 'USB'; Label = 'USB (USB redirection)'; Tip = 'USB device pass-through to virtual sessions.' }
        @{ Name = 'SSON'; Label = 'SSON (single sign-on component)'; Tip = 'The single sign-on credential provider.' }
        @{ Name = 'CitrixEnterpriseBrowser'; Label = 'CitrixEnterpriseBrowser (LTSR only)'; Tip = 'The Citrix Enterprise Browser; only the LTSR packagers write this key.' }
    )
    $sections.Add(@{
            Id = 'cwa'; Title = 'Citrix Workspace App'; Parent = 'applications'; Kind = 'Fields'; Keywords = 'citrix store installation plugins update telemetry store policy components addlocal'
            Description = 'The switches of the Current Release and LTSR packagers, in six groups: store, installation, plug-ins and add-ons, update and telemetry, store policy and components.'
            Post = { param($Context, $Page) Initialize-OptionsCitrixSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'cwa.preview'; Title = 'Command-line preview'; Kind = 'Actions'; Tip = 'Shows the install command line that the Current Release and LTSR packagers build.'; Keywords = 'cwa citrix arguments'; Buttons = @(@{ Id = 'preview'; Label = 'Preview'; Tip = 'Shows the command line for both streams.' }) }
                @{ Id = 'cwa.storeGroup'; Kind = 'Group'; Title = 'Store' }
                @{ Id = 'cwa.storeName'; Title = 'Store name'; Kind = 'Text'; Pref = 'cwa:Store.Name'; Tip = 'The friendly name of the StoreFront store (STORE0).' }
                @{ Id = 'cwa.storeUrl'; Title = 'Store URL'; Kind = 'Text'; Pref = 'cwa:Store.Url'; Tip = 'STORE0 URL: StoreFront https://host/Citrix/Store/discovery, Gateway https://host#StoreName.'; Keywords = 'storefront gateway' }
                @{ Id = 'cwa.installGroup'; Kind = 'Group'; Title = 'Installation' }
                @{ Id = 'cwa.clean'; Title = 'Clean install'; Kind = 'Check'; Pref = 'cwa:Installation.CleanInstall'; Label = 'Remove leftover configuration before install (/CleanInstall)'; Tip = 'Removes configuration and registry data of an earlier installation before the install.' }
                @{ Id = 'cwa.sso'; Title = 'Single sign-on'; Kind = 'Check'; Pref = 'cwa:Installation.IncludeSSON'; Label = 'Install the SSO component (/includeSSON, ENABLE_SSON=Yes)'; Tip = 'Installs the SSO component and turns on domain pass-through authentication.' }
                @{ Id = 'cwa.appprot'; Title = 'App Protection'; Kind = 'Check'; Pref = 'cwa:Installation.AppProtection'; Label = 'Start the App Protection component (startAppProtection)'; Tip = 'Starts the component; the installer already installs it, and it needs an admin install.' }
                @{ Id = 'cwa.prelaunch'; Title = 'Session pre-launch'; Kind = 'Check'; Pref = 'cwa:Installation.SessionPreLaunch'; Label = 'Pre-launch a session at logon (ENABLEPRELAUNCH)'; Tip = 'Starts a Citrix session at logon so that applications start faster.' }
                @{ Id = 'cwa.selfservice'; Title = 'Self-service mode'; Kind = 'Check'; Pref = 'cwa:Installation.SelfServiceMode'; Label = 'Show the self-service window (SELFSERVICEMODE)'; Tip = 'Shows the Citrix Workspace self-service window; cleared hides it.' }
                @{ Id = 'cwa.forceinstall'; Title = 'Force install'; Kind = 'Check'; Pref = 'cwa:Installation.ForceInstall'; Label = 'Clean old Citrix data after a failed install (/forceinstall)'; Tip = 'Cleans old data when an install or upgrade fails; use it for a retry package.' }
                @{ Id = 'cwa.featuresGroup'; Kind = 'Group'; Title = 'Features and client name' }
                @{ Id = 'cwa.bidir'; Title = 'Content redirection'; Kind = 'Check'; Pref = 'cwa:Installation.BidirContentRedirection'; Label = 'Allow redirection in both directions (ALLOW_BIDIRCONTENTREDIRECTION=1)'; Tip = 'Allows content redirection between the client and the host in both directions.' }
                @{ Id = 'cwa.hostedurl'; Title = 'URL redirection'; Kind = 'Check'; Pref = 'cwa:Installation.ClientHostedAppsUrl'; Label = 'Turn on URL redirection (ALLOW_CLIENTHOSTEDAPPSURL=1)'; Tip = 'Turns on URL redirection on the device.' }
                @{ Id = 'cwa.laa'; Title = 'Local App Access'; Kind = 'Check'; Pref = 'cwa:Installation.ForceLaa'; Label = 'Install Local App Access (FORCE_LAA=1)'; Tip = 'Installs Local App Access; it needs an admin install.' }
                @{ Id = 'cwa.legacyicons'; Title = 'File-type icons'; Kind = 'Check'; Pref = 'cwa:Installation.LegacyFtaIcons'; Label = 'Hide file-type icons of subscribed apps (LEGACYFTAICONS=True)'; Tip = 'Hides the file-type icons of subscribed apps.' }
                @{ Id = 'cwa.disablesetting'; Title = 'Advanced Preferences'; Kind = 'Combo'; Pref = 'cwa:Installation.DisableSetting'; Default = ''; Width = 300; Items = @($notSet, @{ Label = '0 - show Shortcuts and Reconnect'; Value = '0' }, @{ Label = '1 - show Reconnect only'; Value = '1' }, @{ Label = '2 - show Shortcuts only'; Value = '2' }, @{ Label = '3 - hide both'; Value = '3' }); Tip = 'Hides the Shortcuts and Reconnect options in Advanced Preferences (DisableSetting).' }
                @{ Id = 'cwa.clientname'; Title = 'Client name'; Kind = 'Text'; Pref = 'cwa:Installation.ClientName'; MaxLength = 20; Tip = 'A fixed client device name; empty uses the computer name.'; Keywords = 'client_name dynamic' }
                @{ Id = 'cwa.pluginsGroup'; Kind = 'Group'; Title = 'Plugins and add-ons' }
                @{ Id = 'cwa.teams'; Title = 'MS Teams VDI plugin'; Kind = 'Check'; Pref = 'cwa:Plugins.MSTeamsPlugin'; Label = 'Install the plug-in (/InstallMSTeamsPlugin)'; Tip = 'Current Release only; the LTSR packagers ignore this setting.' }
                @{ Id = 'cwa.zoom'; Title = 'Zoom VDI plugin'; Kind = 'Check'; Pref = 'cwa:Plugins.ZoomPlugin'; Label = 'Install the plug-in (Installzoomplugin, ADDONS=ZoomVDIPlugin)'; Tip = 'Current Release adds Installzoomplugin=N when cleared; LTSR adds ADDONS=ZoomVDIPlugin when set.' }
                @{ Id = 'cwa.webex'; Title = 'WebEx VDI plugin'; Kind = 'Check'; Pref = 'cwa:Plugins.WebExPlugin'; Label = 'Install the plug-in (ADDONS=WebexVDIPlugin)'; Tip = 'Installs the WebEx VDI plugin engine.' }
                @{ Id = 'cwa.uber'; Title = 'uberAgent'; Kind = 'Check'; Pref = 'cwa:Plugins.UberAgent'; Label = 'Install or upgrade uberAgent (/InstallUberAgent)'; Tip = 'Installs or upgrades the uberAgent monitoring plug-in.' }
                @{ Id = 'cwa.uberskip'; Title = 'uberAgent upgrade'; Kind = 'Check'; Pref = 'cwa:Plugins.UberAgentSkipUpgrade'; Label = 'Skip the upgrade when uberAgent is present (/SkipUberAgentUpgrade)'; Tip = 'Installs uberAgent only when it is not on the device yet.' }
                @{ Id = 'cwa.epa'; Title = 'EPA client'; Kind = 'Check'; Pref = 'cwa:Plugins.EPAClient'; Label = 'Install the EPA client (cleared adds InstallEPAClient=N)'; Tip = 'The Endpoint Analysis client for device posture checks.' }
                @{ Id = 'cwa.sr'; Title = 'Session Recording agent'; Kind = 'Check'; Pref = 'cwa:Plugins.SessionRecording'; Label = 'Install the agent (/InstallSRAgent)'; Tip = 'Current Release only; the LTSR packagers ignore this setting.' }
                @{ Id = 'cwa.srskip'; Title = 'Session Recording upgrade'; Kind = 'Check'; Pref = 'cwa:Plugins.SkipSRAgentUpgrade'; Label = 'Keep the agent that is on the device (/SkipSRAgentUpgrade)'; Tip = 'Keeps an installed Session Recording agent; Current Release only.' }
                @{ Id = 'cwa.updateGroup'; Kind = 'Group'; Title = 'Update and telemetry' }
                @{ Id = 'cwa.autoupdate'; Title = 'Auto-update'; Kind = 'Combo'; Pref = 'cwa:UpdateAndTelemetry.AutoUpdateCheck'; Default = ''; Width = 200; Items = @($notSet, @{ Label = 'auto'; Value = 'auto' }, @{ Label = 'manual'; Value = 'manual' }, @{ Label = 'disabled'; Value = 'disabled' }); Tip = 'AutoUpdateCheck; (not set) omits the switch and the installer default applies.' }
                @{ Id = 'cwa.stream'; Title = 'Update stream'; Kind = 'Combo'; Pref = 'cwa:UpdateAndTelemetry.AutoUpdateStream'; Default = ''; Width = 200; Items = @($notSet, @{ Label = 'LTSR'; Value = 'LTSR' }, @{ Label = 'Current'; Value = 'Current' }); Tip = 'AutoUpdateStream; it works only when Auto-update is set.' }
                @{ Id = 'cwa.rollout'; Title = 'Rollout priority'; Kind = 'Combo'; Pref = 'cwa:UpdateAndTelemetry.AURolloutPriority'; Default = ''; Width = 200; Items = @($notSet, @{ Label = 'Auto'; Value = 'Auto' }, @{ Label = 'Fast'; Value = 'Fast' }, @{ Label = 'Medium'; Value = 'Medium' }, @{ Label = 'Slow'; Value = 'Slow' }); Tip = 'AURolloutPriority; it works only when Auto-update is set.' }
                @{ Id = 'cwa.ceip'; Title = 'CEIP and telemetry'; Kind = 'Check'; Pref = 'cwa:UpdateAndTelemetry.EnableCEIP'; Label = 'Join the Customer Experience Improvement Program (EnableCEIP)'; Tip = 'Sends Citrix usage data when selected.' }
                @{ Id = 'cwa.tracing'; Title = 'Always-on tracing'; Kind = 'Check'; Pref = 'cwa:UpdateAndTelemetry.EnableTracing'; Label = 'Turn on always-on diagnostic tracing (EnableTracing)'; Tip = 'Writes diagnostic traces on the device.' }
                @{ Id = 'cwa.policyGroup'; Kind = 'Group'; Title = 'Store policy' }
                @{ Id = 'cwa.addstore'; Title = 'Allow add store'; Kind = 'Combo'; Pref = 'cwa:StorePolicy.AllowAddStore'; Default = ''; Width = 280; Items = $cwaPolicyItems; Tip = 'ALLOWADDSTORE; (not set) omits the switch and the installer default S applies.' }
                @{ Id = 'cwa.savepwd'; Title = 'Allow save password'; Kind = 'Combo'; Pref = 'cwa:StorePolicy.AllowSavePwd'; Default = ''; Width = 280; Items = $cwaPolicyItems; Tip = 'ALLOWSAVEPWD; (not set) omits the switch and the installer default S applies.' }
                @{ Id = 'cwa.componentsGroup'; Kind = 'Group'; Title = 'Components ADDLOCAL'; Description = 'With Customize cleared, the installer uses its default components and the packager omits ADDLOCAL.' }
                @{ Id = 'cwa.customize'; Title = 'Customize'; Kind = 'Check'; Pref = 'cwa:Components.Customize'; Label = 'Write ADDLOCAL with the components below'; Tip = 'When cleared, the packager omits ADDLOCAL and the installer default applies.' }
            ) +
            @($cwaComponents | ForEach-Object { @{ Id = 'cwa.comp.' + $_.Name; Title = $_.Name; Kind = 'Check'; Pref = 'cwa:Components.' + $_.Name; Label = $_.Label; Tip = $_.Tip; Required = [bool]$_.Required; Indent = 0 } })
        })

    $sections.Add(@{
            Id = 'dbeaver'; Title = 'DBeaver Community'; Parent = 'applications'; Kind = 'Fields'
            Fields = @(
                @{ Id = 'dbv.scope'; Title = 'Install scope'; Kind = 'Combo'; Pref = 'DBeaverInstallOptions.InstallScope'; Default = 'System'; Width = 420; Items = @(@{ Label = 'System - all users, in Program Files (/allusers)'; Value = 'System' }, @{ Label = 'User - current user, no elevation (/currentuser)'; Value = 'User' }); Tip = 'User deploys per user, so deploy it in user context for the detection rule to match.' }
                @{ Id = 'dbv.ai'; Title = 'AI features'; Kind = 'Check'; Pref = 'DBeaverInstallOptions.DisableAI'; Label = 'Set DBEAVER_AI_DISABLED=true after install'; Tip = 'Sets the variable at machine scope for System and at user scope for User.'; Keywords = 'disable ai assistant' }
            )
        })

    $sections.Add(@{
            Id = 'm365'; Title = 'Microsoft 365 Apps'; Parent = 'applications'; Kind = 'Fields'; Prefs = @('M365CustomConfig')
            Description = 'The Office Deployment Tool settings of the Microsoft 365 Apps, Project and Visio packagers.'
            Post = { param($Context, $Page) Initialize-OptionsM365Section -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'm365.links'; Title = 'Documentation'; Kind = 'Custom'; Tip = 'Opens the Office Deployment Tool download and the configuration.xml reference.'; Keywords = 'odt download office deployment tool'
                    Build = { param($Field, $Context) New-OptionsLinkEditor -Link @(@{ Label = 'Download the Office Deployment Tool'; Url = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' }, @{ Label = 'configuration.xml options (Microsoft Learn)'; Url = 'https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options' }) } }
                @{ Id = 'm365.customGroup'; Kind = 'Group'; Title = 'Your own configuration.xml' }
                @{ Id = 'm365.useCustom'; Title = 'Use my own configuration.xml'; Kind = 'Check'; Pref = 'M365CustomConfig.Enabled'; Label = 'Use the imported file for every Microsoft 365 package'; Tip = 'The packager uses the imported file and keeps its own version, source and detection.'; Keywords = 'import custom' }
                @{ Id = 'm365.file'; Title = 'Imported file'; Kind = 'Status'; Tip = 'The name and size of the stored configuration.xml; the file stays on this computer.'; Keywords = 'import remove'; Buttons = @(@{ Id = 'import'; Label = 'Import...'; Tip = 'Checks a configuration.xml and stores a copy.' }, @{ Id = 'remove'; Label = 'Remove'; Tip = 'Deletes the stored copy.' }) }
                @{ Id = 'm365.reason'; Kind = 'Note'; Text = '' }
                @{ Id = 'm365.generatedGroup'; Kind = 'Group'; Title = 'Generated configuration' }
                @{ Id = 'm365.company'; Title = 'Company name'; Kind = 'Text'; Pref = 'CompanyName'; MaxLength = 100; Tip = 'The organization name that the packagers write into the Office settings.' }
                @{ Id = 'm365.channel'; Title = 'Channel'; Kind = 'Combo'; Pref = 'M365Channel'; Default = 'MonthlyEnterprise'; Width = 260; Items = $channelItems; Tip = 'The update channel of the Microsoft 365 Apps, Project and Visio packagers.' }
                @{ Id = 'm365.deploy'; Title = 'Deploy mode'; Kind = 'Combo'; Pref = 'M365DeployMode'; Default = 'Managed'; Width = 260; Items = $deployItems; Tip = 'Managed downloads the Office source and pins the version; Online installs from the CDN.' }
                @{ Id = 'm365.languages'; Title = 'Languages'; Kind = 'Text'; Pref = 'M365Options.Languages'; Watermark = 'en-us'
                    ToText = { param($Value) (@($Value) -join ', ') }
                    FromText = { param($Text) @(([string]$Text) -split '[,;\s]+' | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() }) }
                    Tip = 'Culture codes of the Office display languages, such as en-us; the first sets the shell language.' }
                @{ Id = 'm365.exclude'; Title = 'Excluded apps'; Kind = 'Multi'; Pref = 'M365ExcludeApps'; Columns = 2; Tip = 'The ExcludeApp IDs that the packagers write into the Office configuration.'
                    Items = @(
                        @{ Id = 'Access'; Label = 'Access'; Tip = 'Excludes Microsoft Access.' }
                        @{ Id = 'Excel'; Label = 'Excel'; Tip = 'Excludes Excel.' }
                        @{ Id = 'Groove'; Label = 'OneDrive for Business (Groove)'; Tip = 'Excludes the OneDrive sync client of Office; the ExcludeApp ID is Groove.' }
                        @{ Id = 'Lync'; Label = 'Skype for Business (Lync)'; Tip = 'Excludes Skype for Business; the ExcludeApp ID is Lync.' }
                        @{ Id = 'OneDrive'; Label = 'OneDrive'; Tip = 'Excludes the OneDrive client that Office installs.' }
                        @{ Id = 'OneNote'; Label = 'OneNote'; Tip = 'Excludes OneNote.' }
                        @{ Id = 'Outlook'; Label = 'Outlook (classic)'; Tip = 'Excludes classic Outlook.' }
                        @{ Id = 'OutlookForWindows'; Label = 'Outlook for Windows (new)'; Tip = 'Excludes the new Outlook for Windows app.' }
                        @{ Id = 'PowerPoint'; Label = 'PowerPoint'; Tip = 'Excludes PowerPoint.' }
                        @{ Id = 'Publisher'; Label = 'Publisher'; Tip = 'Excludes Publisher.' }
                        @{ Id = 'Teams'; Label = 'Teams'; Tip = 'Excludes the Teams installer that Office bundles.' }
                        @{ Id = 'Word'; Label = 'Word'; Tip = 'Excludes Word.' }) }
                @{ Id = 'm365.behaviorGroup'; Kind = 'Group'; Title = 'Install behavior' }
                @{ Id = 'm365.officeMgmt'; Title = 'Update source'; Kind = 'Check'; Pref = 'M365Options.OfficeMgmtCom'; Label = 'Let ConfigMgr software updates patch Office (OfficeMgmtCOM)'; Tip = 'Selected lets ConfigMgr deliver the updates; cleared lets Office update from the CDN.' }
                @{ Id = 'm365.updates'; Title = 'Office updates'; Kind = 'Check'; Pref = 'M365Options.UpdatesEnabled'; Label = 'Let Office check for updates (Updates Enabled)'; Tip = 'Selected lets Office check for updates; ConfigMgr stays the source with OfficeMgmtCOM.' }
                @{ Id = 'm365.forceShutdown'; Title = 'Force app shutdown'; Kind = 'Check'; Pref = 'M365Options.ForceAppShutdown'; Label = 'Close Office apps that block setup (FORCEAPPSHUTDOWN)'; Tip = 'Closes Office apps that block setup; users can lose unsaved data.' }
                @{ Id = 'm365.removeMsi'; Title = 'Remove MSI Office'; Kind = 'Check'; Pref = 'M365Options.RemoveMsi'; Label = 'Remove MSI versions of Office before the install (RemoveMSI)'; Tip = 'Removes Office 2007 to 2016 MSI installations; it does not remove Click-to-Run Office.' }
                @{ Id = 'm365.migrate'; Title = 'Migrate architecture'; Kind = 'Check'; Pref = 'M365Options.MigrateArch'; Label = 'Change an installed Office to the architecture of this package (MigrateArch)'; Tip = 'Use it for 32-bit to 64-bit upgrades only; an x86 package would downgrade 64-bit Office.' }
                @{ Id = 'm365.shared'; Title = 'Shared computer activation'; Kind = 'Check'; Pref = 'M365Options.SharedComputerLicensing'; Label = 'Turn on for RDS and VDI hosts (SharedComputerLicensing)'; Tip = 'Leave it off for normal computers; user-based licensing is the default.' }
                @{ Id = 'm365.device'; Title = 'Device-based licensing'; Kind = 'Check'; Pref = 'M365Options.DeviceBasedLicensing'; Label = 'License Office to the device (DeviceBasedLicensing)'; Tip = 'For shared devices; the device must be synced to Microsoft Entra ID.' }
                @{ Id = 'm365.previewGroup'; Kind = 'Group'; Title = 'Preview' }
                @{ Id = 'm365.preview'; Title = 'Install XML preview'; Kind = 'Actions'; Tip = 'Shows the install XML that the packagers build from the settings of this section.'; Buttons = @(@{ Id = 'preview'; Label = 'Preview'; Tip = 'Shows the install XML of every Microsoft 365 package.' }) }
            )
        })

    $sections.Add(@{
            Id = 'ssms'; Title = 'SQL Server Management Studio'; Parent = 'applications'; Kind = 'Fields'
            Description = 'The command line of the SSMS 22 setup. Recommended and optional components apply to the selected workloads.'
            Post = { param($Context, $Page) Initialize-OptionsSsmsSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'ssms.ui'; Title = 'UI mode'; Kind = 'Combo'; Pref = 'SSMSInstallOptions.UIMode'; Default = 'Quiet'; Width = 200; Items = @('Quiet', 'Passive'); Tip = 'Quiet adds --quiet; Passive adds --passive and shows progress, which a required install cannot use.' }
                @{ Id = 'ssms.path'; Title = 'Install path'; Kind = 'Text'; Pref = 'SSMSInstallOptions.InstallPath'; Tip = 'The --installPath value; empty uses the default folder, and detection and uninstall follow it.' }
                @{ Id = 'ssms.dti'; Title = 'Download first'; Kind = 'Check'; Pref = 'SSMSInstallOptions.DownloadThenInstall'; Label = 'Download all packages before the install (--downloadThenInstall)'; Tip = 'Excludes --installWhileDownloading, which is the setup default.' }
                @{ Id = 'ssms.noupdate'; Title = 'Installer update'; Kind = 'Check'; Pref = 'SSMSInstallOptions.NoUpdateInstaller'; Label = 'Do not update the Visual Studio Installer (--noUpdateInstaller)'; Tip = 'Documented with --quiet; setup fails when an installer update is required.' }
                @{ Id = 'ssms.workloads'; Title = 'Workloads'; Kind = 'Multi'; Pref = 'SSMSInstallOptions.Workloads'; Columns = 2; Tip = 'The workloads to install; each one adds an --add switch, and none installs the core editor.'
                    Items = @(
                        @{ Id = 'Microsoft.SqlServer.Workload.SSMS.AI'; Label = 'AI assistance'; Tip = 'Adds the workload Microsoft.SqlServer.Workload.SSMS.AI.' }
                        @{ Id = 'Microsoft.SqlServer.Workload.SSMS.BI'; Label = 'Business intelligence'; Tip = 'Adds the workload Microsoft.SqlServer.Workload.SSMS.BI.' }
                        @{ Id = 'Microsoft.SqlServer.Workload.SSMS.CodeTools'; Label = 'Code tools'; Tip = 'Adds the workload Microsoft.SqlServer.Workload.SSMS.CodeTools.' }
                        @{ Id = 'Microsoft.SqlServer.Workload.SSMS.DevOps'; Label = 'DevOps'; Tip = 'Adds the workload Microsoft.SqlServer.Workload.SSMS.DevOps.' }
                        @{ Id = 'Microsoft.SqlServer.Workload.SSMS.HybridAndMigration'; Label = 'Hybrid and migration'; Tip = 'Adds the workload Microsoft.SqlServer.Workload.SSMS.HybridAndMigration.' }) }
                @{ Id = 'ssms.recommended'; Title = 'Recommended components'; Kind = 'Check'; Pref = 'SSMSInstallOptions.IncludeRecommended'; Label = 'Include recommended components (--includeRecommended)'; Tip = 'Needs at least one workload; without a workload the switch has no effect.' }
                @{ Id = 'ssms.optional'; Title = 'Optional components'; Kind = 'Check'; Pref = 'SSMSInstallOptions.IncludeOptional'; Label = 'Include optional components (--includeOptional)'; Tip = 'Needs at least one workload; it increases the install size and duration.' }
                @{ Id = 'ssms.languages'; Title = 'Languages'; Kind = 'Multi'; Pref = 'SSMSInstallOptions.Languages'; Columns = 7; Tip = 'The SSMS display languages (--addProductLang); none uses the language of the operating system.'
                    Items = @('cs-cz', 'de-de', 'en-us', 'es-es', 'fr-fr', 'it-it', 'ja-jp', 'ko-kr', 'pl-pl', 'pt-br', 'ru-ru', 'tr-tr', 'zh-cn', 'zh-tw' | ForEach-Object { @{ Id = $_; Label = $_ } }) }
                @{ Id = 'ssms.nocache'; Title = 'Package cache'; Kind = 'Check'; Pref = 'SSMSInstallOptions.NoCache'; Label = 'Delete the setup packages after the install (--nocache)'; Tip = 'Saves disk space; a repair downloads the packages again.' }
                @{ Id = 'ssms.oos'; Title = 'Out-of-support components'; Kind = 'Check'; Pref = 'SSMSInstallOptions.RemoveOos'; Label = 'Remove out-of-support components (--removeOos true)'; Tip = 'Removes components that are no longer supported during this install or update.' }
                @{ Id = 'ssms.force'; Title = 'Running SSMS'; Kind = 'Check'; Pref = 'SSMSInstallOptions.ForceClose'; Label = 'Close SSMS when it is in use (--force)'; Tip = 'Closes running SSMS processes; users can lose unsaved query windows.' }
            )
        })

    $sections.Add(@{
            Id = 'teamviewer'; Title = 'TeamViewer Host'; Parent = 'applications'; Kind = 'Fields'
            Description = 'The packager installs the TeamViewer Host EXE with /S, and the EXE ignores the four settings below. They apply to a TeamViewer Host MSI only; the packager logs a warning when they are set.'
            Fields = @(
                @{ Id = 'tv.token'; Title = 'API token'; Kind = 'Text'; Pref = 'tv:ApiToken'; Tip = 'MSI only: the script token for APITOKEN; it is stored as text in teamviewer-host-config.json.'; Keywords = 'apitoken' }
                @{ Id = 'tv.config'; Title = 'Custom config ID'; Kind = 'Text'; Pref = 'tv:CustomConfigId'; Tip = 'MSI only: CUSTOMCONFIGID of a custom Host module from the Management Console.' }
                @{ Id = 'tv.assign'; Title = 'Assignment options'; Kind = 'Text'; Pref = 'tv:AssignmentOptions'; Tip = 'MSI only: the flags of ASSIGNMENTOPTIONS, for example --grant-easy-access or --reassign.' }
                @{ Id = 'tv.shortcut'; Title = 'Desktop shortcut'; Kind = 'Check'; Pref = 'tv:RemoveDesktopShortcut'; Label = 'Remove the desktop shortcut (MSI only)'; Tip = 'MSI only: ADDLOCAL=ALL REMOVE=f.DesktopShortcut removes the desktop shortcut.' }
            )
        })

    $sections.Add(@{
            Id = 'localsources'; Title = 'Local Installer Sources'; Parent = 'applications'; Kind = 'Fields'; Prefs = @('LocalSourceFolders')
            Description = 'A packager whose download needs a sign-in stages the newest matching installer from its folder. Stage asks for the folder when it is empty.'
            Fields = @(@{ Id = 'local.rows'; Title = 'Folders'; Kind = 'Custom'; Tip = 'The installer folder of each packager that can stage a local installer.'; Keywords = 'local source installer folder packager file'
                    Build = { param($Field, $Context) New-OptionsLocalSourceEditor -Context $Context } })
        })

    # ---------------------------------------------------------------- Updates and CVEs
    $sections.Add(@{ Id = 'updates'; Title = 'Updates and CVEs'; Kind = 'Group' })
    $sections.Add(@{
            Id = 'oneclick'; Title = 'One Click'; Parent = 'updates'; Kind = 'Panel'; Factory = { New-AppFlowPanel }
            Prefs = @('AppFlow', 'Intune.DeploymentTarget', 'Intune.PublishToIntune')
            SearchFields = @(
                @{ Id = 'oc.action'; Title = 'Action on update'; Control = 'cboAction' }
                @{ Id = 'oc.force'; Title = 'Force on launch'; Control = 'toggleForce' }
                @{ Id = 'oc.destinations'; Title = 'Default destinations'; Control = 'chkOneClickMecm'; Keywords = 'configmgr wsus intune' }
                @{ Id = 'oc.existing'; Title = 'Existing ConfigMgr app'; Control = 'cboOneClickOnExisting' }
                @{ Id = 'oc.track'; Title = 'Tracked applications'; Control = 'dgApps'; Keywords = 'track cadence days' })
        })
    $sections.Add(@{
            Id = 'productfilter'; Title = 'Product Filter'; Parent = 'updates'; Kind = 'Panel'; Factory = { New-ProductFilterPanel }
            Prefs = @('HiddenApplications', 'DiscoveryDismissed')
            SearchFields = @(
                @{ Id = 'pf.hidden'; Title = 'Hidden applications'; Control = 'treeApps'; Keywords = 'vendor select all none show hide' }
                @{ Id = 'pf.dismissed'; Title = 'Not offered by Discover'; Control = 'btnClearDiscoveryDismissed'; Keywords = 'discover unmanaged dismissed clear list' })
        })
    $sections.Add(@{
            Id = 'cve'; Title = 'CVE Lookup'; Parent = 'updates'; Kind = 'Fields'
            Description = 'Check CVEs reads the National Vulnerability Database. NVD allows 5 requests per 30 seconds without a key and 50 with one.'
            Fields = @(
                @{ Id = 'cve.key'; Title = 'NVD API key'; Kind = 'Secret'; Pref = 'Vulnerability.NvdApiKeyProtected'; Tip = 'Stored with DPAPI for the current Windows user; empty keeps the saved key.'; Keywords = 'nvd' }
                @{ Id = 'cve.ttl'; Title = 'Cache time'; Kind = 'Number'; Pref = 'Vulnerability.CacheTtlMinutes'; Min = 0; Max = 525600; Default = 360; MaxLength = 6; Unit = 'minutes'; Tip = 'How long a CVE lookup result stays valid before the next lookup; 0 turns the cache off.'; Keywords = 'ttl' }
                @{ Id = 'cve.after'; Title = 'Automatic lookup'; Kind = 'Check'; Pref = 'Vulnerability.LookupAfterCheckCm'; Label = 'Look up CVEs after Check ConfigMgr'; Tip = 'Runs Check CVEs when Check ConfigMgr finishes.' }
            )
        })

    # ---------------------------------------------------------------- Security
    $sections.Add(@{ Id = 'security'; Title = 'Security'; Kind = 'Group' })
    $sections.Add(@{
            Id = 'signing'; Title = 'Script Signing'; Parent = 'security'; Kind = 'Panel'; Factory = { New-ScriptSigningPanel }
            Prefs = @('ScriptSigning')
            SearchFields = @(
                @{ Id = 'sign.detection'; Title = 'Sign detection scripts'; Control = 'chkSignDetection' }
                @{ Id = 'sign.requirements'; Title = 'Sign requirement scripts'; Control = 'chkSignRequirements' }
                @{ Id = 'sign.deployment'; Title = 'Sign install and uninstall scripts'; Control = 'chkSignDeployment' }
                @{ Id = 'sign.reqdetection'; Title = 'Require valid detection signatures'; Control = 'chkRequireDetection' }
                @{ Id = 'sign.reqrequirements'; Title = 'Require valid requirement signatures'; Control = 'chkRequireRequirements' }
                @{ Id = 'sign.reqdeployment'; Title = 'Require valid deployment signatures'; Control = 'chkRequireDeployment' }
                @{ Id = 'sign.store'; Title = 'Certificate store'; Control = 'cboSignStore' }
                @{ Id = 'sign.certificate'; Title = 'Certificate'; Control = 'cboSignCertificate'; Keywords = 'thumbprint' }
                @{ Id = 'sign.timestamp'; Title = 'Timestamp server'; Control = 'txtSignTimestamp' }
                @{ Id = 'sign.timestampreq'; Title = 'Timestamp required'; Control = 'chkSignTimestampRequired' })
        })

    # ---------------------------------------------------------------- Tools
    $sections.Add(@{
            Id = 'tools'; Title = 'Tools'; Kind = 'Fields'; Prefs = @('DetectedTools')
            Description = 'The tools that AppPackager detects when it starts.'
            Post = { param($Context, $Page) Initialize-OptionsToolsSection -Context $Context -Page $Page }
            Fields = @(
                @{ Id = 'tools.console'; Title = 'Console'; Kind = 'Status'; Tip = 'Shows whether the ConfigMgr console is installed on this computer.'; Keywords = 'configmgr adminui' }
                @{ Id = 'tools.contentprep'; Title = 'Content Prep Tool'; Kind = 'Status'; Tip = 'Shows whether IntuneWinAppUtil.exe is found.'; Keywords = 'intunewin intunewinapputil'; Buttons = @(@{ Id = 'download'; Label = 'Download'; Tip = 'Downloads the Microsoft Win32 Content Prep Tool.' }) }
                @{ Id = 'tools.sevenzip'; Title = '7-Zip'; Kind = 'Status'; Tip = 'Shows whether 7z.exe is found; the Adobe Reader packager needs it.'; Keywords = '7z' }
                @{ Id = 'tools.iconpack'; Title = 'Icon pack'; Kind = 'Status'; Tip = 'Shows the state of the packager icon pack.'; Keywords = 'icons'; Buttons = @(@{ Id = 'download'; Label = 'Download'; Tip = 'Downloads the icon pack and extracts it into Packagers\Icons.' }, @{ Id = 'file'; Label = 'Install from file...'; Tip = 'Installs the icon pack from a local or UNC zip file.' }) }
                @{ Id = 'gh.group'; Kind = 'Group'; Title = 'GitHub Authentication'; Description = 'The GitHub-backed packagers and the version monitor send a token as a bearer token. A token raises the limit from 60 to 5000 requests per hour.' }
                @{ Id = 'gh.status'; Title = 'Status'; Kind = 'Status'; Tip = 'Shows how AppPackager signs in to the GitHub API and the requests that are left this hour.'; Keywords = 'github authentication token rate limit gh' }
                @{ Id = 'gh.token'; Title = 'GitHub token'; Kind = 'Secret'; Pref = 'GitHub.TokenProtected'; Tip = 'Optional, DPAPI-protected; 5000 requests per hour with a token, 60 without.'; Keywords = 'github key personal access token pat' }
            )
        })

    $sections.Add(@{ Id = 'about'; Title = 'About'; Kind = 'Panel'; Factory = { New-AboutPanel }; SearchFields = @(@{ Id = 'about.update'; Title = 'Update now'; Control = 'btnAboutUpdate'; Keywords = 'version release' }) })

    return $sections.ToArray()
}

# =============================================================================
# Options window: section behavior
# =============================================================================
function Get-OptionsLiveEditor {
    # The editor of a field in a section that is built, or $null.
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$FieldId)
    foreach ($page in @($Context.Sections.Values)) {
        if ($page -and $page.Editors -and $page.Editors.Contains($FieldId)) { return $page.Editors[$FieldId] }
    }
    return $null
}

function Get-OptionsLiveText {
    # The text a field holds now: from its editor when the section is built, else from the stored preference.
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$FieldId)
    $editor = Get-OptionsLiveEditor -Context $Context -FieldId $FieldId
    if ($editor -and $editor.Get) { return ([string](& $editor.Get)).Trim() }
    if ($null -eq $Context.FieldTable) { $Context.FieldTable = Get-OptionsFieldTable }
    $field = $Context.FieldTable[$FieldId]
    if ($field -and $field.Pref) { return ([string](Get-OptionsPrefValue -Context $Context -Path ([string]$field.Pref))).Trim() }
    return ''
}

function Get-OptionsLiveFlag {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$FieldId)
    $editor = Get-OptionsLiveEditor -Context $Context -FieldId $FieldId
    if ($editor -and $editor.Get) { return ((& $editor.Get) -eq $true) }
    if ($null -eq $Context.FieldTable) { $Context.FieldTable = Get-OptionsFieldTable }
    $field = $Context.FieldTable[$FieldId]
    if ($field -and $field.Pref) { return ((Get-OptionsPrefValue -Context $Context -Path ([string]$field.Pref)) -eq $true) }
    return $false
}

function Get-OptionsLiveSecret {
    # The protected secret to use: a typed value, else the stored one. Returns plain text, or '' when there is none.
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$FieldId)
    $editor = Get-OptionsLiveEditor -Context $Context -FieldId $FieldId
    if ($editor -and $editor.Get) {
        $typed = & $editor.Get
        if ($typed -and $typed.Length -gt 0) {
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($typed)
            try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
            finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        }
    }
    if ($null -eq $Context.FieldTable) { $Context.FieldTable = Get-OptionsFieldTable }
    $stored = [string](Get-OptionsPrefValue -Context $Context -Path ([string]$Context.FieldTable[$FieldId].Pref))
    if ([string]::IsNullOrWhiteSpace($stored)) { return '' }
    $secure = ConvertTo-SecureString -String $stored -ErrorAction Stop
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Get-OptionsToolStatusText {
    param([Parameter(Mandatory)]$Prefs, [Parameter(Mandatory)][ValidateSet('Console', 'SevenZip', 'ContentPrep')][string]$Tool)
    $check = [string][char]0x2713
    $cross = [string][char]0x2717
    switch ($Tool) {
        'Console' {
            $cm = $Prefs.DetectedTools.ConfigMgrConsole
            if ($cm -and $cm.Found) { return @{ Text = ('{0} Detected  -  {1} v{2}' -f $check, $cm.DisplayName, $cm.DisplayVersion); Tip = ('Module: {0}' -f $cm.ModulePath) } }
            return @{ Text = ($cross + ' Not detected  -  install the ConfigMgr console and restart AppPackager'); Tip = 'Found at start through the uninstall key, SMS_ADMIN_UI_PATH and the usual install folders.' }
        }
        'SevenZip' {
            $sz = $Prefs.DetectedTools.SevenZipCli
            if ($sz -and $sz.Found) { return @{ Text = ('{0} Detected  -  {1} v{2}' -f $check, $sz.DisplayName, $sz.DisplayVersion); Tip = ('7z.exe: {0}' -f $sz.ExePath) } }
            return @{ Text = ($cross + ' Not detected  -  the Adobe Reader packager needs 7-Zip'); Tip = 'Detected at start through the uninstall registry key and the 7-Zip program folder.' }
        }
        'ContentPrep' {
            $iw = $Prefs.DetectedTools.IntuneWinAppUtil
            if ($iw -and $iw.Found) {
                $version = $(if ([string]::IsNullOrWhiteSpace([string]$iw.DisplayVersion)) { '' } else { ' v' + $iw.DisplayVersion })
                return @{ Text = ('{0} Detected  -  IntuneWinAppUtil{1}' -f $check, $version); Tip = ('IntuneWinAppUtil.exe: {0}' -f $iw.ExePath) }
            }
            return @{ Text = ($cross + ' Not detected  -  download it here or put IntuneWinAppUtil.exe on the PATH'); Tip = 'Checked once at start: preference path, LOCALAPPDATA tool cache, PATH.' }
        }
    }
}

function Initialize-OptionsConfigMgrSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $status = Get-OptionsToolStatusText -Prefs $Context.Prefs -Tool Console
    & $Page.Editors['cm.console'].Set $status.Text
    $Page.Editors['cm.console'].Text.ToolTip = $status.Tip
}

function Initialize-OptionsFoldersSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $root = [string]$PackagersRoot
    & $Page.Editors['pkg.packagers'].Set $root
    $Page.Editors['pkg.packagers'].Buttons['open'].Add_Click({
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $root) } catch { <# Opening a folder in the shell is a best-effort convenience. #> $null = $_ }
        }.GetNewClosure())
}

function Initialize-OptionsToolsSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $prefs = $Context.Prefs
    $console = Get-OptionsToolStatusText -Prefs $prefs -Tool Console
    & $Page.Editors['tools.console'].Set $console.Text
    $Page.Editors['tools.console'].Text.ToolTip = $console.Tip
    $sevenZip = Get-OptionsToolStatusText -Prefs $prefs -Tool SevenZip
    & $Page.Editors['tools.sevenzip'].Set $sevenZip.Text
    $Page.Editors['tools.sevenzip'].Text.ToolTip = $sevenZip.Tip
    Start-OptionsGitHubCheck -StatusText $Page.Editors['gh.status'].Text

    $prepEditor = $Page.Editors['tools.contentprep']
    $downloadButton = $prepEditor.Buttons['download']
    $showPrep = {
        $prep = Get-OptionsToolStatusText -Prefs $prefs -Tool ContentPrep
        & $prepEditor.Set $prep.Text
        $prepEditor.Text.ToolTip = $prep.Tip
        $downloadButton.Visibility = $(if ($prefs.DetectedTools.IntuneWinAppUtil -and $prefs.DetectedTools.IntuneWinAppUtil.Found) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible })
    }.GetNewClosure()
    & $showPrep
    $downloadButton.Add_Click({
            $downloadButton.IsEnabled = $false
            & $prepEditor.Set 'Downloading the Microsoft Win32 Content Prep Tool...'
            $downloaded = $false
            try {
                $exePath = Install-IntuneWinAppUtil -DestinationFolder (Get-IntuneWinToolCachePath)
                $prefs.DetectedTools.IntuneWinAppUtil = Invoke-DetectIntuneWinAppUtil -KnownPath $exePath
                Save-Preferences -Prefs $prefs
                $downloaded = $true
            }
            catch { & $prepEditor.Set ([string][char]0x2717 + ' Download failed  -  ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
            finally { $downloadButton.IsEnabled = $true }
            if ($downloaded) { & $showPrep }
        }.GetNewClosure())

    $iconEditor = $Page.Editors['tools.iconpack']
    $showIcon = {
        & $iconEditor.Set (Get-IconPackStatusText -Manifest (Read-IconPackManifest -Path (Get-IconPackManifestPath)))
        $iconEditor.Text.ToolTip = ('Manifest: {0}' -f (Get-IconPackManifestPath))
    }.GetNewClosure()
    & $showIcon
    $iconDownload = $iconEditor.Buttons['download']
    $iconFile = $iconEditor.Buttons['file']
    $runIcon = {
        param($Action)
        $iconDownload.IsEnabled = $false
        $iconFile.IsEnabled = $false
        $outcome = $null
        try { $outcome = & $Action; Add-LogLine -Message $outcome.Message }
        finally {
            $iconDownload.IsEnabled = $true
            $iconFile.IsEnabled = $true
            & $showIcon
            # A failed fetch leaves no manifest change, so the refreshed line would repeat itself; the reason replaces it.
            if ($outcome -and -not $outcome.Installed) { & $iconEditor.Set ([string][char]0x2717 + ' ' + $outcome.Message) }
        }
    }.GetNewClosure()
    $iconDownload.Add_Click({
            & $iconEditor.Set 'Downloading the packager icon pack...'
            & $runIcon { Install-IconPack }
        }.GetNewClosure())
    $iconFile.Add_Click({
            $open = New-Object Microsoft.Win32.OpenFileDialog
            $open.Title = 'Select Icon Pack Zip'
            $open.Filter = 'Icon pack (*.zip)|*.zip'
            if (-not $open.ShowDialog()) { return }
            $zip = $open.FileName
            & $iconEditor.Set 'Installing the icon pack from the file...'
            & $runIcon ({ Install-IconPackFromFile -ZipPath $zip }.GetNewClosure())
        }.GetNewClosure())
}

function Start-OptionsGitHubCheck {
    # The check needs a web request to api.github.com and may start gh.exe; a slow proxy would hold the
    # window for seconds, so a background runspace answers and a timer fills the line in.
    param([Parameter(Mandatory)]$StatusText)
    $show = {
        param($gh)
        $quota = ''
        if ($gh.Limit) {
            $quota = ' (' + $gh.Limit + ' requests/hour'
            if ($gh.Remaining) { $quota += ', ' + $gh.Remaining + ' remaining' }
            $quota += ')'
        }
        if ($gh.Authenticated) { $StatusText.Text = ([string][char]0x2713 + ' Authenticated  -  ' + $gh.Source + $quota) }
        else {
            $remaining = $(if ($gh.Remaining) { ', ' + $gh.Remaining + ' remaining' } else { '' })
            $StatusText.Text = ([string][char]0x2717 + ' Anonymous  -  60 requests/hour per address' + $remaining + '; type a token below, set GITHUB_TOKEN or run gh auth login')
        }
    }.GetNewClosure()
    $StatusText.Text = 'Checking GitHub access...'
    $state = [hashtable]::Synchronized(@{ Done = $false; Result = $null })
    $runspace = $null
    $shell = $null
    try {
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = 'MTA'
        $runspace.Open()
        $shell = [powershell]::Create()
        $shell.Runspace = $runspace
        [void]$shell.AddScript({
                param($State, $ModulePath, $FunctionText)
                try {
                    Import-Module -Name $ModulePath -DisableNameChecking -ErrorAction Stop
                    Set-Item -Path 'function:Get-GitHubApiAuthStatus' -Value ([scriptblock]::Create($FunctionText))
                    $State.Result = Get-GitHubApiAuthStatus
                }
                catch { $State.Result = $null }
                finally { $State.Done = $true }
            }).AddArgument($state).AddArgument((Join-Path $PSScriptRoot 'Packagers\AppPackagerCommon.psd1')).AddArgument((Get-Item -Path 'function:Get-GitHubApiAuthStatus').ScriptBlock.ToString())
        $handle = $shell.BeginInvoke()
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(250)
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        $timer.Add_Tick({
                if (-not $state.Done) {
                    if ([DateTime]::UtcNow -lt $deadline) { return }
                    $timer.Stop()
                    $StatusText.Text = 'GitHub access could not be checked: no answer within 20 seconds.'
                    # Stopping a pipeline that waits on a network call blocks until it returns, so the stop is asynchronous.
                    try { [void]$shell.BeginStop($null, $null) } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ }
                    try { [void]$runspace.CloseAsync() } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ }
                    return
                }
                $timer.Stop()
                if ($state.Result) { & $show $state.Result } else { $StatusText.Text = 'GitHub access could not be checked.' }
                try { $shell.EndInvoke($handle) } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ }
                try { $shell.Dispose() } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ }
                try { $runspace.Close() } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ }
            }.GetNewClosure())
        $timer.Start()
    }
    catch {
        $StatusText.Text = 'GitHub access could not be checked.'
        if ($shell) { try { [void]$shell.BeginStop($null, $null) } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ } }
        if ($runspace) { try { [void]$runspace.CloseAsync() } catch { <# Stopping and closing the GitHub check runspace is best-effort cleanup. #> $null = $_ } }
    }
}


# ---------------------------------------------------------------------------
# WSUS and the connection tests
# ---------------------------------------------------------------------------
function Get-OptionsWsusSetting {
    # The WSUS settings as the fields hold them now, validated by the WSUS module.
    param([Parameter(Mandatory)][hashtable]$Context)
    $classification = Get-OptionsLiveText -Context $Context -FieldId 'wsus.classification'
    return (ConvertTo-WsusPublishSettings -InputObject @{
            ServerName        = (Get-OptionsLiveText -Context $Context -FieldId 'wsus.server')
            PortNumber        = (Get-OptionsLiveText -Context $Context -FieldId 'wsus.port')
            UseSsl            = (Get-OptionsLiveFlag -Context $Context -FieldId 'wsus.ssl')
            Classification    = $(if ($classification) { $classification } else { 'Updates' })
            ApprovalGroup     = (Get-OptionsLiveText -Context $Context -FieldId 'wsus.group')
            DeclineSuperseded = (Get-OptionsLiveFlag -Context $Context -FieldId 'wsus.decline')
        })
}

function Format-OptionsWsusStatus {
    # Two lines for one status: the connection and the signing certificate.
    param([Parameter(Mandatory)]$Status)
    $check = [string][char]0x2713
    $cross = [string][char]0x2717
    if (-not $Status.Connected) { return @{ Connection = ($cross + ' ' + [string]$Status.Message); Certificate = 'Not connected.' } }
    $connection = ('{0} Connected  -  {1}, WSUS {2}, role {3}, {4} connection. {5}' -f $check, [string]$Status.ServerName, [string]$Status.Version, [string]$Status.Role, $(if ($Status.SecureConnection) { 'secure' } else { 'unencrypted' }), [string]$Status.Message).Trim()
    $cert = $Status.SigningCertificate
    if (-not $cert) {
        $text = $(if ([string]$Status.Message -match 'could not be read') { 'The signing certificate could not be read; see the connection line.' } else { $cross + ' No signing certificate is set on this server. Publishing fails until one is created or imported.' })
        return @{ Connection = $connection; Certificate = $text }
    }
    $flags = @()
    if ($cert.SelfSigned) { $flags += 'self-signed' }
    if ($cert.Expired) { $flags += 'EXPIRED' }
    $certText = ('Subject: {0}   Thumbprint: {1}   Valid to: {2}   Key: {3} bits{4}' -f [string]$cert.Subject, [string]$cert.Thumbprint, ([datetime]$cert.NotAfter).ToString('yyyy-MM-dd'), [string]$cert.KeyLength, $(if ($flags) { '   (' + ($flags -join ', ') + ')' } else { '' }))
    return @{ Connection = $connection; Certificate = $certText }
}

function Initialize-OptionsWsusSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $ctx = $Context
    $certEditor = $Page.Editors['wsus.cert']
    $buttons = $Page.Editors['wsus.certActions'].Buttons
    $groupEditor = $Page.Editors['wsus.group']
    $loadButton = $Page.Editors['wsus.groupLoad'].Buttons['load']
    & $certEditor.Set 'Test Connection reads the signing certificate of the server.'

    $api = $(if (Get-Command -Name Test-WsusAdministrationApi -ErrorAction SilentlyContinue) { Test-WsusAdministrationApi }
        else { [pscustomobject]@{ Available = $false; Version = ''; Reason = 'The WSUS publishing module did not load; reinstall AppPackager.' } })
    $serverButtons = @($buttons['create'], $buttons['import'], $buttons['export'], $loadButton)
    if (-not $api.Available) {
        foreach ($b in $serverButtons) { $b.IsEnabled = $false }
        & $certEditor.Set ([string][char]0x2717 + ' ' + [string]$api.Reason)
    }

    # Test Connection fills the certificate line when this section is built.
    $Context.OnWsusStatus = {
        param($Status)
        $formatted = Format-OptionsWsusStatus -Status $Status
        & $certEditor.Set $formatted.Certificate
    }.GetNewClosure()

    $runOnServer = {
        param([scriptblock]$Action, $Extra)
        $window = $ctx.Owner
        if ($window) { Set-DialogBusyStatus -Dialog $window -Control $certEditor.Text -Text 'Contacting the WSUS server...' }
        try {
            $settings = Get-OptionsWsusSetting -Context $ctx
            if ([string]::IsNullOrWhiteSpace([string]$settings.ServerName)) { throw 'Enter the WSUS server first.' }
            return (& $Action $settings $Extra)
        }
        catch { & $certEditor.Set ([string][char]0x2717 + ' ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')); return $null }
        finally { if ($window) { $window.Cursor = $null } }
    }.GetNewClosure()

    $showStatus = {
        param($Status)
        $formatted = Format-OptionsWsusStatus -Status $Status
        & $certEditor.Set ($formatted.Connection + [Environment]::NewLine + $formatted.Certificate)
    }.GetNewClosure()

    $buttons['create'].Add_Click({
            $settings = & $runOnServer { param($s) $s }
            if (-not $settings) { return }
            $answer = Show-ThemedMessage -Owner $ctx.Owner -Title 'Create Signing Certificate' -Buttons YesNo -Icon Warning `
                -Message ("Have {0} create a self-signed signing certificate and use it for every update published from now on?`n`nThis replaces any certificate the server uses today. Clients that trust only the current certificate reject updates signed with the new one until the new certificate reaches their Trusted Publishers and Trusted Root Certification Authorities stores. With ConfigMgr, keep ConfigMgr in charge of the certificate if it already manages it." -f [string]$settings.ServerName)
            if ($answer -ne 'Yes') { return }
            $status = & $runOnServer { param($s) New-WsusSelfSignedSigningCertificate -Settings $s }
            if ($status) { & $showStatus $status }
        }.GetNewClosure())

    $buttons['import'].Add_Click({
            $settings = & $runOnServer { param($s) $s }
            if (-not $settings) { return }
            $open = New-Object Microsoft.Win32.OpenFileDialog
            $open.Title = 'Select the Signing Certificate PFX'
            $open.Filter = 'PFX files (*.pfx;*.p12)|*.pfx;*.p12'
            if (-not $open.ShowDialog()) { return }
            $password = Show-WsusPasswordPrompt -Owner $ctx.Owner -Prompt ('Password for {0}' -f [System.IO.Path]::GetFileName($open.FileName))
            if ($null -eq $password) { return }
            try {
                $answer = Show-ThemedMessage -Owner $ctx.Owner -Title 'Import Signing Certificate' -Buttons YesNo -Icon Warning `
                    -Message ("Make this certificate the signing certificate of {0}?`n`nIt replaces any certificate the server uses today. Clients must trust the new certificate before they install updates signed with it." -f [string]$settings.ServerName)
                if ($answer -ne 'Yes') { return }
                $status = & $runOnServer { param($s, $x) Set-WsusSigningCertificate -Settings $s -PfxPath $x.Path -Password $x.Password } @{ Path = $open.FileName; Password = $password }
                if ($status) { & $showStatus $status }
            }
            finally { $password.Dispose() }
        }.GetNewClosure())

    $buttons['export'].Add_Click({
            $settings = & $runOnServer { param($s) $s }
            if (-not $settings) { return }
            $save = New-Object Microsoft.Win32.SaveFileDialog
            $save.Title = 'Save the Public Signing Certificate'
            $save.Filter = 'Certificate (*.cer)|*.cer'
            $save.FileName = ('WSUS-signing-{0}.cer' -f [string]$settings.ServerName)
            if (-not $save.ShowDialog()) { return }
            $written = & $runOnServer { param($s, $x) Export-WsusSigningCertificate -Settings $s -Path $x } $save.FileName
            if ($written) { & $certEditor.Set ([string][char]0x2713 + ' Saved ' + [string]$written) }
        }.GetNewClosure())

    $loadButton.Add_Click({
            $result = & $runOnServer { param($s) [pscustomobject]@{ Names = @(Get-WsusComputerGroupNames -Settings $s) } }
            if ($null -eq $result) { return }
            $typed = [string]$groupEditor.Combo.Text
            $groupEditor.Combo.Items.Clear()
            foreach ($name in @($result.Names)) { [void]$groupEditor.Combo.Items.Add($name) }
            $groupEditor.Combo.Text = $typed
            & $certEditor.Set ('{0} computer group(s) loaded.' -f @($result.Names).Count)
        }.GetNewClosure())
}

function Test-OptionsConfigMgrConnection {
    # Opens the SMS Provider and the site database with the values given. Never throws.
    param([string]$SiteCode, [string]$Provider, [string]$SqlServer, [string]$SqlInstance, [int]$SqlPort, [string]$Database)
    $check = [string][char]0x2713
    $cross = [string][char]0x2717
    $connection = $null
    try {
        if ($SiteCode -notmatch '^[A-Za-z0-9]{3}$') { return ($cross + ' Enter the three-character site code first.') }
        $providerName = Resolve-DiscoveryProviderName -Preference $Provider -SiteCode $SiteCode
        $arguments = @{ SiteCode = $SiteCode; ServerOverride = $SqlServer; DatabaseOverride = $Database; InstanceOverride = $SqlInstance; PortOverride = $SqlPort }
        if ($providerName) { $arguments.ProviderMachineName = $providerName }
        $target = Resolve-CMDatabaseTarget @arguments
        $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $builder['Data Source'] = [string]$target.Server
        $builder['Initial Catalog'] = [string]$target.Database
        $builder['Integrated Security'] = $true
        $builder['ApplicationIntent'] = 'ReadOnly'
        $builder['Application Name'] = 'AppPackager Options'
        $builder['Connect Timeout'] = 10
        $connection = New-Object System.Data.SqlClient.SqlConnection($builder.ConnectionString)
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = 'SELECT COUNT(*) FROM dbo.v_Site WHERE SiteCode = @code'
        $command.CommandTimeout = 10
        $parameter = $command.Parameters.Add('@code', [System.Data.SqlDbType]::NVarChar, 3)
        $parameter.Value = $SiteCode.ToUpperInvariant()
        $found = [int]$command.ExecuteScalar()
        if ($found -lt 1) { return ('{0} The database {1} on {2} opened, and it has no site {3}.' -f $cross, $target.Database, $target.Server, $SiteCode.ToUpperInvariant()) }
        $via = $(if ($providerName) { 'SMS Provider ' + $providerName + ', ' } else { '' })
        return ('{0} Site {1}: {2}database {3} on {4} opened.' -f $check, $SiteCode.ToUpperInvariant(), $via, $target.Database, $target.Server)
    }
    catch { return ($cross + ' ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The connection failed. See the log file.')) }
    finally { if ($connection) { $connection.Dispose() } }
}

function Test-OptionsIntuneConnection {
    # Requests an app-only Graph token with the values given. Never throws.
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    $check = [string][char]0x2713
    $cross = [string][char]0x2717
    if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) { return ($cross + ' Enter the tenant ID, the client ID and the client secret first.') }
    try {
        $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        if ($token) { return ($check + ' The tenant answered and issued a Microsoft Graph token.') }
        return ($cross + ' The tenant issued no token.')
    }
    catch { return ($cross + ' ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The token request failed. See the log file.')) }
}

function Initialize-OptionsTestSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $ctx = $Context
    $cmEditor = $Page.Editors['test.configmgr']
    $wsusEditor = $Page.Editors['test.wsus']
    $intuneEditor = $Page.Editors['test.intune']
    foreach ($e in @($cmEditor, $wsusEditor, $intuneEditor)) { & $e.Set 'Not tested.' }

    $busy = {
        param($Editor, [scriptblock]$Work)
        $Editor.Buttons['run'].IsEnabled = $false
        if ($ctx.Owner) { Set-DialogBusyStatus -Dialog $ctx.Owner -Control $Editor.Text -Text 'Testing...' }
        else { & $Editor.Set 'Testing...' }
        try { & $Editor.Set (& $Work) }
        finally {
            if ($ctx.Owner) { $ctx.Owner.Cursor = $null }
            $Editor.Buttons['run'].IsEnabled = $true
        }
    }.GetNewClosure()

    $cmEditor.Buttons['run'].Add_Click({
            & $busy $cmEditor ({
                $port = 0
                [void][int]::TryParse((Get-OptionsLiveText -Context $ctx -FieldId 'cm.sqlPort'), [ref]$port)
                Test-OptionsConfigMgrConnection -SiteCode (Get-OptionsLiveText -Context $ctx -FieldId 'cm.siteCode') -Provider (Get-OptionsLiveText -Context $ctx -FieldId 'cm.provider') `
                    -SqlServer (Get-OptionsLiveText -Context $ctx -FieldId 'cm.sqlServer') -SqlInstance (Get-OptionsLiveText -Context $ctx -FieldId 'cm.sqlInstance') -SqlPort $port `
                    -Database (Get-OptionsLiveText -Context $ctx -FieldId 'cm.sqlDatabase')
            }.GetNewClosure())
        }.GetNewClosure())

    $wsusEditor.Buttons['run'].Add_Click({
            & $busy $wsusEditor ({
                try {
                    $settings = Get-OptionsWsusSetting -Context $ctx
                    if ([string]::IsNullOrWhiteSpace([string]$settings.ServerName)) { return ([string][char]0x2717 + ' Enter the WSUS server first.') }
                    $status = Get-WsusServerStatus -Settings $settings
                    if ($ctx.OnWsusStatus) { & $ctx.OnWsusStatus $status }
                    $formatted = Format-OptionsWsusStatus -Status $status
                    return $formatted.Connection
                }
                catch { return ([string][char]0x2717 + ' ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
            }.GetNewClosure())
        }.GetNewClosure())

    $intuneEditor.Buttons['run'].Add_Click({
            & $busy $intuneEditor ({
                try {
                    Test-OptionsIntuneConnection -TenantId (Get-OptionsLiveText -Context $ctx -FieldId 'int.tenant') -ClientId (Get-OptionsLiveText -Context $ctx -FieldId 'int.client') -ClientSecret (Get-OptionsLiveSecret -Context $ctx -FieldId 'int.secret')
                }
                catch { [string][char]0x2717 + ' ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
            }.GetNewClosure())
        }.GetNewClosure())
}

# ---------------------------------------------------------------------------
# Applications
# ---------------------------------------------------------------------------
function Initialize-OptionsAdobeSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $edition = $Page.Editors['adobe.edition'].Focus
    $languages = $Page.Editors['adobe.languages']
    $all = $languages.Boxes['All']
    $others = @($languages.Boxes.Keys | Where-Object { $_ -ne 'All' } | ForEach-Object { $languages.Boxes[$_] })
    # Multilingual installs the languages that are checked; the English installer takes none.
    $refresh = {
        $mui = ($edition.SelectedItem -and [string]$edition.SelectedItem.Tag -eq 'MUI')
        $all.IsEnabled = $mui
        foreach ($box in $others) { $box.IsEnabled = ($mui -and $all.IsChecked -ne $true) }
    }.GetNewClosure()
    $edition.Add_SelectionChanged($refresh)
    $all.Add_Checked($refresh)
    $all.Add_Unchecked($refresh)
    & $refresh
}

function Initialize-OptionsSsmsSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $ui = $Page.Editors['ssms.ui'].Focus
    $noUpdate = $Page.Editors['ssms.noupdate'].Focus
    $workloads = $Page.Editors['ssms.workloads']
    $recommended = $Page.Editors['ssms.recommended'].Focus
    $optional = $Page.Editors['ssms.optional'].Focus
    $boxes = @($workloads.Boxes.Values)
    # --noUpdateInstaller is documented with --quiet, and the component switches act on workloads only.
    $refresh = {
        $quiet = (-not $ui.SelectedItem -or [string]$ui.SelectedItem.Tag -eq 'Quiet')
        $noUpdate.IsEnabled = $quiet
        $any = (@($boxes | Where-Object { $_.IsChecked -eq $true }).Count -gt 0)
        $recommended.IsEnabled = $any
        $optional.IsEnabled = $any
    }.GetNewClosure()
    $ui.Add_SelectionChanged($refresh)
    foreach ($box in $boxes) { $box.Add_Checked($refresh); $box.Add_Unchecked($refresh) }
    & $refresh
}

function Initialize-OptionsCitrixSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    Initialize-OptionsCitrixInstallSection -Context $Context -Page $Page
    Initialize-OptionsCitrixPluginSection -Context $Context -Page $Page
    Initialize-OptionsCitrixComponentSection -Context $Context -Page $Page
    Initialize-OptionsCitrixPreview -Context $Context -Page $Page
}

function Initialize-OptionsCitrixInstallSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    # ENABLE_SSON follows the SSO component: the packager writes it only together with /includeSSON.
    $sso = $Page.Editors['cwa.sso'].Focus
    $Page.Group.ExtraCommit.Add({
            param($ctx)
            Set-OptionsPrefValue -Context $ctx -Path 'cwa:Installation.EnableSSON' -Value ($sso.IsChecked -eq $true)
        }.GetNewClosure()) | Out-Null
}

function Initialize-OptionsCitrixPluginSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $uber = $Page.Editors['cwa.uber'].Focus
    $uberSkip = $Page.Editors['cwa.uberskip'].Focus
    $sr = $Page.Editors['cwa.sr'].Focus
    $srSkip = $Page.Editors['cwa.srskip'].Focus
    $refresh = {
        $uberSkip.IsEnabled = ($uber.IsChecked -eq $true)
        $srSkip.IsEnabled = ($sr.IsChecked -eq $true)
    }.GetNewClosure()
    foreach ($box in @($uber, $sr)) { $box.Add_Checked($refresh); $box.Add_Unchecked($refresh) }
    & $refresh
}

function Initialize-OptionsCitrixComponentSection {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $customize = $Page.Editors['cwa.customize'].Focus
    $required = @('ReceiverInside', 'ICA_Client', 'AM')
    $componentBoxes = @{}
    foreach ($id in @($Page.Editors.Keys | Where-Object { $_ -like 'cwa.comp.*' })) { $componentBoxes[$id.Substring(9)] = $Page.Editors[$id].Focus }
    foreach ($name in $required) { $componentBoxes[$name].IsChecked = $true }
    $refresh = {
        foreach ($name in $componentBoxes.Keys) { $componentBoxes[$name].IsEnabled = ($customize.IsChecked -eq $true) }
        if ($customize.IsChecked -ne $true) { foreach ($name in $required) { $componentBoxes[$name].IsChecked = $true } }
    }.GetNewClosure()
    $customize.Add_Checked($refresh)
    $customize.Add_Unchecked($refresh)
    & $refresh
    # The required components are always written.
    $Page.Group.ExtraCommit.Add({
            param($ctx)
            foreach ($name in $required) { Set-OptionsPrefValue -Context $ctx -Path ('cwa:Components.' + $name) -Value $true }
        }.GetNewClosure()) | Out-Null
}

function Initialize-OptionsCitrixPreview {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $ctx = $Context
    $button = $Page.Editors['cwa.preview'].Buttons['preview']
    $button.Add_Click({
            $owner = $ctx.Owner
            try {
                # The sections hold the settings in a working copy that OK saves; the preview reads the copy after the open sections write to it.
                $citrixPage = $ctx.Sections['cwa']
                if ($citrixPage -and $citrixPage.Commit) { & $citrixPage.Commit }
                $packagerPath = Join-Path $PackagersRoot 'package-citrixworkspace-cr.ps1'
                $pt = $null; $pe = $null
                $ast = [System.Management.Automation.Language.Parser]::ParseFile($packagerPath, [ref]$pt, [ref]$pe)
                $function = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-CwaInstallArguments' }, $false)
                if ($null -eq $function) { throw "ConvertTo-CwaInstallArguments not found in $packagerPath" }
                . ([scriptblock]::Create($function.Extent.Text))
                $text = [System.Text.StringBuilder]::new()
                foreach ($stream in @('Current', 'LTSR')) {
                    $built = ConvertTo-CwaInstallArguments -Config $ctx.Aux.cwa -Stream $stream
                    [void]$text.AppendLine(('# {0}' -f $(if ($stream -eq 'Current') { 'Current Release' } else { 'LTSR (x86, x64, ARM64)' })))
                    [void]$text.AppendLine('CitrixWorkspaceApp.exe ' + ($built.Arguments -join ' '))
                    foreach ($w in $built.Warnings) { [void]$text.AppendLine('WARN: ' + $w) }
                    [void]$text.AppendLine('')
                }
                Show-PreviewDialog -Owner $owner -Title 'CWA Preview' -Content $text.ToString() -Width 820 -Height 360
            }
            catch {
                Show-PreviewDialog -Owner $owner -Title 'CWA Preview (Error)' -Content ("Failed to build preview:`r`n{0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) -Width 600 -Height 260
            }
        }.GetNewClosure())
}

function New-OptionsLinkEditor {
    # A row of buttons that open a web page.
    param([Parameter(Mandatory)][object[]]$Link)
    $panel = New-Object System.Windows.Controls.WrapPanel
    $first = $null
    foreach ($item in $Link) {
        $url = [string]$item.Url
        $button = New-OptionsButton -Label ([string]$item.Label) -Tip $url
        $button.Margin = New-Object System.Windows.Thickness(0, 0, 8, 4)
        $button.Add_Click({ try { Start-Process -FilePath $url } catch { <# Opening a link in the shell is a best-effort convenience. #> $null = $_ } }.GetNewClosure())
        [void]$panel.Children.Add($button)
        if (-not $first) { $first = $button }
    }
    return @{ Element = $panel; Focus = $first; Get = { $null }; Set = { param($v) } }
}

function New-OptionsLocalSourceEditor {
    # One folder box for each packager that can stage a local installer.
    param([Parameter(Mandatory)][hashtable]$Context)
    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.MaxWidth = 640
    $boxes = [ordered]@{}
    $first = $null
    foreach ($row in @(Get-OptionsPackagerList -Context $Context | Where-Object { $_.LocalSource })) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script)
        $caption = New-OptionsTextBlock -Text ('{0}  ({1})' -f [string]$row.Application, [string]$row.Script) -Size 12
        $caption.Margin = New-Object System.Windows.Thickness(0, 0, 0, 2)
        $tip = ('The folder that holds the installer for {0}.' -f [string]$row.Script)
        $caption.ToolTip = $tip
        $box = New-Object System.Windows.Controls.TextBox
        $box.FontSize = 13
        $box.ToolTip = $tip
        $box.Height = (Get-OptionsInputHeight)
        $box.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Center
        $box.MinWidth = 160
        if ($Context.Prefs.LocalSourceFolders -and $Context.Prefs.LocalSourceFolders.PSObject.Properties[$base]) { $box.Text = [string]$Context.Prefs.LocalSourceFolders.$base }
        $browse = New-OptionsButton -Label 'Browse...' -Tip 'Choose the folder.'
        $browse.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
        $browse.Add_Click({ $picked = Select-OptionsFolder -Current $box.Text; if ($picked) { $box.Text = $picked } }.GetNewClosure())
        $line = New-Object System.Windows.Controls.DockPanel
        $line.LastChildFill = $true
        $line.MaxWidth = 640
        $line.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
        [System.Windows.Controls.DockPanel]::SetDock($browse, [System.Windows.Controls.Dock]::Right)
        [void]$line.Children.Add($browse)
        [void]$line.Children.Add($box)
        $line.Margin = New-Object System.Windows.Thickness(0, 0, 0, 8)
        [void]$stack.Children.Add($caption)
        [void]$stack.Children.Add($line)
        $boxes[$base] = $box
        if (-not $first) { $first = $box }
    }
    if ($boxes.Count -eq 0) { [void]$stack.Children.Add((New-OptionsTextBlock -Text 'No packager of this catalog stages a local installer.' -Size 12)) }
    $commit = {
        param($ctx)
        if ($boxes.Count -eq 0) { return }
        $folders = [ordered]@{}
        if ($ctx.Prefs.LocalSourceFolders) { foreach ($p in $ctx.Prefs.LocalSourceFolders.PSObject.Properties) { $folders[$p.Name] = $p.Value } }
        foreach ($key in $boxes.Keys) {
            $value = [string]$boxes[$key].Text.Trim()
            if ($value) { $folders[$key] = $value }
            elseif ($folders.Contains($key)) { $folders.Remove($key) }
        }
        $ctx.Prefs.LocalSourceFolders = [pscustomobject]$folders
    }.GetNewClosure()
    return @{ Element = $stack; Focus = $first; Get = $null; Set = $null; Commit = $commit }
}

# =============================================================================
# Options window: Microsoft 365 Apps section
# =============================================================================
function Get-M365CustomConfigPath {
    Join-Path (Join-Path $PSScriptRoot 'Packagers') 'm365-configuration.user.xml'
}

function Format-OptionsFileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return ('{0} bytes' -f $Bytes)
}

function Initialize-OptionsM365Section {
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][hashtable]$Page)
    $ctx = $Context
    $e = $Page.Editors
    $useBox = $e['m365.useCustom'].Focus
    $fileEditor = $e['m365.file']
    $importButton = $fileEditor.Buttons['import']
    $removeButton = $fileEditor.Buttons['remove']
    $reasonText = $e['m365.reason'].Text
    $reasonText.Text = 'The imported file sets these. The packager still sets the version, the source path, the architecture and a silent display, and keeps its own detection.'
    $generated = @('m365.company', 'm365.channel', 'm365.deploy', 'm365.languages', 'm365.exclude', 'm365.officeMgmt', 'm365.updates', 'm365.forceShutdown', 'm365.removeMsi', 'm365.migrate', 'm365.shared', 'm365.device') | ForEach-Object { $e[$_] }

    # State: the stored file, a file that was imported in this window, or a removal. OK applies it, Cancel drops it.
    $storedPath = Get-M365CustomConfigPath
    $stored = $Context.Prefs.M365CustomConfig
    $state = @{ Name = ''; Size = 0L; Pending = $null; Removed = $false; HasStored = ($stored -and (Test-Path -LiteralPath $storedPath -PathType Leaf)) }
    if ($state.HasStored) { $state.Name = [string]$stored.FileName; $state.Size = [long](Get-Item -LiteralPath $storedPath).Length }

    $refresh = {
        $hasFile = ($state.Pending -ne $null) -or ($state.HasStored -and -not $state.Removed)
        if ($hasFile) { & $fileEditor.Set ('{0}  ({1})' -f $state.Name, (Format-OptionsFileSize -Bytes $state.Size)) }
        else { & $fileEditor.Set 'No file imported.' }
        $removeButton.IsEnabled = $hasFile
        $useBox.IsEnabled = $hasFile
        if (-not $hasFile) { $useBox.IsChecked = $false }
        $custom = ($useBox.IsChecked -eq $true)
        foreach ($editor in $generated) { $editor.Element.IsEnabled = -not $custom }
        $reasonText.Visibility = $(if ($custom) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed })
    }.GetNewClosure()
    $useBox.Add_Checked($refresh)
    $useBox.Add_Unchecked($refresh)
    & $refresh

    # The file is read, decoded by its byte order mark, checked and held in the window until OK stores it.
    $importFile = {
        param([string]$Path)
        try {
            $info = Get-Item -LiteralPath $Path -ErrorAction Stop
            if ($info.Length -gt 2MB) { [void](Show-ThemedMessage -Owner $ctx.Owner -Title 'Import Failed' -Message 'The file is larger than 2 MB. A configuration.xml is a small text file.' -Buttons OK -Icon Error); return $false }
            $reader = New-Object System.IO.StreamReader($info.FullName, [System.Text.Encoding]::UTF8, $true)
            try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $result = Test-OdtConfigurationXml -Text $text
            if (-not $result.Valid) {
                [void](Show-ThemedMessage -Owner $ctx.Owner -Title 'Import Failed' -Message ("The file is not a usable Office Deployment Tool configuration:`n`n" + (($result.Errors | ForEach-Object { '- ' + $_ }) -join "`n")) -Buttons OK -Icon Error)
                return $false
            }
            # The stored copy is UTF-8 without a byte order mark.
            $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($text.TrimStart([char]0xFEFF))
            $state.Pending = @{ Bytes = $bytes; Text = $text }
            $state.Removed = $false
            $state.Name = $info.Name
            $state.Size = [long]$bytes.Length
            & $refresh
            if ($result.Notes.Count -gt 0) {
                [void](Show-ThemedMessage -Owner $ctx.Owner -Title 'File Imported' -Message ("The packager changes these parts of the file:`n`n" + (($result.Notes | ForEach-Object { '- ' + $_ }) -join "`n")) -Buttons OK -Icon Info)
            }
            return $true
        }
        catch {
            [void](Show-ThemedMessage -Owner $ctx.Owner -Title 'Import Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The file could not be read.') -Buttons OK -Icon Error)
            return $false
        }
    }.GetNewClosure()
    $Page.Actions = @{ ImportM365 = $importFile }
    $importButton.Add_Click({
            $open = New-Object Microsoft.Win32.OpenFileDialog
            $open.Title = 'Select Configuration File'
            $open.Filter = 'Office Deployment Tool configuration (*.xml)|*.xml|All files (*.*)|*.*'
            if ($open.ShowDialog() -eq $true) { [void](& $importFile $open.FileName) }
        }.GetNewClosure())

    $removeButton.Add_Click({
            $state.Pending = $null
            $state.Removed = $true
            & $refresh
        }.GetNewClosure())

    # OK stores the file next to the other packager files and records its name and size.
    $Page.Group.ExtraCommit.Add({
            param($context)
            $path = Get-M365CustomConfigPath
            $record = $context.Prefs.M365CustomConfig
            if ($state.Pending) {
                [System.IO.File]::WriteAllBytes($path, $state.Pending.Bytes)
                $record.FileName = $state.Name
                $record.SizeBytes = [long]$state.Size
                $state.HasStored = $true
                $state.Pending = $null
            }
            elseif ($state.Removed) {
                if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
                $record.FileName = ''
                $record.SizeBytes = 0
                $record.Enabled = $false
                $state.HasStored = $false
            }
        }.GetNewClosure()) | Out-Null

    # The preview shows the install XML of each product with the values in the fields, or the imported file with the packager changes.
    $previewButton = $e['m365.preview'].Buttons['preview']
    $previewButton.Add_Click({
            $owner = $ctx.Owner
            try {
                if (-not (Get-Command -Name New-OdtConfigXml -ErrorAction SilentlyContinue)) {
                    Show-PreviewDialog -Owner $owner -Title 'M365 Preview (Error)' -Content 'New-OdtConfigXml not available. Ensure Packagers\AppPackagerCommon.psm1 is importable.' -Width 600 -Height 220
                    return
                }
                $builder = [System.Text.StringBuilder]::new()
                $products = @(
                    @{ Label = 'M365 Apps for Enterprise (x64)'; Edition = '64'; Ids = @('O365ProPlusRetail') },
                    @{ Label = 'M365 Apps for Enterprise (x86)'; Edition = '32'; Ids = @('O365ProPlusRetail') },
                    @{ Label = 'Project Pro (x64)'; Edition = '64'; Ids = @('ProjectProRetail') },
                    @{ Label = 'Visio Pro (x64)'; Edition = '64'; Ids = @('VisioProRetail') })
                $channel = Get-OptionsLiveText -Context $ctx -FieldId 'm365.channel'
                if (-not $channel) { $channel = 'MonthlyEnterprise' }
                $custom = ($useBox.IsChecked -eq $true)
                $customText = ''
                if ($custom) { $customText = $(if ($state.Pending) { [string]$state.Pending.Text } else { [System.IO.File]::ReadAllText((Get-M365CustomConfigPath), [System.Text.Encoding]::UTF8) }) }
                $options = @{
                    Languages               = @(([string](Get-OptionsLiveText -Context $ctx -FieldId 'm365.languages')) -split '[,;\s]+' | Where-Object { $_ })
                    OfficeMgmtCom           = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.officeMgmt')
                    UpdatesEnabled          = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.updates')
                    ForceAppShutdown        = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.forceShutdown')
                    RemoveMsi               = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.removeMsi')
                    MigrateArch             = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.migrate')
                    SharedComputerLicensing = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.shared')
                    DeviceBasedLicensing    = (Get-OptionsLiveFlag -Context $ctx -FieldId 'm365.device')
                }
                if ($options.Languages.Count -eq 0) { $options.Remove('Languages') }
                $excluded = @(& $e['m365.exclude'].Get)
                foreach ($product in $products) {
                    [void]$builder.AppendLine(('# ===== {0} =====' -f $product.Label))
                    if ($custom) { $xml = ConvertTo-OdtUserConfigXml -Text $customText -OfficeClientEdition $product.Edition -Channel $channel }
                    else { $xml = New-OdtConfigXml -OfficeClientEdition $product.Edition -ProductIds $product.Ids -Channel $channel -CompanyName (Get-OptionsLiveText -Context $ctx -FieldId 'm365.company') -ExcludeApps $excluded @options }
                    [void]$builder.AppendLine($xml)
                    [void]$builder.AppendLine('')
                }
                Show-PreviewDialog -Owner $owner -Title 'M365 Preview (Install XML)' -Content $builder.ToString() -Width 820 -Height 640
            }
            catch {
                Show-PreviewDialog -Owner $owner -Title 'M365 Preview (Error)' -Content ("Failed to build preview:`r`n{0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) -Width 600 -Height 260
            }
        }.GetNewClosure())

    # The languages field accepts culture codes only.
    $Page.Group.ExtraValidate.Add({
            param($context)
            if ($useBox.IsEnabled -and $useBox.IsChecked -eq $true) { return $null }
            $codes = @(([string](Get-OptionsLiveText -Context $context -FieldId 'm365.languages')) -split '[,;\s]+' | Where-Object { $_ })
            $bad = @($codes | Where-Object { $_ -notmatch '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){1,2}$' })
            if ($bad.Count -gt 0) { return @{ FieldId = 'm365.languages'; Message = ('Languages must be culture codes such as en-us or de-de. Not valid: {0}.' -f ($bad -join ', ')) } }
            return $null
        }.GetNewClosure()) | Out-Null
}

# =============================================================================
# Options window: tree, search, OK and Cancel
# -----------------------------------------------------------------------------
# The tree is the list of sections. A section is built the first time it is
# shown: building every section before the window appears delays the first open,
# and a section that nobody visits has no changes to save. OK validates and
# commits the sections that were built, then saves once.
# =============================================================================
function New-OptionsDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$InitialSection = 'cm'
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="1260" Height="760"
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="290"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <DockPanel Grid.Column="0" Grid.Row="0" Margin="0,10,0,0">
            <TextBox x:Name="txtSearch" DockPanel.Dock="Top" Margin="12,0,12,8" FontSize="13"
                     Controls:TextBoxHelper.Watermark="Search settings"
                     Controls:TextBoxHelper.ClearTextButton="True"
                     ToolTip="Type the name of a setting; Enter opens the first match."/>
            <Grid>
                <TreeView x:Name="treeNav" BorderThickness="0" FontSize="13" ItemContainerStyle="{DynamicResource ThemedTreeViewItemStyle}"/>
                <ListBox x:Name="lstSearch" BorderThickness="0" FontSize="13" Visibility="Collapsed" ItemContainerStyle="{DynamicResource ThemedListBoxItemStyle}"/>
            </Grid>
        </DockPanel>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <ContentControl Grid.Column="2" Grid.Row="0" x:Name="contentArea" Margin="24,18,12,12"/>

        <Border Grid.Column="0" Grid.ColumnSpan="3" Grid.Row="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="20,12,20,12">
                <Button x:Name="btnOK"     Content="OK" Margin="0,0,8,0" IsDefault="True" Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnCancel" Content="Cancel" IsCancel="True" Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = Expand-SharedStyle $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $treeNav = $dlg.FindName('treeNav')
    $lstSearch = $dlg.FindName('lstSearch')
    $txtSearch = $dlg.FindName('txtSearch')
    $contentArea = $dlg.FindName('contentArea')
    $btnOK = $dlg.FindName('btnOK')
    $btnCancel = $dlg.FindName('btnCancel')

    $ctx = New-OptionsContext -Prefs $script:Prefs
    $ctx.Owner = $dlg
    $ctx.FieldTable = Get-OptionsFieldTable
    $definition = @(Get-OptionsSectionDefinition)
    $byId = @{}
    foreach ($d in $definition) { $byId[[string]$d.Id] = $d }
    $index = Get-OptionsSearchIndex -Definition $definition

    $nodes = @{}
    foreach ($d in $definition) {
        $item = New-Object System.Windows.Controls.TreeViewItem
        $item.Header = [string]$d.Title
        $item.Tag = [string]$d.Id
        $item.IsExpanded = $true
        $nodes[[string]$d.Id] = $item
        if ($d.Parent) { [void]$nodes[[string]$d.Parent].Items.Add($item) } else { [void]$treeNav.Items.Add($item) }
    }

    $showSection = {
        param([string]$Id)
        $def = $byId[$Id]
        if (-not $def) { return }
        if ($def.Kind -eq 'Group') {
            if ($nodes[$Id].Items.Count -gt 0) { $nodes[$Id].Items[0].IsSelected = $true }
            return
        }
        if (-not $ctx.Sections.Contains($Id)) {
            try { $ctx.Sections[$Id] = New-OptionsSectionPage -Definition $def -Context $ctx }
            catch {
                Add-LogLine -Message ('Options section {0} could not be built: {1}' -f $def.Title, (Get-UserErrorText -ErrorRecord $_ -Fallback 'See the log file.'))
                $ctx.Sections[$Id] = @{ Id = $Id; Element = (New-OptionsTextBlock -Text 'This section could not be built. See the log file.'); Editors = @{}; Commit = $null; Validate = $null }
            }
        }
        $contentArea.Content = $ctx.Sections[$Id].Element
        $dlg.Title = ('Options: ' + (Get-OptionsSectionPath -Definition $definition -Id $Id))
    }.GetNewClosure()

    $treeNav.Add_SelectedItemChanged({
            $item = $treeNav.SelectedItem
            if ($item) { & $showSection ([string]$item.Tag) }
        }.GetNewClosure())

    $navigate = {
        param([string]$SectionId, [string]$FieldId)
        $node = $nodes[$SectionId]
        if (-not $node) { return }
        $parent = $node.Parent
        while ($parent -is [System.Windows.Controls.TreeViewItem]) { $parent.IsExpanded = $true; $parent = $parent.Parent }
        if ($node.IsSelected) { & $showSection $SectionId } else { $node.IsSelected = $true }
        $node.BringIntoView()
        if ($FieldId) {
            $page = $ctx.Sections[$SectionId]
            $editor = $(if ($page -and $page.Editors -and $page.Editors.Contains($FieldId)) { $page.Editors[$FieldId] } else { $null })
            if ($editor -and $editor.Focus) {
                $target = $editor.Focus
                $focus = { $target.BringIntoView(); [void]$target.Focus() }.GetNewClosure()
                & $focus
                [void]$dlg.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]$focus)
            }
        }
    }.GetNewClosure()

    $showTree = { $lstSearch.Visibility = [System.Windows.Visibility]::Collapsed; $treeNav.Visibility = [System.Windows.Visibility]::Visible }.GetNewClosure()
    $jump = {
        param($Entry)
        if (-not $Entry) { return }
        & $showTree
        & $navigate ([string]$Entry.SectionId) ([string]$Entry.FieldId)
    }.GetNewClosure()

    $txtSearch.Add_TextChanged({
            $query = [string]$txtSearch.Text
            if ([string]::IsNullOrWhiteSpace($query)) { & $showTree; return }
            $lstSearch.Items.Clear()
            $hits = @(Search-OptionsIndex -Index $index -Query $query)
            foreach ($hit in $hits) {
                $li = New-Object System.Windows.Controls.ListBoxItem
                $li.Content = ('{0}   -   {1}' -f $hit.Title, $hit.Path)
                $li.Tag = $hit
                [void]$lstSearch.Items.Add($li)
            }
            if ($hits.Count -eq 0) {
                $none = New-Object System.Windows.Controls.ListBoxItem
                $none.Content = 'No setting found.'
                $none.IsEnabled = $false
                [void]$lstSearch.Items.Add($none)
            }
            $treeNav.Visibility = [System.Windows.Visibility]::Collapsed
            $lstSearch.Visibility = [System.Windows.Visibility]::Visible
        }.GetNewClosure())

    $txtSearch.Add_PreviewKeyDown({
            param($sender, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Enter -and $lstSearch.Items.Count -gt 0) {
                $first = $lstSearch.Items[0]
                if ($first.Tag) { & $jump $first.Tag; $e.Handled = $true }
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Down -and $lstSearch.Visibility -eq [System.Windows.Visibility]::Visible -and $lstSearch.Items.Count -gt 0) {
                $lstSearch.SelectedIndex = 0
                [void]$lstSearch.Focus()
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Escape -and $txtSearch.Text) {
                $txtSearch.Text = ''
                $e.Handled = $true
            }
        }.GetNewClosure())

    $lstSearch.Add_PreviewMouseLeftButtonUp({
            $selected = $lstSearch.SelectedItem
            if ($selected -and $selected.Tag) { & $jump $selected.Tag }
        }.GetNewClosure())
    $lstSearch.Add_KeyDown({
            param($sender, $e)
            $selected = $lstSearch.SelectedItem
            if ($e.Key -eq [System.Windows.Input.Key]::Enter -and $selected -and $selected.Tag) { & $jump $selected.Tag; $e.Handled = $true }
        }.GetNewClosure())

    $result = @{ Saved = $false }
    $apply = {
        # The sections that were built hold the changes. A value that cannot be stored stops the save and opens its field.
        foreach ($id in @($ctx.Sections.Keys)) {
            $page = $ctx.Sections[$id]
            if (-not $page -or -not $page.Validate) { continue }
            $problem = & $page.Validate
            if ($problem) {
                & $navigate $id ([string]$problem.FieldId)
                [void](Show-ThemedMessage -Owner $dlg -Title 'Check the Value' -Message ([string]$problem.Message) -Buttons OK -Icon Warning)
                return $false
            }
        }
        try {
            $siteBefore = Get-SiteIdentityKey -SiteCode $ctx.Prefs.SiteCode -ProviderMachineName $ctx.Prefs.ProviderMachineName
            foreach ($id in @($ctx.Sections.Keys)) { $page = $ctx.Sections[$id]; if ($page -and $page.Commit) { & $page.Commit } }
            $siteChanged = (Get-SiteIdentityKey -SiteCode $ctx.Prefs.SiteCode -ProviderMachineName $ctx.Prefs.ProviderMachineName) -cne $siteBefore
            Save-Preferences -Prefs $ctx.Prefs
            if ($ctx.Touched['cwa']) { Save-CwaSwitches -Switches $ctx.Aux.cwa }
            if ($ctx.Touched['tv']) { Save-TvHostConfig -Config $ctx.Aux.tv }
            Invoke-RefreshGrid -DiscardSiteResults:$siteChanged
            Update-SidebarForSystems
            $result.Saved = $true
            return $true
        }
        catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
            return $false
        }
    }.GetNewClosure()

    $btnOK.Add_Click({ if (& $apply) { $dlg.Close() } }.GetNewClosure())
    $btnCancel.Add_Click({ $dlg.Close() }.GetNewClosure())

    $start = $(if ($byId.ContainsKey($InitialSection)) { $InitialSection } else { 'cm' })
    & $navigate $start ''

    return @{ Window = $dlg; Context = $ctx; Definition = $definition; Index = $index; Nodes = $nodes; Navigate = $navigate; Jump = $jump; Apply = $apply; Result = $result; SearchBox = $txtSearch; SearchList = $lstSearch; Tree = $treeNav }
}

function Show-OptionsDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$InitialSection = 'cm'
    )
    $dialog = New-OptionsDialog -Owner $Owner -InitialSection $InitialSection
    [void]$dialog.Window.ShowDialog()
    if ($dialog.Result.Saved) { Add-LogLine -Message 'Options saved.' }
}

# =============================================================================
# First-run setup wizard - the systems and the minimum settings of each system.
# The fields are the Connections fields of the Options window, built by the
# same code, so a value entered here is the value that Options shows.
# =============================================================================
function Show-FirstRunWizard {
    param([Parameter(Mandatory)]$Owner)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Setup"
    Width="760" Height="640"
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ScrollViewer Grid.Row="0" Margin="20,18,20,12" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="stkWizard">
                <TextBlock Text="Welcome to AppPackager" FontSize="20" Margin="0,0,0,6"/>
                <TextBlock TextWrapping="Wrap" FontSize="12" Margin="0,0,0,12"
                           Text="Select the systems you publish to, then fill in their settings. The sidebar has a publish button for each system. You can change them later in Options."/>
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
                          ToolTip="Do not show this wizard again."/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnWizSave" Content="Save" Margin="0,0,8,0" IsDefault="True" Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <Button x:Name="btnWizSkip" Content="Skip" IsCancel="True" Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = Expand-SharedStyle $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $stack = $dlg.FindName('stkWizard')
    $chkWizDontAsk = $dlg.FindName('chkWizDontAsk')
    $btnWizSave = $dlg.FindName('btnWizSave')
    $btnWizSkip = $dlg.FindName('btnWizSkip')

    $ctx = New-OptionsContext -Prefs $script:Prefs
    $ctx.Owner = $dlg
    $table = Get-OptionsFieldTable
    $newBlock = {
        param([string]$Heading, [string[]]$FieldId)
        $block = New-Object System.Windows.Controls.StackPanel
        $block.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        if ($Heading) {
            $title = New-OptionsTextBlock -Text $Heading -Size 15
            $title.Margin = New-Object System.Windows.Thickness(0, 6, 0, 6)
            [void]$block.Children.Add($title)
        }
        $group = New-OptionsFieldGroup -Context $ctx -Field @($FieldId | ForEach-Object { $table[$_] })
        [void]$block.Children.Add($group.Element)
        [void]$stack.Children.Add($block)
        return @{ Block = $block; Group = $group }
    }
    $systemBlock = & $newBlock 'Systems' @('sys.configmgr', 'sys.intune', 'sys.wsus')
    $mecm = & $newBlock 'ConfigMgr' @('cm.siteCode', 'cm.provider', 'cm.fileShare', 'pkg.downloadRoot')
    $intune = & $newBlock 'Intune' @('int.tenant', 'int.client', 'int.secret')
    $wsus = & $newBlock 'WSUS' @('wsus.server', 'wsus.port', 'wsus.ssl')

    $mecmBox = $systemBlock.Group.Editors['sys.configmgr'].Focus
    $intuneBox = $systemBlock.Group.Editors['sys.intune'].Focus
    $wsusBox = $systemBlock.Group.Editors['sys.wsus'].Focus
    $updateGroupState = {
        $mecm.Block.Visibility = $(if ($mecmBox.IsChecked -eq $true) { 'Visible' } else { 'Collapsed' })
        $intune.Block.Visibility = $(if ($intuneBox.IsChecked -eq $true) { 'Visible' } else { 'Collapsed' })
        $wsus.Block.Visibility = $(if ($wsusBox.IsChecked -eq $true) { 'Visible' } else { 'Collapsed' })
    }.GetNewClosure()
    & $updateGroupState
    foreach ($box in @($mecmBox, $intuneBox, $wsusBox)) { $box.Add_Click($updateGroupState) }

    $script:FirstRunDlgSaved = $false
    $prefsRef = $script:Prefs
    # ShowDialog keeps this function on the stack while these handlers run.
    # Keep its scope: GetNewClosure would hide the script-local save/refresh
    # helpers and give each handler a separate $script:FirstRunDlgSaved flag.
    $btnWizSave.Add_Click({
        try {
            foreach ($part in @($systemBlock, $mecm, $intune, $wsus)) {
                $problem = & $part.Group.Validate
                if ($problem -and $part.Block.Visibility -eq 'Visible') {
                    [void](Show-ThemedMessage -Owner $dlg -Title 'Check the Value' -Message ([string]$problem.Message) -Buttons OK -Icon Warning)
                    return
                }
            }
            & $systemBlock.Group.Commit
            $useMecm = ($mecmBox.IsChecked -eq $true)
            $useIntune = ($intuneBox.IsChecked -eq $true)
            $useWsus = ($wsusBox.IsChecked -eq $true)
            if ($useMecm) { & $mecm.Group.Commit }
            # The built-in default would otherwise persist as a site code the operator never entered.
            else { $prefsRef.SiteCode = '' }
            if ($useIntune) { & $intune.Group.Commit }
            if ($useWsus) { & $wsus.Group.Commit }
            # One Click runs unattended: every selected system becomes a default
            # destination; the sidebar buttons choose per run.
            $prefsRef.AppFlow.DefaultDestinations = [pscustomobject]@{ ConfigMgr = $useMecm; WSUS = $useWsus; Intune = $useIntune }
            $target = ConvertTo-DeploymentTargetName -Destinations $prefsRef.AppFlow.DefaultDestinations
            $prefsRef.Intune.DeploymentTarget = $target
            $prefsRef.Intune.PublishToIntune  = $useIntune
            $prefsRef.FirstRunCompleted       = $true
            Save-Preferences -Prefs $prefsRef
            Invoke-RefreshGrid
            Update-SidebarForSystems
            $script:FirstRunDlgSaved = $true
            $dlg.Close()
        } catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
        }
    })

    # Skip and window close share one path: the flag persists only when the
    # suppression box is checked, otherwise the wizard returns next launch.
    $dlg.Add_Closing({
        if ($script:FirstRunDlgSaved) { return }
        if ([bool]$chkWizDontAsk.IsChecked) {
            $prefsRef.FirstRunCompleted = $true
            try { Save-Preferences -Prefs $prefsRef } catch { <# Saving the first-run flag while the window closes is best-effort. #> $null = $_ }
        }
    })

    $btnWizSkip.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog()

    if ($script:FirstRunDlgSaved) {
        $defaultSet = ConvertTo-OneClickDestinationSet -Value $script:Prefs.AppFlow.DefaultDestinations
        $defaultNames = @(foreach ($d in (Get-OneClickDestinationNames)) { if ($defaultSet.$d) { $d } })
        Add-LogLine -Message ("Setup complete. One Click publishes to {0}; change it in Options, Updates and CVEs, One Click." -f $(if ($defaultNames.Count) { $defaultNames -join ' and ' } else { 'no destination' }))
    }
}

# =============================================================================
# Grid refresh helper
# =============================================================================
function Get-SiteIdentityKey {
    param([AllowNull()][string]$SiteCode, [AllowNull()][string]$ProviderMachineName)
    # Whitespace or case differences in saved values name the same site; a
    # raw string compare reports a site change and discards ConfigMgr results.
    return ('{0}|{1}' -f ([string]$SiteCode).Trim(), ([string]$ProviderMachineName).Trim()).ToUpperInvariant()
}

function Invoke-RefreshGrid {
    # ConfigMgr versions and compare results describe one site; a site change
    # must not carry them onto rows that now point elsewhere.
    param([switch]$DiscardSiteResults)

    # Row objects are reused across rebuilds: a background pipeline holds
    # references to them, and a replaced row loses the results written later.
    # The cache also keeps results for rows hidden by the product filter.
    if ($null -eq $script:GridSessionRows) { $script:GridSessionRows = @{} }
    foreach ($row in @($script:PackagerData)) { $script:GridSessionRows[[string]$row.Script] = $row }
    $script:PackagerData.Clear()

    $items = Get-Packagers -Root $PackagersRoot
    $hidden = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($script:Prefs.HiddenApplications),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    # Pre-load persistent history so first render shows Latest + LastChecked
    # from prior sessions (Full Run + manual Check Latest both write here).
    $history = @{}
    try { $history = Read-PackagerHistory } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }

    $present = @{}
    $hiddenCount = 0
    foreach ($m in $items) {
        $present[[string]$m.Script] = $true

        $baseName     = [System.IO.Path]::GetFileNameWithoutExtension($m.Script)
        $latestStored = ""
        $lastChecked  = ""
        $cmStored     = ""
        $vulnStored   = $null
        if ($history.ContainsKey($baseName)) {
            $entry = $history[$baseName]
            $field = { param($name) if ($entry -is [hashtable]) { $entry[$name] } elseif ($entry.PSObject.Properties[$name]) { $entry.PSObject.Properties[$name].Value } else { $null } }
            if (& $field 'LastKnownVersion') { $latestStored = [string](& $field 'LastKnownVersion') }
            if (& $field 'LastChecked')      { $lastChecked  = [string](& $field 'LastChecked') }
            # The ConfigMgr version from an earlier session belongs to the
            # site it was read from.
            if ((& $field 'LastCmVersion') -and [string](& $field 'LastCmSite') -eq [string]$script:Prefs.SiteCode) { $cmStored = [string](& $field 'LastCmVersion') }
            $vulnStored = & $field 'VulnSummary'
        }
        $cve = Get-CveGridCells -Summary $vulnStored -Cpe $(if ($m.PSObject.Properties['CPE']) { [string]$m.CPE } else { '' }) -CurrentVersion $cmStored -LatestVersion $latestStored
        $advisory = $(if ($m.PSObject.Properties['SecurityAdvisoryUrl']) { [string]$m.SecurityAdvisoryUrl } else { '' })
        # Both versions from the history give the row its compare result at
        # once; a stored CVE count for the same pair rides on the status.
        $preloadStatus = [string]$m.Status
        if ($cmStored -and $latestStored -and $preloadStatus -eq 'Ready') {
            $cmp = Compare-SemVer -A $cmStored -B $latestStored
            if ($cmp -lt 0) { $preloadStatus = 'Update available'; if ($cve.Text) { $preloadStatus += ' (' + $cve.Text + ')' } }
            elseif ($cmp -eq 0) { $preloadStatus = 'Up to date' }
            else { $preloadStatus = 'Current newer' }
        }

        $row = $script:GridSessionRows[[string]$m.Script]
        if ($row) {
            $row.Vendor      = $m.Vendor
            $row.Application = $m.Application
            $row.CMName      = $m.CMName
            $row.FullPath    = $m.FullPath
            $row.VendorURL   = $m.VendorUrl
            $row.Description = $m.Description
            $row.Cpe         = $cve.Cpe
            $row.AdvisoryUrl = $advisory
            if (-not [string]$row.LatestVersion) {
                $row.LatestVersion = $latestStored
                Clear-StaleCveCells -Row $row
            }
            if (-not [string]$row.LastChecked)   { $row.LastChecked   = $lastChecked }
            if (-not [string]$row.CveChecked -and -not [string]$row.CveCount) {
                $row.CveCount = $cve.Count; $row.CveCvss = $cve.Cvss; $row.CveKev = $cve.Kev; $row.CveEpss = $cve.Epss; $row.CveChecked = $cve.Checked
                $row.CveFixed = $cve.Fixed; $row.CveNewest = $cve.Newest; $row.CveUpdated = $cve.Updated
                Set-CveCellsPair -Row $row -Pair $(if ($cve.Checked) { ConvertTo-CvePairKey -Current $cmStored -Latest $latestStored } else { '' })
                Clear-StaleCveCells -Row $row
            }
            if ($DiscardSiteResults -or ([string]$m.Status).StartsWith('Read error')) {
                $row.CurrentVersion = ""
                Clear-StaleCveCells -Row $row
                $row.Status         = $m.Status
                $row.CveCount = $cve.Count; $row.CveCvss = ''; $row.CveKev = ''; $row.CveEpss = ''; $row.CveChecked = ''
                $row.CveFixed = ''; $row.CveNewest = ''; $row.CveUpdated = ''
                if ($cve.Checked) { $row.CveCount = '' }
            }
        }
        else {
            $row = [pscustomobject]@{
                Selected       = $false
                Vendor         = $m.Vendor
                Application    = $m.Application
                CurrentVersion = $cmStored
                LatestVersion  = $latestStored
                Status         = $preloadStatus
                CMName         = $m.CMName
                Script         = $m.Script
                FullPath       = $m.FullPath
                VendorURL      = $m.VendorUrl
                Description    = $m.Description
                LastChecked    = $lastChecked
                Cpe            = $cve.Cpe
                CveCount       = $cve.Count
                CveCvss        = $cve.Cvss
                CveKev         = $cve.Kev
                CveEpss        = $cve.Epss
                CveChecked     = $cve.Checked
                CveFixed       = $cve.Fixed
                CveNewest      = $cve.Newest
                CveUpdated     = $cve.Updated
                CvePair        = $(if ($cve.Checked) { ConvertTo-CvePairKey -Current $cmStored -Latest $latestStored } else { '' })
                AdvisoryUrl    = $advisory
            }
            $script:GridSessionRows[[string]$m.Script] = $row
        }

        if ($hidden.Contains($m.Script)) {
            $row.Selected = $false
            $hiddenCount++
            continue
        }
        $script:PackagerData.Add($row)
    }

    foreach ($key in @($script:GridSessionRows.Keys)) {
        if (-not $present.ContainsKey($key)) { $script:GridSessionRows.Remove($key) }
    }

    if ($hiddenCount -gt 0) {
        $txtStatus.Text = ("{0} packager(s) loaded, {1} hidden. Ready." -f $script:PackagerData.Count, $hiddenCount)
    }
    else {
        $txtStatus.Text = ("Loaded {0} packager(s). Ready." -f $script:PackagerData.Count)
    }

    # A filtered ItemsSource is a snapshot array; it must be recomputed after
    # the collection is repopulated.
    Update-GridFilter
}

function Clear-SucceededSelection {
    # Rows whose Stage or Package step succeeded are unchecked; failed and
    # skipped rows keep their check so a rerun targets only the remainder.
    param(
        [AllowNull()][object[]]$Rows,
        [AllowNull()][string[]]$Scripts
    )
    if (-not $Rows -or -not $Scripts) { return }
    $done = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Scripts), [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in @($Rows)) {
        if ($row -and $done.Contains([string]$row.Script)) { $row.Selected = $false }
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
$script:BgGraveyard = @()

function Import-BackgroundWorkerModule {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.Runspace]$Runspace,
        [Parameter(Mandatory)][string]$Root
    )
    # Every module whose commands the background loop or the inline helpers
    # it calls resolve: Common for history and the ConfigMgr application
    # helpers, Workbench for One Click build records, Discovery for the
    # ConfigMgr title rule of Get-MecmCurrentVersionByCMName. A missing
    # module surfaces only as "The term ... is not recognized" inside a run.
    $modulePath = @(
        (Join-Path $Root 'Packagers\AppPackagerCommon.psm1')
        (Join-Path $Root 'Packagers\AppPackagerWorkbench.psm1')
        (Join-Path $Root 'Packagers\AppPackagerSigning.psm1')
        (Join-Path $Root 'Packagers\AppPackagerWsus.psm1')
        (Join-Path $Root 'Packagers\AppPackagerOneClick.psm1')
        (Join-Path $Root 'Packagers\AppPackagerVulnerability.psm1')
        (Join-Path $Root 'Packagers\AppPackagerDiscovery.psm1')
    )
    $initPS = [powershell]::Create()
    $initPS.Runspace = $Runspace
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
}

function Initialize-BackgroundWorker {
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }

    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()

    Import-BackgroundWorkerModule -Runspace $script:BgRunspace -Root $PSScriptRoot

    # The runspace writes to the same log file as the window.
    if ($script:GuiLogPath) {
        $logPS = [powershell]::Create()
        $logPS.Runspace = $script:BgRunspace
        [void]$logPS.AddScript({ param($Path) Initialize-Logging -LogPath $Path -Attach }).AddArgument($script:GuiLogPath)
        try { [void]$logPS.Invoke() } catch { $null = $_ } finally { $logPS.Dispose() }
    }

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

function Test-PipelineBusy {
    # True while a background run is in flight. A function keeps the script
    # session state, so $script: resolves here where a closure cannot.
    # The state stays set until the completion tick has run every hook and
    # cleared it; Done alone leaves a gap in which a new run would drop
    # the pending hooks. A discovery scan shows the same display, so it counts.
    [bool]($null -ne $script:BgState) -or [bool]$script:CmImportRunning -or [bool]($null -ne $script:DiscoveryJob)
}

function Invoke-MultiAppPipeline {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('CheckLatest','CheckCve','Stage','Package','FullRun','OneClick')]
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
        Succeeded       = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
        Counts          = $null
        # Existing-application overwrite handshake. ConflictDecisionForAll
        # lives for this run only; nothing about it is persisted.
        ConflictRequest        = $null
        ConflictResponse       = $null
        ConflictDecisionForAll = ''
        # One Click progress per application, read by the plan window.
        PlanRows               = $(if ($Context.ContainsKey('OneClickRows')) { $Context.OneClickRows } else { $null })
        # Check CVEs results by script name, read by the CVE window's
        # completion hook (a cache write can fail after a good lookup).
        CveResults             = [hashtable]::Synchronized(@{})
    })
    $script:ConflictPromptOpen = $false
    # The timer tick runs after this function returns; its parameters are out
    # of scope there.
    $script:BgContext   = $Context
    $script:BgOperation = $Operation

    Set-ActionButtonsEnabled -Enabled $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait

    $titleMap = @{
        CheckLatest = 'Checking latest versions'
        CheckCve    = 'Checking CVEs'
        Stage       = 'Staging packages'
        Package     = 'Packaging applications'
        FullRun     = 'One Click flow'
        OneClick    = 'One Click'
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
            Checked = 0; Updated = 0; Reported = 0; Staged = 0; Packaged = 0; StageAndPackage = 0; Published = 0
            NoChange = 0; Skipped = 0; NotSupported = 0; Failed = 0; CheckFailed = 0
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
                            Clear-StaleCveCells -Row $row

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
                                [void]$State.LogQueue.Enqueue(('History write failed for {0}: {1}' -f $baseName, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            }
                            $counts['Checked']++
                            if ($histResult -eq 'NoChange') { $counts['NoChange']++ }
                            else { $counts['Updated']++ }
                        } catch {
                            $row.Status = 'Error'
                            [void]$State.LogQueue.Enqueue(('Error: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            try { Update-PackagerHistory -PackagerName $baseName -Event Checked -Result Failed } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                            $counts['CheckFailed']++
                        }
                    }

                    'CheckCve' {
                        $State.Step = ('CVEs {0}/{1}: {2}' -f $i, $n, $app)
                        [void]$State.LogQueue.Enqueue(('CVEs: {0} ({1})' -f $app, $scrName))
                        $cpe     = [string]$row.Cpe
                        $current = [string]$row.CurrentVersion
                        $latest  = [string]$row.LatestVersion
                        $reason = Get-CveSkipReason -Cpe $cpe -Current $current -Latest $latest
                        if ($reason) {
                            $row.CveCount = $reason
                            $row.CveCvss = ''; $row.CveKev = ''; $row.CveEpss = ''; $row.CveChecked = ''; $row.CveFixed = ''; $row.CveNewest = ''; $row.CveUpdated = ''
                            Set-CveCellsPair -Row $row
                            [void]$State.LogQueue.Enqueue(('Skipped: {0}.' -f $reason))
                            $counts['Skipped']++
                            continue
                        }
                        try {
                            if ($null -eq $Ctx.KevCatalog) { $Ctx.KevCatalog = Get-KevCatalog }
                            $res = Get-ProductVulnerabilities -Cpe $cpe -CurrentVersion $current -LatestVersion $latest `
                                -ApiKey ([string]$Ctx.NvdApiKey) -CacheTtlMinutes ([int]$Ctx.CacheTtlMinutes) -KevCatalog $Ctx.KevCatalog -Force:([bool]$Ctx.ForceRefresh) `
                                -VendorAdvisoryUrl $(if ($row.PSObject.Properties['AdvisoryUrl']) { [string]$row.AdvisoryUrl } else { '' })
                            $State.CveResults[$scrName] = $res
                            if ($Ctx.CveSession -and -not $res.Error) { $Ctx.CveSession[$scrName] = $res }
                            $summary = Get-VulnerabilitySummary -Result $res
                            $cells = Get-CveGridCells -Summary $summary -Cpe $cpe
                            $row.CveCount = $cells.Count; $row.CveCvss = $cells.Cvss; $row.CveKev = $cells.Kev; $row.CveEpss = $cells.Epss; $row.CveChecked = $cells.Checked
                            $row.CveFixed = $cells.Fixed; $row.CveNewest = $cells.Newest; $row.CveUpdated = $cells.Updated
                            Set-CveCellsPair -Row $row
                            if ($res.Error) {
                                # A failed lookup is shown, not recorded: the
                                # history keeps the last good summary.
                                [void]$State.LogQueue.Enqueue(('Error: ' + $res.Error))
                                $counts['CheckFailed']++
                            }
                            else {
                                if ([string]$row.Status -match '^Update available') { $row.Status = 'Update available (' + $cells.Text + ')' }
                                [void]$State.LogQueue.Enqueue(('{0} Source {1}{2}' -f (ConvertTo-VulnerabilityResultText -Result $res), $res.Source, $(if ($null -ne $res.MaxCvss) { '; max CVSS ' + (ConvertTo-VulnerabilityCvssText -Score $res.MaxCvss -Severity $res.MaxSeverity) } else { '' })))
                                $counts['Checked']++
                                try { Update-PackagerHistory -PackagerName $baseName -Event VulnChecked -Data $summary } catch {
                                    [void]$State.LogQueue.Enqueue(('History write failed for {0}: {1}' -f $baseName, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                                }
                            }
                        } catch {
                            $row.CveCount = 'Error'
                            [void]$State.LogQueue.Enqueue(('Error: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
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
                                if (-not [string]::IsNullOrWhiteSpace([string]$Ctx.IntuneWinToolPath)) {
                                    $stageIntuneWin = Invoke-PackagerIntuneWinPostStep -Result $res -PackagerPath $path -DownloadRoot $planDownloadRoot -ToolPath ([string]$Ctx.IntuneWinToolPath) -SkipNetworkCopy
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $stageIntuneWin.Message))
                                }
                                $ver = [string]$row.LatestVersion
                                try {
                                    if ($ver) { Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $ver -Result Updated }
                                    else      { Update-PackagerHistory -PackagerName $baseName -Event Staged -Result Updated }
                                } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                [void]$State.Succeeded.Enqueue($scrName)
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
                            [void]$State.LogQueue.Enqueue(('Stage exception: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
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
                                WsusPublishConfig    = $Ctx.WsusPublishConfig
                                DeploymentTarget     = $(if ([string]$Ctx.DeploymentTarget -in (Get-DeploymentTargetNames)) { [string]$Ctx.DeploymentTarget } else { 'MECM' })
                                ContentLayout        = [string]$Ctx.ContentLayout
                                RequirementsJson     = $reqJson
                                VariantsJson         = $varJson
                                CommandsJson         = $cmdJson
                                InstallMode          = $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' })
                            }
                            $packageArgs['TitleMode'] = $(if ($Ctx.TitleModesByApp -and [string]$Ctx.TitleModesByApp[$baseName]) { [string]$Ctx.TitleModesByApp[$baseName] } else { [string]$Ctx.DefaultTitleMode })
                            $res = Invoke-PackagerPackageWithConflictPrompt -State $State -AppLabel $app -PackageArgs $packageArgs
                            if ($res.PSObject.Properties['WsusPublish'] -and $res.WsusPublish) {
                                [void]$State.LogQueue.Enqueue(('WSUS publish: ' + $res.WsusPublish.Message))
                            }
                            if ($res.ExitCode -eq 0) {
                                if ($res.PSObject.Properties['IntunePublish'] -and $res.IntunePublish) {
                                    [void]$State.LogQueue.Enqueue(('Intune publish: ' + $res.IntunePublish.Message))
                                }
                                if ($res.PSObject.Properties['IntuneWin'] -and $res.IntuneWin) {
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $res.IntuneWin.Message))
                                    if ($res.IntuneWin.NetworkPath) {
                                        [void]$State.LogQueue.Enqueue(('Intunewin on network: ' + $res.IntuneWin.NetworkPath))
                                    }
                                }
                            }
                            $outcome = Get-PublishRowOutcome -Result $res -DeploymentTarget ([string]$packageArgs.DeploymentTarget)
                            $row.Status = $outcome.Status

                            if ($outcome.Kind -eq 'Skipped') {
                                $counts['Skipped']++
                            } elseif ($outcome.Kind -eq 'NotSupported') {
                                $counts['NotSupported']++
                            } elseif ($outcome.Kind -eq 'Success') {
                                [void]$State.LogQueue.Enqueue(('{0}. Logs: {1}' -f $outcome.Status, (Split-Path -Leaf $res.OutLog)))
                                $ver = [string]$row.LatestVersion
                                try {
                                    if ($ver) { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $ver -Result Updated }
                                    else      { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Result Updated }
                                } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                [void]$State.Succeeded.Enqueue($scrName)
                                $counts[$(if ($outcome.Published) { 'Published' } else { 'Packaged' })]++
                            } elseif ($res.ExitCode -eq 0) {
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $res.OutLog)))
                                $counts['Failed']++
                            } else {
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
                            [void]$State.LogQueue.Enqueue(('Package exception: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            $counts['Failed']++
                        }
                    }

                    'OneClick' {
                        # The plan window decided the rows and their
                        # destinations; nothing here prompts. Each step writes
                        # the row's progress record for the plan grid and the
                        # run report.
                        $pr = $State.PlanRows[$baseName]
                        $setStep = { param($text) $State.Step = ('One Click {0}/{1}: {2}' -f $i, $n, $text); if ($pr) { $pr.Step = $text } }
                        $finish = { param($outcome, $reason) if ($pr) { $pr.Outcome = $outcome; if ($reason) { $pr.Reason = $reason }; $pr.Step = $outcome } }
                        $wanted = @()
                        if ($Ctx.OneClickPlan -and $Ctx.OneClickPlan.ContainsKey($baseName)) { $wanted = @($Ctx.OneClickPlan[$baseName].Destinations) }
                        $onExisting = $(if ([string]$Ctx.OnExisting -eq 'Overwrite') { 'Overwrite' } else { 'Skip' })

                        $hist = @{}
                        try { $hist = Read-PackagerHistory } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                        $entry = $null
                        if ($hist.ContainsKey($baseName)) { $entry = $hist[$baseName] }
                        $lastKnown = $null; $lastStaged = $null
                        if ($entry) {
                            if ($entry -is [hashtable]) { $lastKnown = $entry['LastKnownVersion']; $lastStaged = $entry['LastStaged'] }
                            else { $lastKnown = $entry.LastKnownVersion; $lastStaged = $entry.LastStaged }
                        }

                        # 1. Check latest
                        & $setStep ('checking ' + $app)
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
                            Clear-StaleCveCells -Row $row
                            if ($pr) { $pr.Version = [string]$latest }
                            [void]$State.LogQueue.Enqueue(('Latest: ' + $latest))
                        } catch {
                            $row.Status = 'Check error'
                            [void]$State.LogQueue.Enqueue(('Latest check failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            $counts['CheckFailed']++
                            & $finish 'Failed' ('check failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
                            continue
                        }

                        # A vendor feed can step back to an older release. The
                        # run skips it before the history learns the older
                        # version; the FullRun worker applies the same guard.
                        if (-not $Ctx.ForceFlag -and $lastKnown -and (Compare-SemVer -A ([string]$lastKnown) -B ([string]$latest) -Strict) -gt 0) {
                            $row.Status = 'Vendor older than last seen'
                            [void]$State.LogQueue.Enqueue(('{0}: the vendor reports {1}, older than {2} seen before - skipping' -f $app, $latest, $lastKnown))
                            $counts['Skipped']++
                            & $finish 'Skipped' ('vendor reports {0}, older than {1} seen before' -f $latest, $lastKnown)
                            continue
                        }

                        $versionChanged = (-not $lastKnown) -or ($lastKnown -ne $latest)
                        # The version a stage completed for this root, not the newest one a check saw.
                        $oneClickStageScope = [string]$planDownloadRoot
                        $stagedVersion = Get-PackagerHistoryActedVersion -Entry $entry -Event Staged -Scope $oneClickStageScope
                        $stageStale = (-not $stagedVersion) -or ([string]$stagedVersion -ne [string]$latest)
                        try {
                            Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result $(if ($versionChanged) { 'Updated' } else { 'NoChange' })
                            $row.LastChecked = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                        } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }

                        if ($Ctx.Action -eq 'Report') {
                            $row.Status = if ($versionChanged) { 'Update available' } else { 'Up to date' }
                            if ($versionChanged) { $counts['Reported']++ } else { $counts['NoChange']++ }
                            & $finish 'Checked' $(if ($versionChanged) { 'update available' } else { 'up to date' })
                            continue
                        }

                        # The build identity a publish receipt carries: profile,
                        # revision and signing policy. A receipt for another
                        # build does not count as published.
                        $buildKey = ''
                        if ($planApplicationId) { $buildKey = '{0}|{1}|{2}' -f [string]$planProfileId, [string]$planRevision, [string]$Ctx.SigningDigest }
                        $buildStale = $false
                        if ($planApplicationId) {
                            $current = Test-WorkbenchBuildIsCurrent -ApplicationId $planApplicationId -ProfileId $planProfileId `
                                -Version $latest -ProfileRevision $planRevision -PolicyDigest ([string]$Ctx.SigningDigest)
                            if ($null -ne $current -and -not $current) { $buildStale = $true }
                        }

                        # 2. Which destinations still need this version. The
                        # guard reads the history again: the check above may
                        # have found a version the plan did not know. A stale
                        # build republishes even when the version matches.
                        $publishTo = @()
                        $skipReasons = @()
                        # What the site holds now, read once: the publish record
                        # below says only what this tool published earlier.
                        $observedCm = ''
                        $observedUnknown = $false
                        if ($Ctx.Action -eq 'StageAndPackage' -and 'ConfigMgr' -in $wanted -and -not $Ctx.ForceFlag -and $Ctx.AdminUiFound -and -not [string]::IsNullOrWhiteSpace([string]$row.CMName)) {
                            try {
                                $siteRes = Get-MecmCurrentVersionByCMName -SiteCode $Ctx.SiteCode -ProviderMachineName $Ctx.ProviderMachineName -CMName ([string]$row.CMName)
                                if ($siteRes.Found -and -not [string]::IsNullOrWhiteSpace([string]$siteRes.SoftwareVersion)) {
                                    $observedCm = [string]$siteRes.SoftwareVersion
                                    # The grid and the history learn what the site holds, as the FullRun check does.
                                    $row.CurrentVersion = $observedCm
                                    Clear-StaleCveCells -Row $row
                                    try { Update-PackagerHistory -PackagerName $baseName -Event CmChecked -Version $observedCm -Data @{ SiteCode = [string]$Ctx.SiteCode } } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                }
                            } catch {
                                $observedUnknown = $true
                                [void]$State.LogQueue.Enqueue(('ConfigMgr site version for {0} not read: {1}' -f $app, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            }
                        }
                        if ($Ctx.Action -eq 'StageAndPackage') {
                            foreach ($d in $wanted) {
                                $instance = [string]$(if ($Ctx.InstanceKeys) { $Ctx.InstanceKeys[$d] } else { '' })
                                $rec = Get-OneClickPublishedVersion -HistoryEntry $entry -Destination $d -Instance $instance
                                if (-not $Ctx.ForceFlag -and $rec -and (Compare-SemVer -A ([string]$rec.Version) -B ([string]$latest) -Strict) -gt 0) {
                                    $skipReasons += ('{0}: published {1} is newer than {2}' -f $d, $rec.Version, $latest)
                                    if ($pr) { $pr.('Result' + $d) = ('published {0} is newer than {1}' -f $rec.Version, $latest) }
                                    continue
                                }
                                if ($d -eq 'ConfigMgr') {
                                    # With Overwrite, the package step replaces what the site holds without comparing versions;
                                    # a site that could not be read is not overwritten.
                                    if ($observedUnknown -and $onExisting -eq 'Overwrite' -and -not $Ctx.ForceFlag) {
                                        $skipReasons += 'ConfigMgr: site version could not be read'
                                        if ($pr) { $pr.('Result' + $d) = 'site version could not be read' }
                                        continue
                                    }
                                    $siteSkip = Get-OneClickObservedSkipReason -Destination $d -Observed $observedCm -Latest ([string]$latest) -Force:([bool]$Ctx.ForceFlag)
                                    if ($siteSkip) {
                                        $skipReasons += $siteSkip
                                        if ($pr) { $pr.('Result' + $d) = $siteSkip }
                                        continue
                                    }
                                }
                                $sameBuild = (-not $buildKey) -or (-not $rec) -or (-not $rec.Build) -or ($rec.Build -eq $buildKey)
                                if (-not $Ctx.ForceFlag -and -not $buildStale -and $sameBuild -and $rec -and $rec.Version -eq [string]$latest) {
                                    $skipReasons += ('{0}: {1} already published' -f $d, $latest)
                                    if ($pr) { $pr.('Result' + $d) = ('already published ' + $latest) }
                                    continue
                                }
                                $refused = Get-OneClickNotSupportedVersion -HistoryEntry $entry -Destination $d -Instance $instance
                                if (-not $Ctx.ForceFlag -and $refused -and $refused -eq [string]$latest) {
                                    $skipReasons += ('{0}: not supported' -f $d)
                                    $refusedReason = Get-OneClickNotSupportedReason -HistoryEntry $entry -Destination $d -Instance $instance
                                    if ($pr) { $pr.('Result' + $d) = $(if ($refusedReason -match '^not supported: ') { $refusedReason } else { 'not supported ' + $latest }) }
                                    continue
                                }
                                $publishTo += $d
                            }
                        }
                        $needStage = $Ctx.ForceFlag -or $buildStale -or $stageStale -or ($publishTo.Count -gt 0)
                        if (-not $needStage) {
                            $row.Status = 'Up to date'
                            [void]$State.LogQueue.Enqueue(('No change - skipping: ' + $app))
                            $counts['NoChange']++
                            & $finish 'Skipped' 'up to date'
                            continue
                        }
                        if ($Ctx.Action -eq 'StageAndPackage' -and $publishTo.Count -eq 0 -and -not $stageStale -and -not $buildStale -and -not $Ctx.ForceFlag) {
                            $row.Status = 'Up to date'
                            $counts['NoChange']++
                            & $finish 'Skipped' ($skipReasons -join '; ')
                            continue
                        }

                        # 3. Stage
                        & $setStep ('staging ' + $app)
                        $row.Status = 'Staging...'
                        [void]$State.LogQueue.Enqueue(('Stage: ' + $app))
                        $stg = $null
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
                                try { Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $latest -Result Updated -Scope $oneClickStageScope } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                            } else {
                                $row.Status = 'Stage error'
                                $stderrLines = @($stg.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                foreach ($line in ($stderrLines | Select-Object -First 10)) { [void]$State.LogQueue.Enqueue(('  stderr: ' + $line)) }
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $stg.OutLog)))
                            }
                        } catch {
                            $row.Status = 'Stage error'
                            [void]$State.LogQueue.Enqueue(('Stage exception: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                        }
                        if (-not $stageOk) {
                            $counts['Failed']++
                            & $finish 'Failed' 'stage failed'
                            continue
                        }
                        if ($Ctx.Action -eq 'Stage' -or $publishTo.Count -eq 0) {
                            [void]$State.Succeeded.Enqueue($scrName)
                            $counts['Staged']++
                            & $finish 'Staged' $(if ($skipReasons.Count) { $skipReasons -join '; ' } else { '' })
                            continue
                        }

                        # 4. Publish per destination, ConfigMgr first: it is the
                        # only destination that reads the network copy.
                        $published = 0; $failed = 0; $notSupported = 0
                        $intuneWinNote = $null
                        foreach ($d in $publishTo) {
                            & $setStep ('publishing {0} to {1}' -f $app, $d)
                            $row.Status = ('Publishing to {0}...' -f $d)
                            $resultText = ''
                            $detail = ''
                            $id = ''
                            $ok = $false
                            $present = $false
                            try {
                                switch ($d) {
                                    'ConfigMgr' {
                                        $reqJson = ''; if ($Ctx.RequirementsByApp) { $reqJson = [string]$Ctx.RequirementsByApp[$baseName] }
                                        $varJson = ''; if ($Ctx.VariantsByApp) { $varJson = [string]$Ctx.VariantsByApp[$baseName] }
                                        $cmdJson = ''; if ($Ctx.CommandsByApp) { $cmdJson = [string]$Ctx.CommandsByApp[$baseName] }
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
                                            M365Channel          = $Ctx.M365Channel
                                            M365DeployMode       = $Ctx.M365DeployMode
                                            EstimatedRuntimeMins = $Ctx.EstimatedRuntimeMins
                                            MaximumRuntimeMins   = $Ctx.MaximumRuntimeMins
                                            SevenZipPath         = $Ctx.SevenZipPath
                                            CreateIntuneWin      = [bool]$Ctx.IntuneWinCreate
                                            IntuneWinToolPath    = [string]$Ctx.IntuneWinToolPath
                                            DeploymentTarget     = 'MECM'
                                            ContentLayout        = [string]$Ctx.ContentLayout
                                            RequirementsJson     = $reqJson
                                            VariantsJson         = $varJson
                                            CommandsJson         = $cmdJson
                                            OnExisting           = $onExisting
                                            InstallMode          = $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' })
                                            TitleMode            = $(if ($Ctx.TitleModesByApp -and [string]$Ctx.TitleModesByApp[$baseName]) { [string]$Ctx.TitleModesByApp[$baseName] } else { [string]$Ctx.DefaultTitleMode })
                                        }
                                        $pkg = Invoke-PackagerPackage @packageArgs
                                        $id = Get-PackagerConfigMgrApplicationName -Text $pkg.StdOut
                                        if ($pkg.ExitCode -ne 0) {
                                            $stderrLines = @($pkg.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                                            foreach ($line in ($stderrLines | Select-Object -First 10)) { [void]$State.LogQueue.Enqueue(('  stderr: ' + $line)) }
                                            $resultText = ('failed: exit code {0}; logs {1}' -f $pkg.ExitCode, (Split-Path -Leaf $pkg.OutLog))
                                        }
                                        elseif ($pkg.StdOut -match [regex]::Escape((Get-OnExistingConflictMarker))) {
                                            $resultText = ('skipped: application {0} already exists' -f $id)
                                            $skipReasons += ('ConfigMgr: application exists, policy ' + $onExisting)
                                        }
                                        else {
                                            $ok = $true
                                            $resultText = ('published ' + $id)
                                            if ($pkg.PSObject.Properties['IntuneWin'] -and $pkg.IntuneWin) { $intuneWinNote = $pkg.IntuneWin }
                                        }
                                    }
                                    'WSUS' {
                                        $note = Invoke-PackagerWsusPostStep -Result $stg -PackagerPath $path -DownloadRoot $planDownloadRoot -Config $Ctx.WsusPublishConfig -SkipsSite
                                        [void]$State.LogQueue.Enqueue(('WSUS publish: ' + $note.Message))
                                        if ($note.Ok) {
                                            $ok = $true; $id = [string]$note.PackageId
                                            # A repeat publish of the same version reuses the update
                                            # that is not expired; nothing new reaches WSUS.
                                            if ([string]$note.Outcome -eq 'AlreadyPublished') {
                                                $present = $true; $resultText = ('already on the server ' + $id)
                                                $skipReasons += ('WSUS: {0} already on the server' -f $latest)
                                            }
                                            else { $resultText = ('published ' + $id) }
                                        }
                                        elseif ($note.Refused) {
                                            # The cell carries the code; the log and the report carry the text.
                                            $notSupported++; $skipReasons += 'WSUS: not supported'
                                            $resultText = $(if ($note.Message -match '^(not supported: [A-Za-z]+)') { $Matches[1] } else { [string]$note.Message })
                                            $detail = [string]$note.Message
                                            try {
                                                $h = Read-PackagerHistory
                                                [void](Set-OneClickNotSupported -History $h -PackagerName $baseName -Destination 'WSUS' -Version ([string]$latest) -Reason $resultText `
                                                    -Instance ([string]$(if ($Ctx.InstanceKeys) { $Ctx.InstanceKeys['WSUS'] } else { '' })))
                                                Save-PackagerHistory -History $h
                                            } catch { [void]$State.LogQueue.Enqueue(('History write failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))) }
                                        }
                                        else { $resultText = [string]$note.Message }
                                    }
                                    'Intune' {
                                        if (-not $intuneWinNote -or -not $intuneWinNote.Ok) {
                                            $intuneWinNote = Invoke-PackagerIntuneWinPostStep -Result $stg -PackagerPath $path -DownloadRoot $planDownloadRoot -ToolPath ([string]$Ctx.IntuneWinToolPath) -SkipNetworkCopy
                                            [void]$State.LogQueue.Enqueue(('Intunewin: ' + $intuneWinNote.Message))
                                        }
                                        if (-not $intuneWinNote.Ok) { $resultText = ('intunewin failed: ' + $intuneWinNote.Message) }
                                        elseif (-not $Ctx.IntunePublishConfig) { $resultText = 'Intune credentials not configured' }
                                        else {
                                            $note = Invoke-PackagerIntunePublish -IntuneNote $intuneWinNote -Config $Ctx.IntunePublishConfig -Comment $Ctx.Comment
                                            [void]$State.LogQueue.Enqueue(('Intune publish: ' + $note.Message))
                                            if ($note.Ok) { $ok = $true; $id = [string]$note.AppId; $resultText = ('published ' + $id) }
                                            else { $resultText = [string]$note.Message }
                                        }
                                    }
                                }
                            } catch {
                                $resultText = ('failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
                            }
                            if ($pr) { $pr.('Result' + $d) = $resultText; $pr.('Id' + $d) = $id; $pr.('Detail' + $d) = $detail }
                            if ($ok) {
                                if (-not $present) { $published++ }
                                try {
                                    $h = Read-PackagerHistory
                                    [void](Set-OneClickPublishedVersion -History $h -PackagerName $baseName -Destination $d -Version ([string]$latest) -Id $id `
                                        -Instance ([string]$(if ($Ctx.InstanceKeys) { $Ctx.InstanceKeys[$d] } else { '' })) -Build $buildKey)
                                    Save-PackagerHistory -History $h
                                } catch { [void]$State.LogQueue.Enqueue(('History write failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))) }
                            }
                            elseif ($resultText -notmatch '^(skipped|not supported)') { $failed++ }
                        }

                        if ($published -gt 0) {
                            try { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $latest -Result Updated } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                            [void]$State.Succeeded.Enqueue($scrName)
                        }
                        if ($failed -gt 0) {
                            $row.Status = $(if ($published -gt 0) { 'Published, partial' } else { 'Publish error' })
                            $counts['Failed']++
                            & $finish 'Failed' ($skipReasons -join '; ')
                        }
                        elseif ($published -gt 0) {
                            $row.Status = 'Published'
                            $counts['Published']++
                            & $finish 'Published' ($skipReasons -join '; ')
                        }
                        elseif ($notSupported -gt 0) {
                            $row.Status = 'WSUS: not supported'
                            $counts['NotSupported']++
                            & $finish 'Not supported' ($skipReasons -join '; ')
                        }
                        else {
                            $row.Status = 'Skipped'
                            $counts['Skipped']++
                            & $finish 'Skipped' ($skipReasons -join '; ')
                        }
                    }
                    'FullRun' {
                        # Cadence gate (Report only), ConfigMgr pre-flight, then Stage + optional Package.
                        # Mirrors the UI-thread handler behavior 1:1 so a Full Run here lands the
                        # same history entries and row.Status flips as the old path did.
                        $State.Step = ('One Click {0}/{1}: {2}' -f $i, $n, $app)

                        $lastChecked = $null; $lastKnown = $null; $lastStaged = $null; $lastPackaged = $null
                        $fullEntry = $null
                        try {
                            $hist = Read-PackagerHistory
                            if ($hist.ContainsKey($baseName)) {
                                $h = $hist[$baseName]
                                $fullEntry = $h
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
                        } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }

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
                                } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                            }
                            try {
                                $nextDue = ([datetime]$lastChecked).ToUniversalTime().AddDays($cadenceDays)
                                if ($nextDue -gt (Get-Date).ToUniversalTime()) {
                                    $row.Status = 'Skipped (cadence)'
                                    [void]$State.LogQueue.Enqueue(('Skipped {0} (within {1}d cadence)' -f $app, $cadenceDays))
                                    $counts['Skipped']++
                                    continue
                                }
                            } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
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
                            Clear-StaleCveCells -Row $row
                            [void]$State.LogQueue.Enqueue(('Latest: ' + $latest))
                        } catch {
                            $row.Status = 'Check error'
                            [void]$State.LogQueue.Enqueue(('Latest check failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            $counts['CheckFailed']++
                            continue
                        }

                        # 1a. ConfigMgr pre-flight for Stage/StageAndPackage.
                        # IntuneOnly and WSUSOnly have no site to query; the query
                        # would open a provider connection, so it is skipped before that.
                        if ($Ctx.Action -in @('Stage','StageAndPackage') -and (Test-DeploymentTargetSkipsSite -DeploymentTarget ([string]$Ctx.DeploymentTarget))) {
                            [void]$State.LogQueue.Enqueue(('ConfigMgr pre-flight skipped for {0}: deployment target is {1}.' -f $app, $(if ([string]$Ctx.DeploymentTarget -eq 'WSUSOnly') { 'WSUS only' } else { 'Intune only' })))
                        }
                        elseif ($Ctx.Action -in @('Stage','StageAndPackage')) {
                            $cmName = [string]$row.CMName
                            if (-not [string]::IsNullOrWhiteSpace($cmName) -and $Ctx.AdminUiFound) {
                                try {
                                    $mecmRes = Get-MecmCurrentVersionByCMName -SiteCode $Ctx.SiteCode -ProviderMachineName $Ctx.ProviderMachineName -CMName $cmName
                                    if ($mecmRes.Found -and -not [string]::IsNullOrWhiteSpace([string]$mecmRes.SoftwareVersion)) {
                                        $row.CurrentVersion = [string]$mecmRes.SoftwareVersion
                                        Clear-StaleCveCells -Row $row
                                        try { Update-PackagerHistory -PackagerName $baseName -Event CmChecked -Version ([string]$mecmRes.SoftwareVersion) -Data @{ SiteCode = [string]$Ctx.SiteCode } } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                        $cmp = Compare-SemVer -A ([string]$mecmRes.SoftwareVersion) -B $latest
                                        if ($cmp -eq 0) {
                                            $row.Status = 'Up to date (ConfigMgr)'
                                            [void]$State.LogQueue.Enqueue(('ConfigMgr already has {0} at {1} - skipping' -f $app, $latest))
                                            try { Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result NoChange } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                            $counts['NoChange']++
                                            continue
                                        }
                                        # The site holds a newer release than the vendor now reports;
                                        # acting would publish the older build over it.
                                        if ($cmp -gt 0 -and -not $Ctx.ForceFlag) {
                                            $row.Status = 'Current newer (ConfigMgr)'
                                            [void]$State.LogQueue.Enqueue(('ConfigMgr has {0} at {1}, newer than the vendor''s {2} - skipping' -f $app, [string]$mecmRes.SoftwareVersion, $latest))
                                            $counts['Skipped']++
                                            continue
                                        }
                                    }
                                } catch {
                                    [void]$State.LogQueue.Enqueue(('ConfigMgr pre-flight for {0} failed: {1}' -f $app, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                                }
                            }
                        }

                        # A vendor feed can step back to an older release (Dell's
                        # catalog listed 5.7.0 after 5.7.1). The run skips it and
                        # keeps the newer version on record.
                        if (-not $Ctx.ForceFlag -and $lastKnown -and (Compare-SemVer -A ([string]$lastKnown) -B ([string]$latest) -Strict) -gt 0) {
                            $row.Status = 'Vendor older than last seen'
                            [void]$State.LogQueue.Enqueue(('{0}: the vendor reports {1}, older than {2} seen before - skipping' -f $app, $latest, $lastKnown))
                            $counts['Skipped']++
                            continue
                        }

                        $versionChanged = (-not $lastKnown) -or ($lastKnown -ne $latest)
                        # What a stage or a publish completed is not what a check
                        # saw: a check of v2 after a stage of v1 leaves v2 to build.
                        $fullStageScope   = [string]$planDownloadRoot
                        $fullPackageScope = ('{0}|{1}|{2}' -f [string]$Ctx.SiteCode, [string]$Ctx.FileShareRoot, [string]$planProfileId)
                        $actedVersion = switch ([string]$Ctx.Action) {
                            'Stage'           { Get-PackagerHistoryActedVersion -Entry $fullEntry -Event Staged -Scope $fullStageScope }
                            'StageAndPackage' { Get-PackagerHistoryActedVersion -Entry $fullEntry -Event Packaged -Scope $fullPackageScope }
                            default           { $lastKnown }
                        }
                        $actionStale = (-not $actedVersion) -or ([string]$actedVersion -ne [string]$latest)

                        # A vendor-version match alone is not freshness: a
                        # changed profile revision or signing policy makes
                        # the last build stale even at the same version.
                        $buildStale = $false
                        if ($planApplicationId) {
                            $current = Test-WorkbenchBuildIsCurrent -ApplicationId $planApplicationId -ProfileId $planProfileId `
                                -Version $latest -ProfileRevision $planRevision -PolicyDigest ([string]$Ctx.SigningDigest)
                            if ($null -ne $current -and -not $current) { $buildStale = $true }
                        }
                        $shouldAct      = $actionStale -or $Ctx.ForceFlag -or $buildStale
                        if ($buildStale -and -not $actionStale) {
                            [void]$State.LogQueue.Enqueue(('Rebuilding {0} at the same version: profile revision or signing policy changed.' -f $app))
                        }

                        $histResult = if ($versionChanged) { 'Updated' } else { 'NoChange' }
                        try {
                            Update-PackagerHistory -PackagerName $baseName -Event Checked -Version $latest -Result $histResult
                            $row.LastChecked = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                        } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }

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
                                if (-not [string]::IsNullOrWhiteSpace([string]$Ctx.IntuneWinToolPath)) {
                                    $stageIntuneWin = Invoke-PackagerIntuneWinPostStep -Result $stg -PackagerPath $path -DownloadRoot $planDownloadRoot -ToolPath ([string]$Ctx.IntuneWinToolPath) -SkipNetworkCopy
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $stageIntuneWin.Message))
                                }
                                try {
                                    if ($latest) { Update-PackagerHistory -PackagerName $baseName -Event Staged -Version $latest -Result Updated -Scope $fullStageScope }
                                    else         { Update-PackagerHistory -PackagerName $baseName -Event Staged -Result Updated }
                                } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
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
                            [void]$State.LogQueue.Enqueue(('Stage exception: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                        }

                        if (-not $stageOk) {
                            $counts['Failed']++
                            continue
                        }

                        if ($Ctx.Action -eq 'Stage') {
                            [void]$State.Succeeded.Enqueue($scrName)
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
                                WsusPublishConfig    = $Ctx.WsusPublishConfig
                                DeploymentTarget     = $(if ([string]$Ctx.DeploymentTarget -in (Get-DeploymentTargetNames)) { [string]$Ctx.DeploymentTarget } else { 'MECM' })
                                ContentLayout        = [string]$Ctx.ContentLayout
                                RequirementsJson     = $reqJson
                                VariantsJson         = $varJson
                                CommandsJson         = $cmdJson
                                InstallMode          = $(if ($Ctx.InstallModesByApp) { [string]$Ctx.InstallModesByApp[$baseName] } else { '' })
                            }
                            $packageArgs['TitleMode'] = $(if ($Ctx.TitleModesByApp -and [string]$Ctx.TitleModesByApp[$baseName]) { [string]$Ctx.TitleModesByApp[$baseName] } else { [string]$Ctx.DefaultTitleMode })
                            $pkg = Invoke-PackagerPackageWithConflictPrompt -State $State -AppLabel $app -PackageArgs $packageArgs
                            if ($pkg.PSObject.Properties['WsusPublish'] -and $pkg.WsusPublish) {
                                [void]$State.LogQueue.Enqueue(('WSUS publish: ' + $pkg.WsusPublish.Message))
                            }
                            if ($pkg.ExitCode -eq 0) {
                                if ($pkg.PSObject.Properties['IntunePublish'] -and $pkg.IntunePublish) {
                                    [void]$State.LogQueue.Enqueue(('Intune publish: ' + $pkg.IntunePublish.Message))
                                }
                                if ($pkg.PSObject.Properties['IntuneWin'] -and $pkg.IntuneWin) {
                                    [void]$State.LogQueue.Enqueue(('Intunewin: ' + $pkg.IntuneWin.Message))
                                    if ($pkg.IntuneWin.NetworkPath) {
                                        [void]$State.LogQueue.Enqueue(('Intunewin on network: ' + $pkg.IntuneWin.NetworkPath))
                                    }
                                }
                            }
                            $outcome = Get-PublishRowOutcome -Result $pkg -DeploymentTarget ([string]$packageArgs.DeploymentTarget)
                            $row.Status = $outcome.Status

                            if ($outcome.Kind -eq 'Skipped') {
                                $counts['Skipped']++
                            } elseif ($outcome.Kind -eq 'NotSupported') {
                                $counts['NotSupported']++
                            } elseif ($outcome.Kind -eq 'Success') {
                                [void]$State.LogQueue.Enqueue(('{0}. Logs: {1}' -f $outcome.Status, (Split-Path -Leaf $pkg.OutLog)))
                                try {
                                    if ($latest) { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Version $latest -Result Updated -Scope $fullPackageScope }
                                    else         { Update-PackagerHistory -PackagerName $baseName -Event Packaged -Result Updated }
                                } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                                [void]$State.Succeeded.Enqueue($scrName)
                                $counts[$(if ($outcome.Published) { 'Published' } else { 'StageAndPackage' })]++
                            } elseif ($pkg.ExitCode -eq 0) {
                                [void]$State.LogQueue.Enqueue(('Logs: ' + (Split-Path -Leaf $pkg.OutLog)))
                                $counts['Failed']++
                            } else {
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
                            [void]$State.LogQueue.Enqueue(('Package exception: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                            $counts['Failed']++
                        }
                    }
                }
            }
            $State.Counts = $counts
        }
        catch {
            $State.ErrorMsg = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
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
                Add-LogLine -Message ('Conflict prompt failed, keeping the existing application: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            }
            finally {
                $script:ConflictPromptOpen = $false
            }
        }

        # Re-render the grid so row.Status flips done in the bg are visible.
        try { $dataGrid.Items.Refresh() } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
        if ($script:BgContext -and $script:BgContext.ContainsKey('OneClickTick')) { try { & $script:BgContext.OneClickTick } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ } }

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
                        $summaryLabel = switch ($script:BgOperation) {
                            'CheckLatest' { 'Check Latest summary:'; break }
                            'CheckCve'    { 'Check CVEs summary:'; break }
                            'Stage'       { 'Stage summary:'; break }
                            'Package'     { 'Package summary:'; break }
                            'FullRun'     { 'One Click summary:'; break }
                            'OneClick'    { 'One Click summary:'; break }
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

            # A staged app stays checked: Stage then Publish is one workflow on
            # the same selection. A publish clears its rows.
            if ($doneState.Succeeded -and $script:BgOperation -ne 'Stage') {
                Clear-SucceededSelection -Rows $script:PackagerData -Scripts $doneState.Succeeded.ToArray()
            }
            try { $dataGrid.Items.Refresh() } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $window.Cursor = $null
            Set-ActionButtonsEnabled -Enabled $true
            $btnPausePipeline.Content = 'Pause'
            $btnPausePipeline.IsEnabled = $true
            $btnCancelPipeline.IsEnabled = $true
            $script:BgTimer = $null
            $script:BgState = $null
            if ($script:BgContext -and $script:BgContext.ContainsKey('OneClickDone')) { try { & $script:BgContext.OneClickDone $doneState } catch { Add-LogLine -Message ('One Click report failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) } }
            if ($script:BgContext -and $script:BgContext.ContainsKey('CveDone')) { try { & $script:BgContext.CveDone $doneState } catch { Add-LogLine -Message ('CVE window update failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) } }
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
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
        <TextBlock Grid.Row="1" x:Name="txtDropDetected" Margin="0,4,0,12" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>

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

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Application" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtDropAppName" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Publisher" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtDropPublisher" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="2" Grid.Column="0" Text="Version" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtDropVersion" Height="28" Margin="0,0,0,8"/>

            <!-- Shown only when the installer script accepts a mode switch; the
                 two radios swap every mode-dependent field together. -->
            <TextBlock Grid.Row="3" Grid.Column="0" x:Name="lblDropMode" Text="Install For" VerticalAlignment="Center" Margin="0,0,12,8" Visibility="Collapsed"/>
            <StackPanel Grid.Row="3" Grid.Column="1" x:Name="pnlDropMode" Orientation="Horizontal" Margin="0,0,0,8" Visibility="Collapsed">
                <RadioButton x:Name="radDropCurrentUser" Content="Current user" GroupName="DropMode" VerticalAlignment="Center" Margin="0,0,18,0" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <RadioButton x:Name="radDropAllUsers" Content="All users (system)" GroupName="DropMode" VerticalAlignment="Center" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <TextBlock x:Name="txtDropModeSwitch" VerticalAlignment="Center" Margin="14,0,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>

            <TextBlock Grid.Row="4" Grid.Column="0" Text="Install Args" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="4" Grid.Column="1" x:Name="txtDropInstallArgs" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="5" Grid.Column="0" x:Name="lblDropUninstall" Text="Uninstall Command" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,8"/>
            <TextBox   Grid.Row="5" Grid.Column="1" x:Name="txtDropUninstallCmd" Height="28" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="6" Grid.Column="0" Text="Detection" VerticalAlignment="Top" Margin="0,4,12,0"/>
            <TextBlock Grid.Row="6" Grid.Column="1" x:Name="txtDropDetection" TextWrapping="Wrap" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,4,0,0"/>
        </Grid>

        <CheckBox Grid.Row="3" x:Name="chkDropConfirm" Margin="0,12,0,0"
                  Content="I verified the predicted silent switches and detection for this installer"/>

        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnDropSaveApplication" Content="Save as Application" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource DialogButton}"
                    ToolTip="Save this installer as an application that opens in the Application Workbench."/>
            <Button x:Name="btnDropSavePackager" Content="Save as Packager" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource DialogButton}"/>
            <Button x:Name="btnDropStage" Content="Stage" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource DialogButton}"/>
            <Button x:Name="btnDropStagePackage" Content="Stage and Package" Margin="0,0,8,0" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource DialogAccentButton}"/>
            <Button x:Name="btnDropCancel" Content="Cancel" IsCancel="True" Controls:ControlsHelper.ContentCharacterCasing="Normal" Style="{DynamicResource DialogButton}"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = Expand-SharedStyle $dlgXaml
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
                [void](Show-ThemedMessage -Owner $dlg -Title 'Install Mode' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Warning)
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
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save as Application' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
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
                    [void]$State.LogQueue.Enqueue(('Stage failed: {0}: {1}' -f $v.AppName, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
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
                    [void]$State.LogQueue.Enqueue(('Package failed: {0}: {1}' -f $v.AppName, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')))
                }
            }
            $State.Counts = $counts
        }
        catch {
            $State.ErrorMsg = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
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
    if (($script:BgState -and -not [bool]$script:BgState.Done) -or $script:CmImportRunning -or $null -ne $script:DiscoveryJob) {
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
            -Message 'Staging a dropped installer needs a download root. Open Options, Packaging, Folders to set it.' `
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
        $packageAvailable = $false; $packageReason = 'SiteCode is not configured (Options, Connections, ConfigMgr).'
    }
    elseif ([string]::IsNullOrWhiteSpace($script:Prefs.FileShareRoot)) {
        $packageAvailable = $false; $packageReason = 'The content share is not configured (Options, Connections, ConfigMgr).'
    }

    $jobs = @()
    $installerTotal = @($installers).Count
    $installerIndex = 0
    foreach ($installer in $installers) {
        $installerIndex++
        Show-BusyIndicator -Title 'Analyzing dropped installers' -Step (Format-BusyStep -Verb 'Analyze' -Index $installerIndex -Total $installerTotal -Name (Split-Path -Leaf $installer))
        try {
            $analysis = Get-InstallerAnalysis -Path $installer
        }
        catch {
            Hide-BusyIndicator
            Add-LogLine -Message ('Analysis failed for {0}: {1}' -f (Split-Path -Leaf $installer), (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            continue
        }
        Hide-BusyIndicator
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
                [void](Show-ThemedMessage -Owner $window -Title 'Save Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
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
    $dlg.Title  = 'Select Installers to Package'
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
    # Packagers take a site code parameter that a version check never reads;
    # without ConfigMgr in use the packager default stands in.
    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        if ($script:Prefs.Systems.ConfigMgr) {
            Add-LogLine -Message "The site code is required. Set it in Options, Connections, ConfigMgr."
            $txtStatus.Text = "The site code is required."
            return
        }
        $siteCodeValue = 'MCM'
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
    # The check reads the saved site code and site server. The footer shows the same values.
    Update-ApFooter
    if (-not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Check ConfigMgr requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and restart the computer after you install it) before running Check ConfigMgr." `
            -Buttons OK -Icon Warning)
        return
    }

    $siteCodeValue = $script:Prefs.SiteCode
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "The site code is required. Set it in Options, Connections, ConfigMgr."
        $txtStatus.Text = "The site code is required."
        return
    }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    Set-ActionButtonsEnabled -Enabled $false
    Show-BusyIndicator -Title 'Checking ConfigMgr'

    try {
        $txtStatus.Text = "Querying ConfigMgr for selected products..."

        $cmTotal = @($selectedRows).Count
        $cmIndex = 0
        foreach ($row in $selectedRows) {
            $cmIndex++
            Set-BusyStep -Text (Format-BusyStep -Verb 'ConfigMgr' -Index $cmIndex -Total $cmTotal -Name ([string]$row.Application))

            $app    = [string]$row.Application
            $cmName = [string]$row.CMName

            Add-LogLine -Message ("ConfigMgr: {0}" -f $app)
            $row.Status = "Querying ConfigMgr..."
            $dataGrid.Items.Refresh()

            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -ProviderMachineName $script:Prefs.ProviderMachineName -CMName $cmName

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script)
                if (-not $res.Found) {
                    $row.CurrentVersion = ""
                    Clear-StaleCveCells -Row $row
                    $row.Status = "Not found in ConfigMgr"
                    try { Update-PackagerHistory -PackagerName $baseName -Event CmChecked -Version '' -Data @{ SiteCode = $siteCodeValue } } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }
                    Add-LogLine -Message "Not found."
                    continue
                }

                $row.CurrentVersion = [string]$res.SoftwareVersion
                Clear-StaleCveCells -Row $row
                try { Update-PackagerHistory -PackagerName $baseName -Event CmChecked -Version ([string]$res.SoftwareVersion) -Data @{ SiteCode = $siteCodeValue } } catch { <# Reading or writing packager history or cadence data is bookkeeping; the main flow does not depend on it. #> $null = $_ }

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
                Add-LogLine -Message ("Error: {0}" -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            }
        }

        Hide-BusyIndicator
        Select-OnlyUpdateAvailable
        $dataGrid.Items.Refresh()

        $notFound = @(Get-AutoHideOffer -Prefs $script:Prefs -Rows @($script:PackagerData))
        if ($notFound.Count -gt 0) {
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

        $txtStatus.Text = "ConfigMgr query complete."
        if ([bool]$script:Prefs.Vulnerability.LookupAfterCheckCm) { $script:CveAfterCmRows = $selectedRows }
    }
    finally {
        Hide-BusyIndicator
        Set-ActionButtonsEnabled -Enabled $true
    }
    # The CVE lookup starts after this handler released the buttons; it
    # runs off the UI thread with its own overlay.
    if ($script:CveAfterCmRows) {
        $cveRows = $script:CveAfterCmRows
        $script:CveAfterCmRows = $null
        $txtStatus.Text = "Checking CVEs for the queried packagers..."
        Invoke-MultiAppPipeline -Operation CheckCve -Rows $cveRows -Context (Get-CveRunContext)
    }
})

# --- 2b. Discover unmanaged software ---
function Show-DiscoveryDialog {
    param(
        [AllowNull()]$Owner,
        [Parameter(Mandatory)][object[]]$Candidates,
        [string]$StatusText = '',
        [hashtable]$NameByScript = @{},
        # Test hook: receives the window internals instead of showing the window.
        [scriptblock]$Probe = $null
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Discover Unmanaged Software" Width="860" Height="520" MinWidth="700" MinHeight="360"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="16,12,16,8">
            <TextBlock x:Name="txtDiscoverPrompt" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap"/>
            <TextBlock x:Name="txtDiscoverStatus" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
            <TextBlock FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"
                       Text="Each app is hidden by the Product Filter, has no ConfigMgr application, and is installed on at least one client. Enabled apps join the grid for Check, Stage and Package."/>
            <TextBlock x:Name="txtDiscoverNote" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Visibility="Collapsed"/>
        </StackPanel>
        <DataGrid RowStyle="{DynamicResource GridRowStyle}" CellStyle="{DynamicResource GridCellStyle}" Grid.Row="1" x:Name="gridDiscover" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column" Margin="16,0,16,8"
                  SelectionMode="Single" RowHeaderWidth="0">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="40"/>
                <DataGridTextColumn Header="Application" Binding="{Binding Name}" Width="*" MinWidth="200" IsReadOnly="True"/>
                <DataGridTextColumn MinWidth="94" Header="Vendor" Binding="{Binding Vendor}" Width="180" IsReadOnly="True"/>
                <DataGridTextColumn MinWidth="156" Header="Inventory Name" Binding="{Binding MatchedName}" Width="220" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
        <Border Grid.Row="2" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <WrapPanel Margin="16,10,16,10" HorizontalAlignment="Right">
                <Button x:Name="btnEnableSelected" Content="Enable Selected" Margin="0,0,8,0"
                        Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Show the checked apps in the main grid."/>
                <Button x:Name="btnEnableAll" Content="Enable All" Margin="0,0,8,0"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Show every app of this list in the main grid."/>
                <Button x:Name="btnNotNow" Content="Not Now" Margin="0,0,8,0" IsCancel="True"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Close and offer these apps again in the next discovery."/>
                <Button x:Name="btnDismiss" Content="Don't Ask Again for Selected"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Hide the checked apps in this and later discoveries."/>
            </WrapPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $txtPrompt = $dlg.FindName('txtDiscoverPrompt')
    $txtStatusLine = $dlg.FindName('txtDiscoverStatus')
    $txtNote = $dlg.FindName('txtDiscoverNote')
    $grid = $dlg.FindName('gridDiscover')
    $btnEnableSelected = $dlg.FindName('btnEnableSelected')
    $btnEnableAll = $dlg.FindName('btnEnableAll')
    $btnNotNow = $dlg.FindName('btnNotNow')
    $btnDismiss = $dlg.FindName('btnDismiss')

    $rows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    foreach ($c in @($Candidates | Sort-Object -Property CMName)) {
        $scriptName = [string]$c.Script
        $name = $(if ($NameByScript.ContainsKey($scriptName) -and $NameByScript[$scriptName]) { [string]$NameByScript[$scriptName] } else { [string]$c.CMName })
        $rows.Add([pscustomobject]@{ Selected = $false; Script = $scriptName; Name = $name; Vendor = [string]$c.Vendor; MatchedName = [string]$c.MatchedName })
    }
    $grid.ItemsSource = $rows
    $txtPrompt.Text = Get-DiscoveryPromptText -Count $rows.Count
    $txtStatusLine.Text = $StatusText

    $state = @{ Action = 'NotNow'; Scripts = @() }

    $apply = {
        param([string]$Button)
        # A check box edit stays uncommitted until the row loses focus.
        [void]$grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
        $choice = Resolve-DiscoveryButton -Button $Button -Rows @($rows)
        if ($choice.Note) {
            $txtNote.Text = $choice.Note
            $txtNote.Visibility = 'Visible'
            return
        }
        $state.Action = $choice.Action
        $state.Scripts = @($choice.Scripts)
        $dlg.Close()
    }.GetNewClosure()
    $btnEnableSelected.Add_Click({ & $apply 'EnableSelected' }.GetNewClosure())
    $btnEnableAll.Add_Click({ & $apply 'EnableAll' }.GetNewClosure())
    $btnNotNow.Add_Click({ $dlg.Close() }.GetNewClosure())
    $btnDismiss.Add_Click({ & $apply 'Dismiss' }.GetNewClosure())

    if ($Probe) {
        & $Probe @{ Window = $dlg; Grid = $grid; Rows = $rows; Prompt = $txtPrompt; Status = $txtStatusLine; Note = $txtNote
            EnableSelected = $btnEnableSelected; EnableAll = $btnEnableAll; NotNow = $btnNotNow; Dismiss = $btnDismiss; State = $state }
    }
    else {
        if ($Owner) { $dlg.Owner = $Owner }
        [void]$dlg.ShowDialog()
    }
    return [pscustomobject]@{ Action = [string]$state.Action; Scripts = @($state.Scripts) }
}

function Invoke-DiscoveryChoice {
    # Applies the dialog result on the UI thread. Enabling follows the
    # Product Filter path: change HiddenApplications, save, rebuild the grid.
    param([Parameter(Mandatory)]$Choice)

    switch ([string]$Choice.Action) {
        { $_ -in @('EnableSelected', 'EnableAll') } {
            $count = Enable-DiscoveredApplication -Prefs $script:Prefs -Script @($Choice.Scripts)
            # A choice that is not saved holds for this session only; a grid
            # that is not rebuilt does not show it. Each failure is said.
            $problems = @()
            try { Save-Preferences -Prefs $script:Prefs }
            catch { $problems += ('The choice was not saved and holds for this session only: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
            try { Invoke-RefreshGrid }
            catch { $problems += ('The grid was not refreshed: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
            if ($problems.Count -eq 0) {
                Add-LogLine -Message ("Discover unmanaged software: {0} app(s) added to the grid. Manage them in Product Filter." -f $count)
                $txtStatus.Text = ("{0} app(s) added to the grid." -f $count)
            }
            else {
                $text = ("{0} app(s) enabled. {1}" -f $count, ($problems -join ' '))
                Add-LogLine -Message ('Discover unmanaged software: ' + $text)
                $txtStatus.Text = $text
            }
        }
        'Dismiss' {
            $count = Add-DiscoveryDismissed -Prefs $script:Prefs -Script @($Choice.Scripts)
            try {
                Save-Preferences -Prefs $script:Prefs
                Add-LogLine -Message ("Discover unmanaged software: {0} app(s) will not be offered again. Clear the list in Options, Updates and CVEs, Product Filter." -f $count)
            }
            catch {
                $text = ('The dismissal was not saved and holds for this session only: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
                Add-LogLine -Message ('Discover unmanaged software: ' + $text)
                $txtStatus.Text = $text
            }
        }
        default {
            Add-LogLine -Message 'Discover unmanaged software: no change.'
        }
    }
}

function Complete-DiscoveryScan {
    param($Result)

    Set-ActionButtonsEnabled -Enabled $true
    Hide-BusyIndicator
    if (-not $Result -or -not $Result.Ok) {
        $message = 'Discover unmanaged software: ' + $(if ($Result -and $Result.Error) { [string]$Result.Error } else { 'the query returned no result.' })
        Add-LogLine -Message $message
        $txtStatus.Text = $message
        return
    }
    Add-LogLine -Message ("Discover unmanaged software: queried {0} on {1}. {2}" -f $Result.Database, $Result.Server, $Result.StatusText)
    $txtStatus.Text = [string]$Result.StatusText
    if (@($Result.Candidates).Count -eq 0) { return }

    try {
        $choice = Show-DiscoveryDialog -Owner $window -Candidates @($Result.Candidates) -StatusText ([string]$Result.StatusText) -NameByScript $script:DiscoveryNames
        Invoke-DiscoveryChoice -Choice $choice
    }
    catch {
        Add-LogLine -Message ('Discover unmanaged software: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
    }
}

function Stop-DiscoveryScan {
    if ($script:DiscoveryTimer) { $script:DiscoveryTimer.Stop() }
    $job = $script:DiscoveryJob
    $script:DiscoveryJob = $null
    Stop-DiscoveryRunspace -Job $job
    Complete-DiscoveryScan -Result ([pscustomobject]@{ Ok = $false; Error = 'cancelled.' })
}

function Start-DiscoveryScan {
    $packagers = @(Get-DiscoveryPackagerInput -Rows @(Get-Packagers -Root $PackagersRoot))
    $script:DiscoveryNames = @{}
    foreach ($p in $packagers) { $script:DiscoveryNames[[string]$p.Script] = [string]$p.App }

    $scanArgs = @{
        SiteCode            = [string]$script:Prefs.SiteCode
        ProviderMachineName = Resolve-DiscoveryProviderName -Preference ([string]$script:Prefs.ProviderMachineName) -SiteCode ([string]$script:Prefs.SiteCode)
        ServerOverride      = [string]$script:Prefs.DiscoverySqlServer
        DatabaseOverride    = [string]$script:Prefs.DiscoverySqlDatabase
        InstanceOverride    = [string]$script:Prefs.DiscoverySqlInstance
        PortOverride        = [int]$script:Prefs.DiscoverySqlPort
        Packager            = $packagers
        HiddenScript        = [string[]]@($script:Prefs.HiddenApplications)
        DismissedScript     = [string[]]@($script:Prefs.DiscoveryDismissed)
    }

    Set-ActionButtonsEnabled -Enabled $false
    Show-BusyIndicator -Title 'Discovering unmanaged software' -Step 'Reading the Add/Remove Programs inventory from the site database...' -AllowCancel
    $txtStatus.Text = 'Reading Add/Remove Programs inventory from the ConfigMgr site database...'
    Add-LogLine -Message 'Discover unmanaged software: reading hardware inventory from the ConfigMgr site database.'

    $script:DiscoveryJob = $null
    try {
        $script:DiscoveryJob = Start-DiscoveryRunspace -ModulePath (Join-Path $PSScriptRoot 'Packagers\AppPackagerDiscovery.psm1') `
            -ScanText ([string]${function:Invoke-DiscoveryScan}) -ScanArgs $scanArgs

        $script:DiscoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:DiscoveryTimer.Interval = [TimeSpan]::FromMilliseconds(250)
        $script:DiscoveryTimer.Add_Tick({
            try {
                $job = $script:DiscoveryJob
                if (-not $job) { return }
                $state = Get-DiscoveryRunspaceState -Job $job -TimeoutSeconds $script:DiscoveryScanTimeoutSeconds
                if ($state -eq 'Running') { return }
                $script:DiscoveryTimer.Stop()
                $script:DiscoveryJob = $null
                $result = if ($state -eq 'Done') { $job.State.Result } else {
                    [pscustomobject]@{ Ok = $false; Error = ('the query did not finish within {0} seconds and was stopped.' -f $script:DiscoveryScanTimeoutSeconds) }
                }
                Stop-DiscoveryRunspace -Job $job
                Complete-DiscoveryScan -Result $result
            }
            catch {
                try { $script:DiscoveryTimer.Stop() } catch { <# Stopping the discovery timer is cleanup inside the error path. #> $null = $_ }
                Stop-DiscoveryRunspace -Job $script:DiscoveryJob
                $script:DiscoveryJob = $null
                Set-ActionButtonsEnabled -Enabled $true
                Hide-BusyIndicator
                Add-LogLine -Message ('Discover unmanaged software: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
            }
        })
        $script:DiscoveryTimer.Start()
    }
    catch {
        Stop-DiscoveryRunspace -Job $script:DiscoveryJob
        $script:DiscoveryJob = $null
        Complete-DiscoveryScan -Result ([pscustomobject]@{ Ok = $false; Error = ('the query could not start: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) })
    }
}

$btnDiscover.Add_Click({
    if (Test-PipelineBusy) {
        Add-LogLine -Message 'Discover unmanaged software: another run is in progress. Wait for it to finish.'
        return
    }
    $sidebar = Get-SidebarSystemState -Prefs $script:Prefs
    if (-not $sidebar.DiscoverEnabled) {
        Add-LogLine -Message $sidebar.DiscoverToolTip
        $txtStatus.Text = $sidebar.DiscoverToolTip
        return
    }
    Start-DiscoveryScan
})

# --- 2c. Convert a ConfigMgr application to Intune ---
# True while the conversion child runs. Test-PipelineBusy, the drop intake and
# the window close handler read it.
$script:CmImportRunning = $false

function Get-CmImportChildEnvironment {
    # The environment values the conversion child needs so that it finds the
    # console module and the SMS Provider that the GUI found. A value that the
    # parent already has stays as it is: the child inherits it.
    param(
        [AllowEmptyString()][string]$ModulePath = '',
        [AllowEmptyString()][string]$ProviderMachineName = '',
        [AllowEmptyString()][string]$ParentAdminUiPath = '',
        [AllowEmptyString()][string]$ParentProvider = ''
    )

    $values = [ordered]@{}
    # The child looks for ConfigurationManager.psd1 in the parent of SMS_ADMIN_UI_PATH.
    if (-not $ParentAdminUiPath -and $ModulePath) {
        $values['SMS_ADMIN_UI_PATH'] = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ModulePath), 'i386')
    }
    if ($ProviderMachineName.Trim() -and -not $ParentProvider) {
        $values['APP_PACKAGER_CM_PROVIDER'] = $ProviderMachineName.Trim()
    }
    return $values
}

function Get-CmImportRunSpec {
    # Checks the dialog values and builds the command line of
    # Convert-CMApplicationToIntune.ps1. It touches no WPF type and starts
    # nothing, so a test runs it headless. Preview writes the reports and
    # copies no content. Publish implies Package.
    param(
        [AllowEmptyString()][string]$ApplicationName = '',
        [AllowEmptyString()][string]$SiteCode = '',
        [AllowEmptyString()][string]$OutputRoot = '',
        [ValidateSet('PerDeploymentType', 'AssignByGroup')][string]$VariantStrategy = 'PerDeploymentType',
        [bool]$NativeDetectionOnly = $false,
        [bool]$AllowUntranslatedRequirements = $false,
        [bool]$Package = $false,
        [bool]$Publish = $false,
        [bool]$IntuneReady = $false,
        [ValidateSet('Preview', 'Convert')][string]$Mode = 'Convert',
        [AllowEmptyString()][string]$ScriptPath = '',
        [AllowEmptyString()][string]$ToolPath = ''
    )

    $fail = { param([string]$Message) [pscustomobject]@{ Ok = $false; Error = $Message; Arguments = @(); Summary = '' } }
    $name = $ApplicationName.Trim()
    $site = $SiteCode.Trim()
    $root = $OutputRoot.Trim()

    if (-not $name) { return (& $fail 'Enter the exact name of the ConfigMgr application.') }
    if ($name -match '[\x00-\x1F]') { return (& $fail 'The application name holds a control character.') }
    if ($site -notmatch '^[A-Za-z0-9]{3}$') { return (& $fail 'The site code is three letters or digits. Set it in Options, Connections, ConfigMgr.') }
    if (-not $root) { return (& $fail 'Enter the output folder.') }
    # IsPathRooted throws on these characters in Windows PowerShell 5.1, so they are checked first.
    if ($root -match '[\x00-\x1F"<>|]') { return (& $fail 'The output folder holds a character that a path cannot use.') }
    if (-not [System.IO.Path]::IsPathRooted($root)) { return (& $fail 'The output folder must be a full path.') }

    $preview = ($Mode -eq 'Preview')
    $doPublish = (-not $preview -and $Publish)
    $doPackage = (-not $preview -and ($Package -or $Publish))
    if ($doPublish -and -not $IntuneReady) {
        return (& $fail 'Publish needs the Intune tenant, client and secret. Set them in Options, Connections, Intune.')
    }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath,
        '-ApplicationName', $name, '-SiteCode', $site, '-OutputRoot', $root, '-VariantStrategy', $VariantStrategy)
    if ($NativeDetectionOnly) { $arguments += '-NativeDetectionOnly' }
    if ($AllowUntranslatedRequirements) { $arguments += '-AllowUntranslatedRequirements' }
    if ($preview) { $arguments += '-SkipContent' }
    if ($doPackage) {
        $arguments += '-Package'
        if ($ToolPath) { $arguments += @('-ToolPath', $ToolPath) }
    }
    if ($doPublish) { $arguments += @('-Publish', '-UseSavedSettings') }

    $summary = $(if ($preview) { "Preview of '{0}' on site {1}: reports only, no content copied." } elseif ($doPublish) { "Conversion of '{0}' on site {1}: copy, build and publish to Intune." } elseif ($doPackage) { "Conversion of '{0}' on site {1}: copy and build." } else { "Conversion of '{0}' on site {1}: copy content." }) -f $name, $site.ToUpperInvariant()
    return [pscustomobject]@{ Ok = $true; Error = ''; Arguments = $arguments; Summary = $summary; Mode = $Mode }
}

function Get-CmImportOutcomeText {
    # The status line after the child process ends.
    param([int]$ExitCode, [AllowEmptyString()][string]$StdErr = '', [string]$Mode = 'Convert', [AllowEmptyString()][string]$OutputRoot = '')

    if ($ExitCode -eq 0) {
        if ($Mode -eq 'Preview') { return ('Preview finished. The report is in ' + $OutputRoot + '.') }
        return ('Conversion finished. The output is in ' + $OutputRoot + '.')
    }
    $firstError = ''
    foreach ($line in @($StdErr -split "`r?`n")) {
        if ($line.Trim()) { $firstError = $line.Trim(); break }
    }
    if ($firstError) { return ('Convert ConfigMgr App to Intune failed: ' + $firstError) }
    return ('Convert ConfigMgr App to Intune failed with exit code {0}. The log has the details.' -f $ExitCode)
}

function Show-CmImportDialog {
    param(
        [AllowNull()]$Owner,
        [AllowEmptyString()][string]$SiteCode = '',
        [AllowEmptyString()][string]$OutputRoot = '',
        [bool]$IntuneReady = $false,
        [AllowEmptyString()][string]$ScriptPath = '',
        [AllowEmptyString()][string]$ToolPath = '',
        # Test hook: receives the window internals instead of showing the window.
        [scriptblock]$Probe = $null
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Convert ConfigMgr App to Intune" Width="640" Height="560" MinWidth="560" MinHeight="520"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
            <StackPanel Margin="16,12,16,8">
                <TextBlock FontSize="12" TextWrapping="Wrap" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"
                           Text="Reads the application from the ConfigMgr site and makes one Intune Win32 app for each deployment type. Nothing changes in ConfigMgr. Preview writes the report only. Convert also copies the content."/>

                <TextBlock Text="ConfigMgr Application Name" FontSize="12" FontWeight="Bold" Margin="0,12,0,4"/>
                <TextBox x:Name="txtCmApp" FontSize="13" Controls:TextBoxHelper.Watermark="The exact name, as the console shows it"
                         ToolTip="The exact name of the ConfigMgr application."/>

                <TextBlock Text="Site Code" FontSize="14" FontWeight="Bold" Margin="0,10,0,4"/>
                <TextBox x:Name="txtCmSite" FontSize="13" Width="120" HorizontalAlignment="Left" MaxLength="3" IsReadOnly="True"
                         ToolTip="Read-only: change the site code in Options, Connections, ConfigMgr."/>

                <TextBlock Text="Output Folder" FontSize="12" FontWeight="Bold" Margin="0,10,0,4"/>
                <DockPanel>
                    <Button x:Name="btnCmBrowse" Content="Browse..." DockPanel.Dock="Right" Margin="8,0,0,0"
                            Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <TextBox x:Name="txtCmOutput" FontSize="13"
                             ToolTip="The folder that holds one subfolder for each application."/>
                </DockPanel>

                <TextBlock Text="Deployment Types" FontSize="12" FontWeight="Bold" Margin="0,10,0,4"/>
                <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}" x:Name="cmbCmStrategy" FontSize="13" SelectedIndex="0">
                    <ComboBoxItem Content="First match wins, like ConfigMgr (recommended)" Tag="PerDeploymentType"/>
                    <ComboBoxItem Content="Assign each app by group" Tag="AssignByGroup"/>
                </ComboBox>
                <TextBlock x:Name="txtCmStrategyHelp" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"
                           Text="Each app after the first gets a requirement script that is false when an earlier deployment type applies."/>

                <CheckBox x:Name="chkCmNative" Content="Native detection rules only" Margin="0,12,0,0"
                          ToolTip="Refuse a detection that needs a PowerShell script in Intune (Or, wildcards, dates, sizes)."/>
                <CheckBox x:Name="chkCmUntranslated" Content="Convert deployment types with requirements Intune cannot check" Margin="0,6,0,0"
                          ToolTip="The Intune app then installs on devices where ConfigMgr skips that deployment type."/>
                <CheckBox x:Name="chkCmPackage" Content="Build the .intunewin files" Margin="0,6,0,0"
                          ToolTip="Build the .intunewin files with the Win32 Content Prep Tool."/>
                <CheckBox x:Name="chkCmPublish" Content="Publish the apps to Intune" Margin="0,6,0,0"
                          ToolTip="Use the Intune tenant, client and secret of Options."/>

                <TextBlock x:Name="txtCmNote" FontSize="12" TextWrapping="Wrap" Margin="0,12,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Visibility="Collapsed"/>
            </StackPanel>
        </ScrollViewer>
        <Border Grid.Row="1" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <WrapPanel Margin="16,10,16,10" HorizontalAlignment="Right">
                <Button x:Name="btnCmPreview" Content="Preview" Margin="0,0,8,0"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Write the plan and the report without a copy of the content."/>
                <Button x:Name="btnCmConvert" Content="Convert" Margin="0,0,8,0"
                        Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Copy the content and make the Intune app folders."/>
                <Button x:Name="btnCmCancel" Content="Cancel" IsCancel="True"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            </WrapPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $txtApp = $dlg.FindName('txtCmApp')
    $txtSite = $dlg.FindName('txtCmSite')
    $txtOutput = $dlg.FindName('txtCmOutput')
    $btnBrowse = $dlg.FindName('btnCmBrowse')
    $cmbStrategy = $dlg.FindName('cmbCmStrategy')
    $txtStrategyHelp = $dlg.FindName('txtCmStrategyHelp')
    $chkNative = $dlg.FindName('chkCmNative')
    $chkUntranslated = $dlg.FindName('chkCmUntranslated')
    $chkPackage = $dlg.FindName('chkCmPackage')
    $chkPublish = $dlg.FindName('chkCmPublish')
    $txtNote = $dlg.FindName('txtCmNote')
    $btnPreview = $dlg.FindName('btnCmPreview')
    $btnConvert = $dlg.FindName('btnCmConvert')
    $btnCancel = $dlg.FindName('btnCmCancel')

    $txtSite.Text = $SiteCode
    $txtOutput.Text = $OutputRoot
    if (-not $IntuneReady) {
        $chkPublish.IsEnabled = $false
        $chkPublish.ToolTip = 'Set Tenant ID, Client ID and Client secret in Options, Connections, Intune.'
        [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($chkPublish, $true)
    }

    $state = @{ Action = 'Cancel'; Spec = $null }
    $syncPublish = {
        # Publish needs the .intunewin files, so it keeps Package checked.
        if ($chkPublish.IsChecked) { $chkPackage.IsChecked = $true; $chkPackage.IsEnabled = $false }
        else { $chkPackage.IsEnabled = $true }
    }.GetNewClosure()
    $chkPublish.Add_Checked($syncPublish)
    $chkPublish.Add_Unchecked($syncPublish)
    $cmbStrategy.Add_SelectionChanged({
        $txtStrategyHelp.Text = $(if ([string]$cmbStrategy.SelectedItem.Tag -eq 'AssignByGroup') {
            'Each app carries only its own requirements. Assign each app to its own device group.'
        } else {
            'Each app after the first gets a requirement script that is false when an earlier deployment type applies.'
        })
    }.GetNewClosure())

    $run = {
        param([string]$Mode)
        $spec = Get-CmImportRunSpec -ApplicationName $txtApp.Text -SiteCode $txtSite.Text -OutputRoot $txtOutput.Text `
            -VariantStrategy ([string]$cmbStrategy.SelectedItem.Tag) -NativeDetectionOnly ([bool]$chkNative.IsChecked) `
            -AllowUntranslatedRequirements ([bool]$chkUntranslated.IsChecked) -Package ([bool]$chkPackage.IsChecked) `
            -Publish ([bool]$chkPublish.IsChecked) -IntuneReady $IntuneReady -Mode $Mode -ScriptPath $ScriptPath -ToolPath $ToolPath
        if (-not $spec.Ok) {
            $txtNote.Text = $spec.Error
            $txtNote.Visibility = 'Visible'
            return
        }
        $state.Action = $Mode
        $state.Spec = $spec
        $state.OutputRoot = $txtOutput.Text.Trim()
        $dlg.Close()
    }.GetNewClosure()
    $btnPreview.Add_Click({ & $run 'Preview' }.GetNewClosure())
    $btnConvert.Add_Click({ & $run 'Convert' }.GetNewClosure())
    $btnCancel.Add_Click({ $dlg.Close() }.GetNewClosure())
    $btnBrowse.Add_Click({
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $picker = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($txtOutput.Text -and (Test-Path -LiteralPath $txtOutput.Text -PathType Container)) { $picker.SelectedPath = $txtOutput.Text }
        if ($picker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOutput.Text = $picker.SelectedPath }
    }.GetNewClosure())

    if ($Probe) {
        & $Probe @{ Window = $dlg; App = $txtApp; Site = $txtSite; Output = $txtOutput; Strategy = $cmbStrategy; StrategyHelp = $txtStrategyHelp
            Native = $chkNative; Untranslated = $chkUntranslated; Package = $chkPackage; Publish = $chkPublish; Note = $txtNote
            Preview = $btnPreview; Convert = $btnConvert; Cancel = $btnCancel; State = $state }
    }
    else {
        if ($Owner) { $dlg.Owner = $Owner }
        [void]$dlg.ShowDialog()
    }
    return [pscustomobject]@{ Action = [string]$state.Action; Spec = $state.Spec; OutputRoot = [string]$state.OutputRoot }
}

function Start-CmImportRun {
    param([Parameter(Mandatory)]$Spec, [AllowEmptyString()][string]$OutputRoot = '')

    Set-ActionButtonsEnabled -Enabled $false
    Show-BusyIndicator -Title 'Converting a ConfigMgr application' -Step 'Starting the conversion...'
    $script:CmImportRunning = $true
    try {
        $logFolder = Join-Path $PSScriptRoot 'Logs'
        if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $outLog = Join-Path $logFolder ("convert-cmapp-{0}.out.log" -f $stamp)
        $errLog = Join-Path $logFolder ("convert-cmapp-{0}.err.log" -f $stamp)

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.WorkingDirectory = $PSScriptRoot
        Set-ProcessStartInfoArgumentList -StartInfo $psi -Arguments @($Spec.Arguments)
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $childEnvironment = Get-CmImportChildEnvironment `
            -ModulePath $(if ($script:Prefs.DetectedTools -and $script:Prefs.DetectedTools.ConfigMgrConsole) { [string]$script:Prefs.DetectedTools.ConfigMgrConsole.ModulePath } else { '' }) `
            -ProviderMachineName ([string]$script:Prefs.ProviderMachineName) `
            -ParentAdminUiPath ([string]$env:SMS_ADMIN_UI_PATH) -ParentProvider ([string]$env:APP_PACKAGER_CM_PROVIDER)
        foreach ($name in $childEnvironment.Keys) { $psi.EnvironmentVariables[$name] = [string]$childEnvironment[$name] }

        Add-LogLine -Message ('Convert ConfigMgr App to Intune: ' + $Spec.Summary)
        $txtStatus.Text = 'Converting the ConfigMgr application...'
        Set-BusyStep -Text 'Converting the ConfigMgr application. The log shows the progress.'
        # A content copy can print nothing for a long time, so the idle limit is one hour.
        $result = Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -LogTextBox $txtLog -IdleTimeoutSeconds 3600
        $text = Get-CmImportOutcomeText -ExitCode ([int]$result.ExitCode) -StdErr ([string]$result.StdErr) -Mode ([string]$Spec.Mode) -OutputRoot $OutputRoot
        Add-LogLine -Message $text
        $txtStatus.Text = $text
    }
    catch {
        $message = 'Convert ConfigMgr App to Intune failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
        Add-LogLine -Message $message
        $txtStatus.Text = $message
    }
    finally {
        $script:CmImportRunning = $false
        Hide-BusyIndicator
        Set-ActionButtonsEnabled -Enabled $true
    }
}

$btnConvertCm.Add_Click({
    if (Test-PipelineBusy) {
        Add-LogLine -Message 'Convert ConfigMgr App to Intune: another run is in progress. Wait for it to finish.'
        return
    }
    $sidebar = Get-SidebarSystemState -Prefs $script:Prefs
    if (-not $sidebar.ConvertCmEnabled) {
        Add-LogLine -Message $sidebar.ConvertCmToolTip
        $txtStatus.Text = $sidebar.ConvertCmToolTip
        return
    }
    $defaultOutput = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$script:Prefs.DownloadRoot)) { $defaultOutput = Join-Path ([string]$script:Prefs.DownloadRoot) 'ConfigMgrImport' }
    $choice = Show-CmImportDialog -Owner $window -SiteCode ([string]$script:Prefs.SiteCode) -OutputRoot $defaultOutput `
        -IntuneReady ([bool]$sidebar.IntuneEnabled) -ScriptPath (Join-Path $PSScriptRoot 'Convert-CMApplicationToIntune.ps1') `
        -ToolPath (Get-IntuneWinToolPathForContext)
    if ($choice.Action -eq 'Cancel') { return }
    Start-CmImportRun -Spec $choice.Spec -OutputRoot $choice.OutputRoot
})

# --- 3. Stage Packages ---
$btnStage.Add_Click({
    $dlRootValue = $script:Prefs.DownloadRoot
    if ([string]::IsNullOrWhiteSpace($dlRootValue)) {
        Add-LogLine -Message "Download root is required for staging. Set it in Options, Packaging, Folders."
        $txtStatus.Text = "Download root is required."
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
        DownloadRoot      = $dlRootValue
        M365Channel       = $script:Prefs.M365Channel
        M365DeployMode    = $script:Prefs.M365DeployMode
        LogFolder         = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath      = Get-SevenZipPathForContext
        IntuneWinToolPath = Get-IntuneWinToolPathForContext
        RunPlanByApp      = Get-WorkbenchRunPlanForContext -Rows $selectedRows -Target ([string]$script:Prefs.Intune.DeploymentTarget)
        SigningJson       = Get-WorkbenchSigningPolicyJson
        SigningDigest     = Get-WorkbenchSigningPolicyDigest
    }
})

function Confirm-WsusPublishPrerequisites {
    # Runs on the UI thread before a WSUS-bound run: a missing server or a
    # missing administration API would otherwise fail every app after its
    # stage completes.
    param([AllowEmptyString()][string]$Target = '')
    $destination = if ($Target) { $Target } else { [string]$script:Prefs.Intune.DeploymentTarget }
    if (-not (Test-DeploymentTargetPublishesToWsus -DeploymentTarget $destination)) { return $true }
    if (-not (Get-WsusPublishConfigForContext -Target $destination)) {
        Add-LogLine -Message "Publish requires a WSUS server. Set it in Options, Connections, WSUS."
        $txtStatus.Text = "WSUS server required."
        return $false
    }
    $api = Test-WsusAdministrationApi
    if (-not $api.Available) {
        Add-LogLine -Message ("WSUS publishing is unavailable: {0}" -f $api.Reason)
        $txtStatus.Text = "WSUS administration API not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'WSUS API Required' -Message $api.Reason -Buttons OK -Icon Warning)
        return $false
    }
    return $true
}

# --- 4. Publish to ConfigMgr, Intune or WSUS ---
# Each sidebar button names its own destination for this run; One Click keeps
# the destination saved in Options, Updates and CVEs, One Click.
function Start-PublishRun {
    param([Parameter(Mandatory)][ValidateSet('MECM', 'IntuneOnly', 'WSUSOnly')][string]$Target)

    # Intune and WSUS runs never touch the site or the share, so the console,
    # SiteCode, and content share gates apply only to ConfigMgr runs.
    $intuneOnlyRun = ($Target -eq 'IntuneOnly')
    $siteFreeRun = Test-DeploymentTargetSkipsSite -DeploymentTarget $Target
    if (-not $siteFreeRun -and -not $script:Prefs.DetectedTools.ConfigMgrConsole.Found) {
        Add-LogLine -Message "Publish to ConfigMgr requires the ConfigMgr Console. Not detected on this workstation."
        $txtStatus.Text = "ConfigMgr Console not installed."
        [void](Show-ThemedMessage -Owner $window -Title 'Console Required' `
            -Message "The Configuration Manager Console (AdminUI) is not detected on this workstation. Install it (and restart the computer after you install it) before publishing to ConfigMgr." `
            -Buttons OK -Icon Warning)
        return
    }

    $siteCodeValue = $script:Prefs.SiteCode
    if (-not $siteFreeRun -and [string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -Message "The site code is required. Set it in Options, Connections, ConfigMgr."
        $txtStatus.Text = "The site code is required."
        return
    }
    # Invoke-PackagerPackage binds SiteCode as a mandatory string, which
    # refuses an empty value; a site-free run never reads it.
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) { $siteCodeValue = 'MCM' }

    $fsPathValue = $script:Prefs.FileShareRoot
    if (-not $siteFreeRun -and [string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -Message "Content share is required. Set it in Options, Connections, ConfigMgr."
        $txtStatus.Text = "Content share is required."
        return
    }

    if ($intuneOnlyRun -and -not (Get-IntunePublishConfigForContext -Target $Target)) {
        Add-LogLine -Message "Publish requires Intune credentials. Set Tenant ID, Client ID and Client secret in Options, Connections, Intune."
        $txtStatus.Text = "Intune credentials required."
        return
    }

    if (-not (Confirm-WsusPublishPrerequisites -Target $Target)) { return }

    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }

    $txtStatus.Text = switch ($Target) {
        'IntuneOnly' { 'Publishing selected applications to Intune...' }
        'WSUSOnly'   { 'Publishing selected applications to WSUS...' }
        default      { 'Publishing selected applications to ConfigMgr...' }
    }
    $toolPath = Get-IntuneWinToolPathForContext
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
        # A ConfigMgr run places the .intunewin that Stage built beside the
        # network content.
        IntuneWinCreate      = (-not $siteFreeRun -and -not [string]::IsNullOrWhiteSpace($toolPath))
        IntuneWinToolPath    = $toolPath
        RequirementsByApp    = Get-RequirementsMapForContext
        VariantsByApp        = Get-VariantsMapForContext
        CommandsByApp        = Get-CommandsMapForContext
        InstallModesByApp    = Get-InstallModesMapForContext
        TitleModesByApp      = Get-TitleModesMapForContext
        DefaultTitleMode     = Get-DefaultTitleModeForContext
        IntunePublishConfig  = $(if ($intuneOnlyRun) { Get-IntunePublishConfigForContext -Target $Target } else { $null })
        WsusPublishConfig    = Get-WsusPublishConfigForContext -Target $Target
        DeploymentTarget     = $Target
        RunPlanByApp         = Get-WorkbenchRunPlanForContext -Rows $rowsForPlan -Target $Target
        SigningJson          = Get-WorkbenchSigningPolicyJson
        SigningDigest        = Get-WorkbenchSigningPolicyDigest
        # Receipts and refusals count only for the destination instance they
        # were written for.
        InstanceKeys         = Get-OneClickInstanceKeys -Prefs $script:Prefs
    }
}

$btnPackage.Add_Click({ Start-PublishRun -Target 'MECM' })
$btnPublishIntune.Add_Click({ Start-PublishRun -Target 'IntuneOnly' })
$btnPublishWsus.Add_Click({ Start-PublishRun -Target 'WSUSOnly' })

$btnWsusUpdates.Add_Click({
    $settings = Get-WsusPublishConfigForContext -Target 'WSUSOnly'
    if (-not $settings) {
        Add-LogLine -Message "WSUS Updates needs a WSUS server. Set it in Options, Connections, WSUS."
        $txtStatus.Text = "WSUS server required."
        return
    }
    $api = Test-WsusAdministrationApi
    if (-not $api.Available) {
        Add-LogLine -Message ("WSUS Updates is unavailable: {0}" -f $api.Reason)
        [void](Show-ThemedMessage -Owner $window -Title 'WSUS API Required' -Message $api.Reason -Buttons OK -Icon Warning)
        return
    }
    Show-WsusPublishedUpdatesDialog -Owner $window -Settings $settings
})

# --- 4b. CVEs ---
# Check CVEs runs Get-ProductVulnerabilities per checked row off the UI
# thread; the CVE cell opens the detail window from the cache; CVE Report
# writes the files under Logs\Reports from the same cache.

function Open-ExternalUrl {
    # Only https URLs leave the application; everything else is logged and
    # dropped (a vendor reference can be any string).
    param([AllowNull()][AllowEmptyString()][string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url) -or $Url -notmatch '^https://') {
        Add-LogLine -Message ('Link not opened (not https): {0}' -f $Url)
        return
    }
    try { Start-Process $Url } catch { Add-LogLine -Message ('Link failed: {0}' -f (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
}

function Write-CveReportForRows {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Rows, [string]$Name = '')
    $reportRows = Get-CveReportRows -Rows $Rows -Results $script:CveSessionResults
    $paths = Export-VulnerabilityReport -Rows $reportRows -Folder (Get-CveReportFolder) -Name $Name -Title 'AppPackager CVE report'
    Add-LogLine -Message ('CVE report: {0}' -f $paths.Html)
    Add-LogLine -Message ('CVE report: {0}' -f $paths.ApplicationsCsv)
    Add-LogLine -Message ('CVE report: {0}' -f $paths.CvesCsv)
    Add-LogLine -Message ('CVE report: {0}' -f $paths.Json)
    return $paths
}

function ConvertTo-CveDialogRows {
    param($Result)
    $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    # The comma keeps an empty collection from unrolling into nothing.
    if (-not $Result) { return ,$rows }
    # The affecting CVEs first, then the records without a version.
    foreach ($c in @($Result.Cves) + @($Result.UnversionedCves)) {
        if ($null -eq $c) { continue }
        $links = @()
        foreach ($k in $c.Links.Keys) { $links += [pscustomobject]@{ Name = [string]$k; Url = [string]$c.Links[$k] } }
        if ($c.References.Count -gt 0 -and [string]$c.References[0] -match '^https://') { $links += [pscustomobject]@{ Name = 'Reference'; Url = [string]$c.References[0] } }
        $rows.Add([pscustomobject]@{
            Id        = [string]$c.Id
            Status    = $(if ([string]$c.Status -eq 'Unversioned') { 'No version' } else { [string]$c.Status })
            Cvss      = (ConvertTo-VulnerabilityCvssText -Score $c.Score -Severity $c.Severity)
            Epss      = $(if ($null -ne $c.Epss) { ('{0:P1}' -f [double]$c.Epss) } else { '' })
            Kev       = $(if ($c.Kev) { 'Yes' } else { '' })
            FixedIn   = [string]$c.FixedIn + $(if ($c.FixedInLatest) { ' (latest)' } else { '' })
            Published = (ConvertTo-VulnerabilityDayText -Value ([string]$c.Published))
            Modified  = (ConvertTo-VulnerabilityDayText -Value ([string]$c.LastModified))
            Summary   = [string]$c.Summary
            Links     = $links
        })
    }
    return ,$rows
}

function Show-CveDetailDialog {
    param(
        [AllowNull()]$Owner,
        [Parameter(Mandatory)]$Row,
        # Test hook: receives the window internals instead of showing the window.
        [scriptblock]$Probe = $null
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="CVEs" Width="1280" Height="680" MinWidth="980" MinHeight="420"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->
            <Style x:Key="CveLinkButton" TargetType="Button" BasedOn="{StaticResource GridLinkButtonStyle}">
                <Setter Property="Margin" Value="0,0,8,2"/>
                <Setter Property="FontSize" Value="12"/>
            </Style>
            <Style x:Key="ProductLinkButton" TargetType="Button" BasedOn="{StaticResource LinkButtonStyle}">
                <Setter Property="Margin" Value="0,0,8,2"/>
                <Setter Property="FontSize" Value="12"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="16,12,16,8">
            <TextBlock x:Name="txtCveTitle" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap"/>
            <TextBlock x:Name="txtCveRange" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
            <TextBlock x:Name="txtCveNote" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Visibility="Collapsed"/>
            <DockPanel Margin="0,6,0,0">
                <TextBlock Text="Product:" FontSize="12" Margin="0,0,8,0" VerticalAlignment="Center" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
                <ItemsControl x:Name="productLinks">
                    <ItemsControl.ItemsPanel><ItemsPanelTemplate><WrapPanel Orientation="Horizontal"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
                    <ItemsControl.ItemTemplate>
                        <DataTemplate>
                            <Button Tag="{Binding Url}" ToolTip="{Binding Url}" Style="{StaticResource ProductLinkButton}"><TextBlock Text="{Binding Name}" TextDecorations="Underline"/></Button>
                        </DataTemplate>
                    </ItemsControl.ItemTemplate>
                </ItemsControl>
            </DockPanel>
        </StackPanel>
        <DataGrid RowStyle="{DynamicResource GridRowStyle}" CellStyle="{DynamicResource GridCellStyle}" Grid.Row="1" x:Name="gridCves" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column" Margin="16,0,16,8"
                  SelectionMode="Single" RowHeaderWidth="0" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn MinWidth="64" Header="CVE" Binding="{Binding Id}" Width="140"/>
                <DataGridTextColumn MinWidth="100" Header="Status" Binding="{Binding Status}" Width="100"/>
                <DataGridTextColumn MinWidth="72" Header="CVSS" Binding="{Binding Cvss}" Width="110"/>
                <DataGridTextColumn MinWidth="70" Header="EPSS" Binding="{Binding Epss}" Width="70"/>
                <DataGridTextColumn MinWidth="65" Header="KEV" Binding="{Binding Kev}" Width="65"/>
                <DataGridTextColumn MinWidth="94" Header="Fixed In" Binding="{Binding FixedIn}" Width="110"/>
                <DataGridTextColumn MinWidth="111" Header="Published" Binding="{Binding Published}" Width="111"/>
                <DataGridTextColumn MinWidth="104" Header="Modified" Binding="{Binding Modified}" Width="104"/>
                <DataGridTextColumn x:Name="colSummary" Header="Summary" Binding="{Binding Summary}" Width="*" MinWidth="260"/>
                <DataGridTemplateColumn Header="Links" Width="380" MinWidth="300">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ItemsControl ItemsSource="{Binding Links}">
                                <ItemsControl.ItemsPanel><ItemsPanelTemplate><WrapPanel Orientation="Horizontal"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
                                <ItemsControl.ItemTemplate>
                                    <DataTemplate>
                                        <Button Tag="{Binding Url}" ToolTip="{Binding Url}" Style="{StaticResource CveLinkButton}"><TextBlock Text="{Binding Name}" TextDecorations="Underline"/></Button>
                                    </DataTemplate>
                                </ItemsControl.ItemTemplate>
                            </ItemsControl>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
            </DataGrid.Columns>
        </DataGrid>
        <Border Grid.Row="2" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <DockPanel Margin="16,10,16,10">
                <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Close"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnRefresh" Content="Refresh" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Look up the CVEs again and skip the cache."/>
                    <Button x:Name="btnCopy" Content="Copy CVE IDs" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <Button x:Name="btnExport" Content="Export Report" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Write the report of this app and open the HTML file."/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $txtTitle = $dlg.FindName('txtCveTitle')
    $txtRange = $dlg.FindName('txtCveRange')
    $txtNote = $dlg.FindName('txtCveNote')
    $productLinks = $dlg.FindName('productLinks')
    $gridCves = $dlg.FindName('gridCves')
    $btnRefresh = $dlg.FindName('btnRefresh')
    $btnCopy = $dlg.FindName('btnCopy')
    $btnExport = $dlg.FindName('btnExport')
    $btnClose = $dlg.FindName('btnClose')

    $colSummary = $dlg.FindName('colSummary')
    $wrap = [System.Windows.Style]::new([System.Windows.Controls.TextBlock], $colSummary.ElementStyle)
    $wrap.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::Wrap))
    # A long summary (the Terrapin record lists every affected product) is
    # clipped to six lines; the tooltip carries the full text.
    $wrap.Setters.Add([System.Windows.Setter]::new([System.Windows.FrameworkElement]::MaxHeightProperty, [double]96))
    $wrap.Setters.Add([System.Windows.Setter]::new([System.Windows.FrameworkElement]::ToolTipProperty, (New-Object System.Windows.Data.Binding 'Summary')))
    $colSummary.ElementStyle = $wrap

    $state = @{ Result = $null; Row = $Row }
    $txtTitle.Text = ('{0} ({1})' -f [string]$Row.Application, [string]$Row.Vendor)

    $render = {
        $res = $state.Result
        $productLinks.ItemsSource = $null
        if (-not $res) {
            $txtRange.Text = ('ConfigMgr {0}, latest {1}' -f $(if ([string]$Row.CurrentVersion) { [string]$Row.CurrentVersion } else { '-' }), $(if ([string]$Row.LatestVersion) { [string]$Row.LatestVersion } else { '-' }))
            $reason = [string]$Row.CveCount
            if (-not $reason -or $reason -match '^\d+$' -or $reason -eq 'Not in NVD') { $reason = 'Not checked. Run Check CVEs for this app. A lookup needs the ConfigMgr version or the latest version.' }
            $txtNote.Text = $reason
            $txtNote.Visibility = 'Visible'
            $links = @()
            if ([string]$Row.Cpe) {
                $advisory = $(if ($Row.PSObject.Properties['AdvisoryUrl']) { [string]$Row.AdvisoryUrl } else { '' })
                $set = Get-VulnerabilityLinks -Cpe ([string]$Row.Cpe) -VersionStart $(if ([string]$Row.CurrentVersion) { [string]$Row.CurrentVersion } else { [string]$Row.LatestVersion }) -VendorAdvisoryUrl $advisory
                foreach ($k in $set.Product.Keys) { $links += [pscustomobject]@{ Name = [string]$k; Url = [string]$set.Product[$k] } }
            }
            $productLinks.ItemsSource = $links
            $gridCves.ItemsSource = (ConvertTo-CveDialogRows -Result $null)
            return
        }
        $fixText = ConvertTo-VulnerabilityFixText -Result $res
        if ($fixText) { $fixText = ' ' + $fixText + '.' }
        $dates = ''
        if ($res.NewestPublished) { $dates = ' Newest CVE {0}, NVD updated {1}.' -f $res.NewestPublished, $res.LastNvdUpdate }
        $txtRange.Text = ('{0}{1} Source {2}, checked {3}.{4}' -f (ConvertTo-VulnerabilityResultText -Result $res), $fixText, $(if ($res.Source) { $res.Source } else { '-' }), $res.QueriedAt, $dates)
        if ($res.Error) { $txtNote.Text = [string]$res.Error; $txtNote.Visibility = 'Visible' } else { $txtNote.Visibility = 'Collapsed' }
        $links = @()
        foreach ($k in $res.Links.Keys) { $links += [pscustomobject]@{ Name = [string]$k; Url = [string]$res.Links[$k] } }
        $productLinks.ItemsSource = $links
        $gridCves.ItemsSource = (ConvertTo-CveDialogRows -Result $res)
    }.GetNewClosure()

    $state.Result = Get-CveCachedResult -Row $Row
    & $render

    # Every link in the window is a Button with its https URL in Tag.
    $dlg.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        $src = $e.OriginalSource
        if ($src -is [System.Windows.Controls.Button] -and ([string]$src.Tag) -match '^https://') {
            $e.Handled = $true
            Open-ExternalUrl -Url ([string]$src.Tag)
        }
    })

    # Refresh runs the lookup through the background pipeline (rate-limit
    # waits included) and re-reads the cache when the run reports done; the
    # window is modeless, so the main window's overlay and Cancel stay usable.
    $btnRefresh.Add_Click({
        $cpe = [string]$Row.Cpe
        $from = [string]$Row.CurrentVersion
        $to = [string]$Row.LatestVersion
        if (Get-CveSkipReason -Cpe $cpe -Current $from -Latest $to) {
            $txtNote.Text = 'Refresh needs a CPE header and a ConfigMgr version or a latest version.'
            $txtNote.Visibility = 'Visible'
            return
        }
        # $script: inside a closure is the closure's own module, so the busy
        # check lives in a script function.
        if (Test-PipelineBusy) {
            $txtNote.Text = 'Another run is in progress. Wait for it to finish.'
            $txtNote.Visibility = 'Visible'
            return
        }
        $btnRefresh.IsEnabled = $false
        $txtNote.Text = 'Looking up...'
        $txtNote.Visibility = 'Visible'
        $ctx = Get-CveRunContext
        $ctx.ForceRefresh = $true
        # A closure made inside a closure copies the locals of this body
        # only, not the outer closure's captures; the captures are copied
        # to locals first.
        $stateRef = $state; $rowRef = $Row; $renderRef = $render; $noteRef = $txtNote; $btnRef = $btnRefresh
        $ctx.CveDone = {
            param($doneState)
            try {
                # The run's own result first (a cache write can fail after a
                # good lookup); the cache when the lookup itself failed.
                $fresh = $null
                if ($doneState -and $doneState.CveResults -and $doneState.CveResults.ContainsKey([string]$rowRef.Script)) { $fresh = $doneState.CveResults[[string]$rowRef.Script] }
                if ($fresh -and -not $fresh.Error) { $stateRef.Result = $fresh }
                else { $stateRef.Result = Get-CveCachedResult -Row $rowRef }
                & $renderRef
                if ([string]$rowRef.CveCount -eq 'Error') {
                    $noteRef.Text = 'The lookup failed; the last good result is shown. See the log.'
                    $noteRef.Visibility = 'Visible'
                }
            }
            catch { Add-LogLine -Message ('CVE window update failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
            finally { $btnRef.IsEnabled = $true }
        }.GetNewClosure()
        Invoke-MultiAppPipeline -Operation CheckCve -Rows @($Row) -Context $ctx
    }.GetNewClosure())

    $btnCopy.Add_Click({
        $ids = @()
        if ($state.Result) { foreach ($c in @($state.Result.Cves)) { $ids += [string]$c.Id } }
        try { [System.Windows.Clipboard]::SetText(($ids -join "`r`n")) } catch { <# Copying text to the clipboard is a best-effort convenience. #> $null = $_ }
        Add-LogLine -Message ('{0} CVE id(s) copied.' -f $ids.Count)
    }.GetNewClosure())

    $btnExport.Add_Click({
        try {
            $paths = Write-CveReportForRows -Rows @($Row)
            if ($paths.Html) { Start-Process $paths.Html }
        }
        catch { Add-LogLine -Message ('CVE report failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) }
    }.GetNewClosure())

    $btnClose.Add_Click({ $dlg.Close() }.GetNewClosure())

    if ($Probe) {
        & $Probe @{ Window = $dlg; Grid = $gridCves; Title = $txtTitle; Range = $txtRange; Note = $txtNote; ProductLinks = $productLinks; State = $state; Render = $render }
        return
    }
    $dlg.Owner = $Owner
    $dlg.Show()
}

# A click on a CVE count cell opens the detail window for that row.
$dataGrid.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($snd, $e)
        $src = $e.OriginalSource
        if (-not ($src -is [System.Windows.Controls.Button])) { return }
        if ([string]$src.Tag -ne 'CveLink') { return }
        $row = $src.DataContext
        if (-not $row -or -not $row.PSObject.Properties['Cpe']) { return }
        $e.Handled = $true
        Show-CveDetailDialog -Owner $window -Row $row
    })

$btnCheckCve.Add_Click({
    $selectedRows = Get-SelectedRows
    if ($selectedRows.Count -eq 0) {
        Add-LogLine -Message "No rows selected."
        return
    }
    $txtStatus.Text = "Checking CVEs for selected packagers..."
    Invoke-MultiAppPipeline -Operation CheckCve -Rows $selectedRows -Context (Get-CveRunContext)
})

$btnCveReport.Add_Click({
    $rows = Get-SelectedRows
    if ($rows.Count -eq 0) { $rows = @($script:PackagerData) }
    if ($rows.Count -eq 0) {
        Add-LogLine -Message "No rows to report."
        return
    }
    Show-BusyIndicator -Title 'Writing the CVE report' -Step ('Writing the report for {0} application(s)...' -f $rows.Count)
    try {
        $txtStatus.Text = "Writing the CVE report..."
        $paths = Write-CveReportForRows -Rows $rows
        $txtStatus.Text = ("CVE report written for {0} application(s)." -f $rows.Count)
        if ($paths.Html) { Start-Process $paths.Html }
    }
    catch {
        Add-LogLine -Message ('CVE report failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
        $txtStatus.Text = "CVE report failed."
    }
    finally { Hide-BusyIndicator }
})

# --- 5. One Click ---
# The button opens the plan window; the window starts the run through
# Invoke-MultiAppPipeline -Operation OneClick and writes the report.

function Get-AppLogFolder {
    param([string]$AppRoot = $PSScriptRoot)
    Join-Path $AppRoot 'Logs'
}

function Get-OneClickPlanApps {
    # The tracked applications as the plan module reads them, from the
    # visible grid rows.
    param([Parameter(Mandatory)][array]$Rows)
    $apps = foreach ($row in $Rows) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$row.Script)
        $cadence = 0
        $wsusUnsupported = $false
        try {
            $meta = Get-PackagerMetadata -Path ([string]$row.FullPath)
            if ($meta.UpdateCadenceDays) { $cadence = [int]$meta.UpdateCadenceDays }
            $wsusUnsupported = ($meta.WsusSupport -eq 'No')
        } catch { <# Reading packager metadata for the cadence and WSUS flag is optional; defaults apply. #> $null = $_ }
        [pscustomobject]@{
            Packager        = $base
            Application     = [string]$row.Application
            Vendor          = [string]$row.Vendor
            LatestVersion   = [string]$row.LatestVersion
            CurrentVersion  = [string]$row.CurrentVersion
            CadenceDays     = $cadence
            WsusUnsupported = $wsusUnsupported
            Row             = $row
        }
    }
    return @($apps)
}

function Show-OneClickPlanDialog {
    # The plan before the run, the progress grid during it, and the report
    # after it. Modeless: the pipeline timer on the main window drives the
    # progress through the OneClickTick and OneClickDone hooks.
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Action,
        [bool]$Force = $false,
        # Test hook: receives the window internals instead of showing the window.
        [scriptblock]$Probe = $null
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="One Click" Width="1280" Height="720" MinWidth="1100" MinHeight="480"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="16,12,16,8">
            <TextBlock x:Name="txtPlanSummary" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap"/>
            <TextBlock x:Name="txtPlanAction" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
            <TextBlock x:Name="txtPlanScope" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"/>
            <TextBlock x:Name="txtPlanBlocking" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Visibility="Collapsed"/>
            <TextBlock x:Name="txtPlanHint" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray.MouseOver}"
                       Text="Include and the destination boxes apply to this run only; Options, Updates and CVEs, One Click holds the saved values. Run asks nothing after it starts."/>
        </StackPanel>
        <DataGrid RowStyle="{DynamicResource GridRowStyle}" CellStyle="{DynamicResource GridCellStyle}" Grid.Row="1" x:Name="gridPlan" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column" Margin="16,0,16,8"
                  SelectionMode="Single" RowHeaderWidth="0">
            <DataGrid.Columns>
                <DataGridTemplateColumn MinWidth="95" x:Name="colInclude" Header="Include" Width="95">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Include, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" Tag="Include" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Application" Binding="{Binding Application}" Width="*" MinWidth="160" IsReadOnly="True"/>
                <DataGridTextColumn MinWidth="96" Header="Version" Binding="{Binding Version}" Width="96" IsReadOnly="True"/>
                <DataGridTemplateColumn MinWidth="119" x:Name="colMecm" Header="ConfigMgr" Width="119">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding ConfigMgr, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" Tag="ConfigMgr" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn MinWidth="78" x:Name="colWsus" Header="WSUS" Width="78">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding WSUS, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" Tag="WSUS" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn MinWidth="89" x:Name="colIntune" Header="Intune" Width="89">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Intune, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" Tag="Intune" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn MinWidth="145" x:Name="colLast" Header="Last Published" Binding="{Binding LastPublished}" Width="145" IsReadOnly="True"/>
                <DataGridTextColumn MinWidth="102" x:Name="colPlanned" Header="Planned" Binding="{Binding Planned}" Width="280" IsReadOnly="True"/>
                <DataGridTextColumn MinWidth="70" x:Name="colStep" Header="Step" Binding="{Binding Step}" Width="130" IsReadOnly="True" Visibility="Collapsed"/>
                <DataGridTextColumn x:Name="colResultMecm" Header="ConfigMgr Result" Binding="{Binding ResultConfigMgr}" Width="1.2*" MinWidth="168" IsReadOnly="True" Visibility="Collapsed"/>
                <DataGridTextColumn x:Name="colResultWsus" Header="WSUS Result" Binding="{Binding ResultWSUS}" Width="1.2*" MinWidth="180" IsReadOnly="True" Visibility="Collapsed"/>
                <DataGridTextColumn x:Name="colResultIntune" Header="Intune Result" Binding="{Binding ResultIntune}" Width="1.2*" MinWidth="180" IsReadOnly="True" Visibility="Collapsed"/>
                <DataGridTextColumn MinWidth="93" x:Name="colReason" Header="Reason" Binding="{Binding Reason}" Width="230" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
        <Border Grid.Row="2" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <DockPanel Margin="16,10,16,10">
                <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Close"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnOpenReport" DockPanel.Dock="Right" Content="Open Report" Margin="0,0,8,0" IsEnabled="False"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnHistory" DockPanel.Dock="Right" Content="Run History" Margin="0,0,8,0"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Opens the folder with the reports of earlier runs."/>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnRun" Content="Run" Margin="0,0,8,0"
                            Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <Button x:Name="btnPlanOnly" Content="Plan Only" Margin="0,0,16,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Writes the plan as a report without running anything."/>
                    <CheckBox x:Name="chkForce" Content="Force: ignore the cadence and publish again" FontSize="12" VerticalAlignment="Center"
                              Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $gridPlan = $dlg.FindName('gridPlan')
    $txtSummary = $dlg.FindName('txtPlanSummary')
    $txtAction = $dlg.FindName('txtPlanAction')
    $txtScope = $dlg.FindName('txtPlanScope')
    $txtBlocking = $dlg.FindName('txtPlanBlocking')
    $chkForce = $dlg.FindName('chkForce')
    $btnRun = $dlg.FindName('btnRun')
    $btnPlanOnly = $dlg.FindName('btnPlanOnly')
    $btnOpenReport = $dlg.FindName('btnOpenReport')
    $btnHistory = $dlg.FindName('btnHistory')
    $btnClose = $dlg.FindName('btnClose')
    $chkForce.IsChecked = $Force

    # A package ID or a not-supported message is wider than its column; a
    # cell that does not wrap cuts it off.
    foreach ($name in 'colLast', 'colPlanned', 'colResultMecm', 'colResultWsus', 'colResultIntune', 'colReason') {
        $col = $dlg.FindName($name)
        $wrap = [System.Windows.Style]::new([System.Windows.Controls.TextBlock], $col.ElementStyle)
        $wrap.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::Wrap))
        $col.ElementStyle = $wrap
    }

    # Closures made with GetNewClosure resolve $script: in their own module
    # scope, so the preferences object is captured here.
    $prefsRef = $script:Prefs
    $apps = Get-OneClickPlanApps -Rows $Rows
    $appByName = @{}
    foreach ($a in $apps) { $appByName[$a.Packager] = $a }
    $history = @{}
    try { $history = Read-PackagerHistory } catch { <# Reading packager history is optional; an empty history is the fallback. #> $null = $_ }
    $logFolder = Get-AppLogFolder
    $reportFolder = Get-OneClickReportFolder -LogFolder $logFolder
    $state = @{ Running = $false; ReportPath = ''; Started = $null }

    $planRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
    $gridPlan.ItemsSource = $planRows

    # The plan is recomputed from the boxes in the grid, on a copy of the
    # preferences that carries this run's selections only.
    $prefsForRun = {
        $copy = [pscustomobject]@{}
        foreach ($prop in $prefsRef.PSObject.Properties) { $copy | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force }
        $appFlow = [pscustomobject]@{}
        foreach ($prop in $prefsRef.AppFlow.PSObject.Properties) { $appFlow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force }
        $dest = [ordered]@{}
        foreach ($r in $planRows) { $dest[$r.Packager] = [pscustomobject]@{ ConfigMgr = [bool]$r.ConfigMgr; WSUS = [bool]$r.WSUS; Intune = [bool]$r.Intune } }
        $appFlow.Destinations = [pscustomobject]$dest
        $copy.AppFlow = $appFlow
        return $copy
    }.GetNewClosure()

    $refresh = {
        $prefs = $(if ($planRows.Count -gt 0) { & $prefsForRun } else { $prefsRef })
        $plan = @(Get-OneClickPlan -Apps $apps -Prefs $prefs -History $history -Action $Action -Force:($chkForce.IsChecked -eq $true))
        $existing = @{}
        foreach ($r in $planRows) { $existing[$r.Packager] = $r }
        foreach ($p in $plan) {
            $last = @()
            if ($p.LastConfigMgr) { $last += ('ConfigMgr ' + $p.LastConfigMgr) }
            if ($p.LastWSUS) { $last += ('WSUS ' + $p.LastWSUS) }
            if ($p.LastIntune) { $last += ('Intune ' + $p.LastIntune) }
            if ($existing.ContainsKey($p.Packager)) {
                # A box the operator cleared stays cleared; Include follows
                # the plan until the operator changes it.
                $r = $existing[$p.Packager]
                $r.Version = $p.Version; $r.Planned = $p.Planned; $r.Reason = $p.Reason; $r.LastPublished = ($last -join ', ')
                $r.PublishTo = @($p.PublishTo)
                $r.NotPublishable = @($p.NotPublishable)
                if (-not $r.IncludeTouched) { $r.Include = [bool]$p.Include }
            }
            else {
                $planRows.Add([pscustomobject]@{
                    Packager = $p.Packager; Application = $p.Application; Version = $p.Version
                    ConfigMgr = [bool]$p.ConfigMgr; WSUS = [bool]$p.WSUS; Intune = [bool]$p.Intune
                    LastPublished = ($last -join ', '); Planned = $p.Planned; Reason = $p.Reason
                    PublishTo = @($p.PublishTo); NotPublishable = @($p.NotPublishable); Include = [bool]$p.Include; IncludeTouched = $false
                    Step = ''; Outcome = ''; ResultConfigMgr = ''; ResultWSUS = ''; ResultIntune = ''; IdConfigMgr = ''; IdWSUS = ''; IdIntune = ''
                    DetailConfigMgr = ''; DetailWSUS = ''; DetailIntune = ''
                })
            }
        }
        foreach ($p in $plan) { if ($existing.ContainsKey($p.Packager)) { $p.Include = [bool]$existing[$p.Packager].Include } }
        $summary = Get-OneClickPlanSummary -Plan $plan -Prefs $prefs -Action $Action
        $txtSummary.Text = $summary.Line
        $txtAction.Text = ('Action: {0}. Existing ConfigMgr application: {1}.' -f $(switch ($Action) { 'Report' { 'check only' } 'Stage' { 'check and stage' } default { 'check, stage and publish' } }), $(if ([string]$prefsRef.AppFlow.OnExisting -eq 'Overwrite') { 'overwrite' } else { 'skip' }))
        $scopeParts = @(foreach ($d in (Get-OneClickDestinationNames)) { if (@($plan | Where-Object { $_.Include -and $_.PublishTo -contains $d }).Count -gt 0) { '{0}: {1}' -f $d, (Get-OneClickDestinationScope -Prefs $prefs -Destination $d) } })
        $txtScope.Text = $(if ($scopeParts.Count) { 'Scope. ' + ($scopeParts -join '. ') + '.' } else { '' })
        $txtScope.Visibility = $(if ($scopeParts.Count) { 'Visible' } else { 'Collapsed' })
        if ($summary.Blocking.Count -gt 0) { $txtBlocking.Text = 'Not ready: ' + ($summary.Blocking -join '; '); $txtBlocking.Visibility = 'Visible' } else { $txtBlocking.Visibility = 'Collapsed' }
        $btnRun.IsEnabled = (@($planRows | Where-Object { $_.Include }).Count -gt 0 -and -not $state.Running)
        try { $gridPlan.Items.Refresh() } catch { <# Refreshing the grid view is cosmetic. #> $null = $_ }
    }.GetNewClosure()

    # A destination box or Include changed: recompute the plan. Click fires
    # only for the operator; Checked also fires for every box that
    # Items.Refresh regenerates, and a refresh from it never ends.
    $boxChanged = [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        if ($state.Running) { return }
        $box = $e.OriginalSource
        if ($box -is [System.Windows.Controls.CheckBox] -and $box.DataContext -and $box.DataContext.PSObject.Properties['Packager']) {
            if ([string]$box.Tag -eq 'Include') { $box.DataContext.IncludeTouched = $true }
            & $refresh
        }
    }.GetNewClosure()
    $gridPlan.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $boxChanged)
    $chkForce.Add_Click({ & $refresh }.GetNewClosure())

    $writeReport = {
        param($Ended, $Canceled, [string]$Label = 'One Click report')
        $used = @($planRows | Where-Object { $_.Include })
        $scope = [ordered]@{}
        if (@($used | Where-Object { $_.ConfigMgr }).Count) { $scope.ConfigMgr = Get-OneClickDestinationScope -Prefs $prefsRef -Destination 'ConfigMgr' }
        if (@($used | Where-Object { $_.WSUS }).Count) { $scope.WSUS = Get-OneClickDestinationScope -Prefs $prefsRef -Destination 'WSUS' }
        if (@($used | Where-Object { $_.Intune }).Count) { $scope.Intune = Get-OneClickDestinationScope -Prefs $prefsRef -Destination 'Intune' }
        $runInfo = [pscustomobject]@{
            Started  = $state.Started
            Ended    = $Ended
            Operator = ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
            Computer = $env:COMPUTERNAME
            Action   = $(switch ($Action) { 'Report' { 'Report only' } 'Stage' { 'Stage' } default { 'Stage and Publish' } })
            Force    = ($chkForce.IsChecked -eq $true)
            Canceled = [bool]$Canceled
            Scope    = [pscustomobject]$scope
        }
        $reportRows = @($planRows | Where-Object { $_.Include } | ForEach-Object {
            $r = $_
            [pscustomobject]@{
                Packager = $r.Packager; Application = $r.Application; Version = $r.Version
                Outcome = $(if ($r.Outcome) { $r.Outcome } elseif ($Canceled) { 'Skipped' } else { 'Planned' })
                Reason = $(if ($r.Outcome -or -not $Canceled) { $r.Reason } else { 'run canceled before this application' })
                ResultConfigMgr = $r.ResultConfigMgr; IdConfigMgr = $r.IdConfigMgr
                ResultWSUS = $r.ResultWSUS; IdWSUS = $r.IdWSUS
                ResultIntune = $r.ResultIntune; IdIntune = $r.IdIntune
                DetailConfigMgr = $r.DetailConfigMgr; DetailWSUS = $r.DetailWSUS; DetailIntune = $r.DetailIntune
            }
        })
        $paths = Write-OneClickReport -Run $runInfo -Rows $reportRows -Folder $reportFolder
        $state.ReportPath = $paths.MarkdownPath
        $btnOpenReport.IsEnabled = $true
        Add-LogLine -Message ($Label + ': ' + $paths.MarkdownPath)
        return $paths
    }.GetNewClosure()

    $btnPlanOnly.Add_Click({
        $state.Started = Get-Date
        $paths = & $writeReport (Get-Date) $false 'One Click plan'
        $txtSummary.Text = 'Plan written: ' + $paths.MarkdownPath
    }.GetNewClosure())

    $btnRun.Add_Click({
        $included = @($planRows | Where-Object { $_.Include })
        if ($included.Count -eq 0) { return }
        $rowsToRun = @($included | ForEach-Object { $appByName[$_.Packager].Row })
        if ($Action -in @('Stage', 'StageAndPackage')) {
            $rowsToRun = @(Confirm-LocalSourceFolders -Rows $rowsToRun)
            $kept = [System.Collections.Generic.HashSet[string]]::new([string[]]@($rowsToRun | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension([string]$_.Script) }), [System.StringComparer]::OrdinalIgnoreCase)
            foreach ($r in $included) { if (-not $kept.Contains($r.Packager)) { $r.Include = $false; $r.IncludeTouched = $true; $r.Reason = 'no local source folder' } }
            $included = @($planRows | Where-Object { $_.Include })
            if ($included.Count -eq 0) { $txtSummary.Text = 'Nothing to run.'; return }
        }

        $state.Running = $true
        $state.Started = Get-Date
        $btnRun.IsEnabled = $false; $btnPlanOnly.IsEnabled = $false; $chkForce.IsEnabled = $false; $btnClose.IsEnabled = $false
        foreach ($name in 'colInclude', 'colMecm', 'colWsus', 'colIntune') { $dlg.FindName($name).IsReadOnly = $true }
        # The result columns take the place of the destination boxes; with
        # both shown, a result column is narrower than one package ID.
        foreach ($name in 'colMecm', 'colWsus', 'colIntune') { $dlg.FindName($name).Visibility = 'Collapsed' }
        $dlg.FindName('colReason').Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
        $dlg.FindName('colStep').Visibility = 'Visible'
        $dlg.FindName('colResultMecm').Visibility = $(if (@($included | Where-Object { $_.ConfigMgr }).Count) { 'Visible' } else { 'Collapsed' })
        $dlg.FindName('colResultWsus').Visibility = $(if (@($included | Where-Object { $_.WSUS }).Count) { 'Visible' } else { 'Collapsed' })
        $dlg.FindName('colResultIntune').Visibility = $(if (@($included | Where-Object { $_.Intune }).Count) { 'Visible' } else { 'Collapsed' })
        $dlg.FindName('colPlanned').Visibility = 'Collapsed'
        $dlg.FindName('colLast').Visibility = 'Collapsed'
        $txtSummary.Text = ('Running: {0} application(s)' -f $included.Count)

        # GetNewClosure captures local variables only; the references this
        # handler's own closure holds are copied into locals so that the tick
        # and done hooks below see them.
        $state = $state; $planRows = $planRows; $gridPlan = $gridPlan
        $writeReport = $writeReport; $txtSummary = $txtSummary; $btnClose = $btnClose

        # Progress records the background loop writes and the tick copies
        # back into the grid rows.
        $progress = [hashtable]::Synchronized(@{})
        $planByApp = @{}
        foreach ($r in $included) {
            $progress[$r.Packager] = [hashtable]::Synchronized(@{ Step = ''; Outcome = ''; Reason = ''; Version = ''; ResultConfigMgr = ''; ResultWSUS = ''; ResultIntune = ''; IdConfigMgr = ''; IdWSUS = ''; IdIntune = ''; DetailConfigMgr = ''; DetailWSUS = ''; DetailIntune = '' })
            $selected = @()
            if ($r.ConfigMgr) { $selected += 'ConfigMgr' }
            if ($r.WSUS) { $selected += 'WSUS' }
            if ($r.Intune) { $selected += 'Intune' }
            # A destination the plan reported as not ready or not supported
            # stays out of the run; the check box alone does not publish.
            $planByApp[$r.Packager] = @{ Destinations = @(Select-OneClickRunDestinations -Selected $selected -NotPublishable @($r.NotPublishable)) }
            $r.Step = 'Pending'
        }

        $tick = {
            foreach ($r in $planRows) {
                if (-not $progress.ContainsKey($r.Packager)) { continue }
                $pr = $progress[$r.Packager]
                foreach ($f in 'Step', 'Outcome', 'Reason', 'ResultConfigMgr', 'ResultWSUS', 'ResultIntune', 'IdConfigMgr', 'IdWSUS', 'IdIntune', 'DetailConfigMgr', 'DetailWSUS', 'DetailIntune') {
                    if ($pr[$f] -and $r.$f -ne $pr[$f]) { $r.$f = $pr[$f] }
                }
                if ($pr['Version'] -and $r.Version -ne $pr['Version']) { $r.Version = $pr['Version'] }
            }
            try { $gridPlan.Items.Refresh() } catch { <# Refreshing the grid view is cosmetic. #> $null = $_ }
        }.GetNewClosure()

        $done = {
            param($DoneState)
            & $tick
            $state.Running = $false
            $canceled = [bool]($DoneState -and $DoneState.Canceled)
            $paths = & $writeReport (Get-Date) $canceled
            $published = @($planRows | Where-Object { $_.Outcome -eq 'Published' }).Count
            $failed = @($planRows | Where-Object { $_.Outcome -eq 'Failed' }).Count
            $txtSummary.Text = ('{0}: {1} published, {2} failed. Report: {3}' -f $(if ($canceled) { 'Canceled' } else { 'Complete' }), $published, $failed, (Split-Path -Leaf $paths.MarkdownPath))
            $btnClose.IsEnabled = $true
        }.GetNewClosure()

        $siteCodeValue = [string]$prefsRef.SiteCode
        if ([string]::IsNullOrWhiteSpace($siteCodeValue)) { $siteCodeValue = 'MCM' }
        $wantsMecm = @($included | Where-Object { $_.ConfigMgr }).Count -gt 0
        Add-LogSeparator
        Add-LogLine -Message ("One Click: {0} app(s), action={1}{2}" -f $included.Count, $Action, $(if ($chkForce.IsChecked) { ', force=on' } else { '' }))
        $txtStatus.Text = ("One Click: {0} app(s)..." -f $included.Count)
        Invoke-MultiAppPipeline -Operation OneClick -Rows $rowsToRun -Context @{
            SiteCode             = $siteCodeValue
            ProviderMachineName  = $prefsRef.ProviderMachineName
            Action               = $Action
            ForceFlag            = ($chkForce.IsChecked -eq $true)
            OnExisting           = [string]$prefsRef.AppFlow.OnExisting
            Comment              = $txtComment.Text.Trim()
            FileShareRoot        = $prefsRef.FileShareRoot
            ContentLayout        = $prefsRef.ContentLayout
            DownloadRoot         = $prefsRef.DownloadRoot
            M365Channel          = $prefsRef.M365Channel
            M365DeployMode       = $prefsRef.M365DeployMode
            EstimatedRuntimeMins = $prefsRef.EstimatedRuntimeMins
            MaximumRuntimeMins   = $prefsRef.MaximumRuntimeMins
            AdminUiFound         = $prefsRef.DetectedTools.ConfigMgrConsole.Found
            LogFolder            = $logFolder
            SevenZipPath         = Get-SevenZipPathForContext
            IntuneWinCreate      = ($wantsMecm -and -not [string]::IsNullOrWhiteSpace((Get-IntuneWinToolPathForContext)))
            IntuneWinToolPath    = Get-IntuneWinToolPathForContext
            RequirementsByApp    = Get-RequirementsMapForContext
            VariantsByApp        = Get-VariantsMapForContext
            CommandsByApp        = Get-CommandsMapForContext
            InstallModesByApp    = Get-InstallModesMapForContext
            TitleModesByApp      = Get-TitleModesMapForContext
            DefaultTitleMode     = Get-DefaultTitleModeForContext
            IntunePublishConfig  = Get-IntunePublishConfigForContext -Target 'IntuneOnly'
            WsusPublishConfig    = Get-WsusPublishConfigForContext -Target 'WSUSOnly'
            DeploymentTarget     = 'MECM'
            RunPlanByApp         = Get-WorkbenchRunPlanForContext -Rows $rowsToRun -Target 'MECM'
            SigningJson          = Get-WorkbenchSigningPolicyJson
            SigningDigest        = Get-WorkbenchSigningPolicyDigest
            OneClickPlan         = $planByApp
            OneClickRows         = $progress
            OneClickTick         = $tick
            OneClickDone         = $done
        }
    }.GetNewClosure())

    $btnOpenReport.Add_Click({ if ($state.ReportPath -and (Test-Path -LiteralPath $state.ReportPath)) { try { Start-Process $state.ReportPath } catch { <# Opening a link or file in the shell is a best-effort convenience. #> $null = $_ } } }.GetNewClosure())
    $btnHistory.Add_Click({
        if (-not (Test-Path -LiteralPath $reportFolder)) { [void](New-Item -ItemType Directory -Path $reportFolder -Force) }
        try { Start-Process explorer.exe -ArgumentList ('"' + $reportFolder + '"') } catch { <# Opening a link or file in the shell is a best-effort convenience. #> $null = $_ }
    }.GetNewClosure())
    $btnClose.Add_Click({ if (-not $state.Running) { $dlg.Close() } }.GetNewClosure())
    $dlg.Add_Closing({ param($sender, $e) if ($state.Running) { $e.Cancel = $true } }.GetNewClosure())

    & $refresh
    if ($Probe) {
        & $Probe ([pscustomobject]@{ Dialog = $dlg; Rows = $planRows; Refresh = $refresh; Summary = $txtSummary; Blocking = $txtBlocking; Scope = $txtScope; Run = $btnRun; Force = $chkForce; WriteReport = $writeReport; State = $state })
        return
    }
    $dlg.Owner = $Owner
    $dlg.Show()
}

$btnFullRun.Add_Click({
    $action = [string]$script:Prefs.AppFlow.Action
    $trackedBases = @($script:Prefs.AppFlow.Tracked)
    if ($trackedBases.Count -eq 0) {
        Add-LogLine -Message "No apps are tracked for One Click. Open Options, Updates and CVEs, One Click."
        $txtStatus.Text = "No apps tracked."
        [void](Show-ThemedMessage -Owner $window -Title 'One Click Not Configured' `
            -Message "No apps are tracked yet.`n`nOpen Options (sidebar) and select Updates and CVEs, One Click, then choose which packagers to include, pick an action (Report / Stage / Stage and Publish) and the destinations, and click OK." `
            -Buttons OK -Icon Info)
        return
    }
    if ($action -in @('Stage','StageAndPackage') -and [string]::IsNullOrWhiteSpace([string]$script:Prefs.DownloadRoot)) {
        Add-LogLine -Message ("Download root is required for action '{0}'. Set it in Options, Packaging, Folders." -f $action)
        $txtStatus.Text = "Download root is required."
        return
    }
    $trackedSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$trackedBases, [System.StringComparer]::OrdinalIgnoreCase)
    $rows = @($script:PackagerData | Where-Object { $trackedSet.Contains([System.IO.Path]::GetFileNameWithoutExtension([string]$_.Script)) })
    if ($rows.Count -eq 0) {
        Add-LogLine -Message ("No tracked apps are visible in the grid. Check Product Filter.")
        $txtStatus.Text = "No visible tracked apps."
        return
    }
    Show-OneClickPlanDialog -Owner $window -Rows $rows -Action $action -Force ([bool]$script:Prefs.AppFlow.ForceOnLaunch)
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
    } catch { <# Reading runtime minutes from preferences is optional; the defaults stay. #> $null = $_ }

    $resolved = $null
    try { $resolved = Resolve-EffectiveSettings -GlobalDefaults $globals -BaseManifest $null -Profile $null } catch { <# Reading runtime minutes from preferences is optional; the defaults stay. #> $null = $_ }

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
        $results.Add('Parser failure: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="18,16,18,14">
        <TextBlock x:Name="txtUnsaved" TextWrapping="Wrap" FontSize="12" Margin="0,0,0,16"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnUnsavedSave" Content="Save" Margin="0,0,8,0"
                    Style="{DynamicResource DialogAccentButton}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnUnsavedDiscard" Content="Discard" Margin="0,0,8,0"
                    Style="{DynamicResource DialogButton}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnUnsavedCancel" Content="Cancel" IsCancel="True"
                    Style="{DynamicResource DialogButton}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@
    [xml]$dx = Expand-SharedStyle $dlgXaml
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
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="18,16,18,14">
        <TextBlock x:Name="txtNamePrompt" FontSize="12" Margin="0,0,0,8"/>
        <TextBox x:Name="txtNameValue" FontSize="12" Height="28" Margin="0,0,0,16"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnNameOk" Content="OK" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource DialogAccentButton}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnNameCancel" Content="Cancel" IsCancel="True"
                    Style="{DynamicResource DialogButton}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@
    [xml]$dx = Expand-SharedStyle $dlgXaml
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
    if ([string]$Target -notin (@('ContentOnly') + (Get-DeploymentTargetNames))) { $Target = 'MECM' }
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
            } catch { <# Reading the display name from the profile is optional; the fallback title stays. #> $null = $_ }
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
            Add-LogLine -Message ('Run snapshot not created for {0}: {1}' -f $base, (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
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
            Add-LogLine -Message ('Legacy preference migration failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.'))
        }
    }

    $xamlPath = Get-WorkbenchXamlPath
    if (-not (Test-Path -LiteralPath $xamlPath)) {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Application Workbench' `
            -Message ('WorkbenchWindow.xaml was not found at ' + $xamlPath + '.') -Buttons OK -Icon Error)
        return
    }
    [xml]$wbXaml = Expand-SharedStyle (Get-Content -LiteralPath $xamlPath -Raw)
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
    $btnPublishMecm   = & $ctl 'btnPublishMecm'
    $btnPublishIntune = & $ctl 'btnPublishIntune'
    $btnPublishWsus   = & $ctl 'btnPublishWsus'

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
            try { [void](Save-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId -Draft ([pscustomobject]$wb.Profile)) } catch { <# Saving a draft copy of the profile is best-effort autosave. #> $null = $_ }
        }
    }

    $refreshFieldIndicators = {
        foreach ($d in $fieldDescriptors) {
            if (-not $d.SourceLabel) { continue }
            $inheritedValue = $null
            try { $inheritedValue = $wb.Inherited.($d.InheritedKey) } catch { <# Reading an inherited value for an indicator is optional; an empty value is shown. #> $null = $_ }
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
                    Version  = [string]$r.SoftwareVersion
                    Revision = [string]$r.ProfileRevision
                    Result   = (Get-WorkbenchBuildResultText -Record $r)
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
            try { $inheritedValue = [string]$wb.Inherited.($d.InheritedKey) } catch { <# Reading an inherited value for an indicator is optional; an empty value is shown. #> $null = $_ }
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
                try { $downloadRoot = [string]$script:Prefs.DownloadRoot } catch { <# Reading the download root for the icon preview is optional. #> $null = $_ }
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
                } catch { <# Loading the icon preview image is cosmetic. #> $null = $_ }
            }
            $txtAppNote.Text = 'The displayed title is separate from the stable identity: renaming a profile never renames an existing site application, and the output folder keeps the packager name.'

            $cboReqTemplate.Items.Clear()
            try {
                foreach ($c in @((Get-ConditionTemplates).Conditions)) { [void]$cboReqTemplate.Items.Add([string]$c.Id) }
            } catch { <# Filling the optional condition template list; an empty list is acceptable. #> $null = $_ }
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
            try { $draft = Get-Draft -ApplicationId ([string]$application.ApplicationId) -ProfileId $wb.ProfileId } catch { <# Restoring a saved draft is optional; the stored profile stays. #> $null = $_ }
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
                $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for This Profile' -Value 'Managed'
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
            [void](Show-ThemedMessage -Owner $win -Title 'Save Failed' -Message (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') -Buttons OK -Icon Error)
            return $false
        }
        $wb.Profile = ConvertTo-WorkbenchUiHashtable -InputObject $saved
        $wb.ProfileId = [string]$wb.Profile['ProfileId']
        try { [void](Remove-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId) } catch { <# Removing the draft file is cleanup after a save or discard. #> $null = $_ }
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
                try { [void](Remove-Draft -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId $wb.ProfileId) } catch { <# Removing the draft file is cleanup after a save or discard. #> $null = $_ }
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
            $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for the New Profile' -Value ([string]$wb.Profile.Name + ' copy')
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
        $name = Show-WorkbenchNameDialog -Owner $win -Title 'New Name for This Profile' -Value ([string]$wb.Profile.Name)
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
        $dlg.Title = 'Add Source File'
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
        $dlg.Title = 'Replace Source File'
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
        $dlg.Title = 'Choose Icon'
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
        $name = Show-WorkbenchNameDialog -Owner $win -Title 'Name for the New Profile' -Value ([string]$wb.Profile.Name + ' copy')
        if ($name) { [void](& $saveProfile $name) }
    })
    $btnValidate.Add_Click({
        [void](& $validateProfile)
        $lstSections.SelectedIndex = 6
        & $setStatus 'Validation inspected the profile without running anything. Results are per target.'
    })

    $runFromWorkbench = {
        param([string]$operation, [string]$target = '')
        if ($wb.Dirty) {
            if (-not (& $confirmDiscard 'Save the profile before this run? A run uses the saved profile revision.')) { return }
        }
        $findings = @(& $validateProfile)
        $blocking = @($findings | Where-Object { $_.Severity -eq 'Blocking' -and ($target -eq 'IntuneOnly' -or [string]$_.Code -notlike '*INTUNE*') })
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
        # The pipeline builds the application's active profile. The window can
        # show a profile that was selected and never saved again, so the
        # displayed profile becomes the active one before the run. This runs
        # while the window is open so a failure is shown in it.
        try { [void](Set-ActiveProfile -ApplicationId ([string]$wb.Application.ApplicationId) -ProfileId ([string]$wb.ProfileId)) }
        catch {
            [void](Show-ThemedMessage -Owner $win -Title $operation `
                -Message ('The profile ' + [string]$wb.Profile.Name + ' could not be made the active profile, so nothing ran: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')) -Buttons OK -Icon Warning)
            return
        }
        $win.Close()
        $row = @($script:PackagerData | Where-Object { [string]$_.FullPath -eq [string]$wb.Application.ScriptPath })
        if ($row.Count -eq 0) {
            $cliHint = ('.\Invoke-AppPackagerBuild.ps1 -Application {0} -Profile {1} -{2}{3}' -f [string]$wb.Application.ApplicationId, [string]$wb.ProfileId, $operation, $(if ($target -and $target -ne 'MECM') { ' -Stage -Target ' + $target } else { '' }))
            Add-LogLine -Message ('Workbench run skipped: {0} is not in the current grid. Build it from the command line: {1}' -f [string]$wb.Application.DisplayName, $cliHint)
            [void](Show-ThemedMessage -Owner $win -Title $operation `
                -Message ('This application is not in the main grid, so it cannot run through the grid pipeline. Build it from the command line:' + "`r`n`r`n" + $cliHint) -Buttons OK -Icon Info)
            return
        }
        Add-LogLine -Message ('Workbench {0}: {1} (profile {2}, revision {3})' -f $operation, [string]$wb.Application.DisplayName, [string]$wb.Profile.Name, [string]$wb.Profile.Revision)
        Invoke-WorkbenchRun -Operation $operation -Rows $row -BuildId $(if ($operation -eq 'Package') { $selectedBuild } else { '' }) -Target $target
    }
    $btnStage.Add_Click({ & $runFromWorkbench 'Stage' })
    $btnPublishMecm.Add_Click({ & $runFromWorkbench 'Package' 'MECM' })
    $btnPublishIntune.Add_Click({ & $runFromWorkbench 'Package' 'IntuneOnly' })
    $btnPublishWsus.Add_Click({ & $runFromWorkbench 'Package' 'WSUSOnly' })

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
        [AllowEmptyString()][string]$BuildId = '',
        [AllowEmptyString()][string]$Target = ''
    )
    $target = if ($Target) { $Target } else { [string]$script:Prefs.Intune.DeploymentTarget }
    $toolPath = Get-IntuneWinToolPathForContext
    $context = @{
        DownloadRoot      = $script:Prefs.DownloadRoot
        M365Channel       = $script:Prefs.M365Channel
        M365DeployMode    = $script:Prefs.M365DeployMode
        LogFolder         = Join-Path $PSScriptRoot 'Logs'
        SevenZipPath      = Get-SevenZipPathForContext
        IntuneWinToolPath = $toolPath
        RunPlanByApp      = Get-WorkbenchRunPlanForContext -Rows $Rows -Target $target
        SigningJson       = Get-WorkbenchSigningPolicyJson
        SigningDigest     = Get-WorkbenchSigningPolicyDigest
    }
    if ($Operation -eq 'Package') {
        $context['SiteCode']             = $script:Prefs.SiteCode
        $context['ProviderMachineName']  = $script:Prefs.ProviderMachineName
        $context['Comment']              = $txtComment.Text.Trim()
        $context['FileShareRoot']        = $script:Prefs.FileShareRoot
        $context['ContentLayout']        = $script:Prefs.ContentLayout
        $context['EstimatedRuntimeMins'] = $script:Prefs.EstimatedRuntimeMins
        $context['MaximumRuntimeMins']   = $script:Prefs.MaximumRuntimeMins
        $context['IntuneWinCreate']      = (-not (Test-DeploymentTargetSkipsSite -DeploymentTarget $target) -and -not [string]::IsNullOrWhiteSpace($toolPath))
        $context['RequirementsByApp']    = Get-RequirementsMapForContext
        $context['VariantsByApp']        = Get-VariantsMapForContext
        $context['CommandsByApp']        = Get-CommandsMapForContext
        $context['InstallModesByApp']    = Get-InstallModesMapForContext
        $context['TitleModesByApp']      = Get-TitleModesMapForContext
        $context['DefaultTitleMode']     = Get-DefaultTitleModeForContext
        $context['IntunePublishConfig']  = Get-IntunePublishConfigForContext -Target $target
        $context['WsusPublishConfig']    = Get-WsusPublishConfigForContext -Target $target
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
        [AllowEmptyString()][string]$BuildId = '',
        [AllowEmptyString()][string]$Target = ''
    )
    if ($Operation -eq 'Stage') {
        $Rows = @(Confirm-LocalSourceFolders -Rows $Rows)
        if ($Rows.Count -eq 0) { return }
    }
    if ($Operation -eq 'Package') {
        if ($Target -eq 'IntuneOnly' -and -not (Get-IntunePublishConfigForContext -Target $Target)) {
            Add-LogLine -Message "Publish requires Intune credentials. Set Tenant ID, Client ID and Client secret in Options, Connections, Intune."
            return
        }
        if (-not (Confirm-WsusPublishPrerequisites -Target $Target)) { return }
    }
    $context = New-WorkbenchPipelineContext -Operation $Operation -Rows $Rows -BuildId $BuildId -Target $Target
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

        <TextBlock Text="Sign" FontSize="15" Margin="0,0,0,6"/>
        <CheckBox x:Name="chkSignDetection" FontSize="12" Margin="0,0,0,4"
                  Content="Sign detection scripts" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Sign the generated and custom detection scripts."/>
        <CheckBox x:Name="chkSignRequirements" FontSize="12" Margin="0,0,0,4"
                  Content="Sign requirement scripts and script global conditions" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Sign the script-based requirement rules."/>
        <CheckBox x:Name="chkSignDeployment" FontSize="12" Margin="0,0,0,12"
                  Content="Sign install and uninstall PowerShell scripts" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                  ToolTip="Sign the install and uninstall scripts and their launchers."/>

        <TextBlock Text="Require valid signatures" FontSize="15" Margin="0,0,0,6"/>
        <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,0,0,6"
                   Text="A publishing constraint, not a change to client execution policy. When a required category fails verification the build stops; there is no unsigned fallback."/>
        <CheckBox x:Name="chkRequireDetection" FontSize="12" Margin="0,0,0,4"
                  Content="Require valid detection signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        <CheckBox x:Name="chkRequireRequirements" FontSize="12" Margin="0,0,0,4"
                  Content="Require valid requirement signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        <CheckBox x:Name="chkRequireDeployment" FontSize="12" Margin="0,0,0,12"
                  Content="Require valid deployment script signatures" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

        <TextBlock Text="Signing identity" FontSize="15" Margin="0,0,0,6"/>
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

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Certificate store" FontWeight="Bold" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}"  Grid.Row="0" Grid.Column="1" x:Name="cboSignStore" FontSize="12" Height="28" Margin="0,0,0,6"
                       ToolTip="Use CurrentUser\My, or LocalMachine\My with access to the private key."/>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Certificate" FontWeight="Bold" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}"  Grid.Row="1" Grid.Column="1" x:Name="cboSignCertificate" FontSize="12" Height="28" Margin="0,0,0,6"
                       ToolTip="Select the certificate by its thumbprint."/>
            <Button    Grid.Row="1" Grid.Column="2" x:Name="btnSignRefresh" Content="Refresh" Margin="8,0,0,6"
                       Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

            <TextBlock Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtSignCertDetail" FontSize="11" TextWrapping="Wrap"
                       Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,0,0,8"/>

            <TextBlock Grid.Row="3" Grid.Column="0" Text="Timestamp server" FontWeight="Bold" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtSignTimestamp" FontSize="12" Height="28" VerticalContentAlignment="Center" Margin="0,0,0,6"
                       Controls:TextBoxHelper.Watermark="none"
                       ToolTip="A timestamp keeps a signature valid after the certificate expires."/>
            <CheckBox  Grid.Row="4" Grid.Column="1" x:Name="chkSignTimestampRequired" FontSize="12" Margin="0,0,0,6"
                       Content="Timestamp required" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                       ToolTip="Fail the build when the timestamp server does not answer."/>
        </Grid>

        <StackPanel Orientation="Horizontal" Margin="0,4,0,8">
            <Button x:Name="btnSignTest" Content="Test Signing Configuration" Margin="0,0,10,0"
                    Style="{DynamicResource InlineButtonStyle}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                    ToolTip="Sign and verify a temporary file with the selected certificate."/>
        </StackPanel>
        <TextBlock x:Name="txtSignTestResult" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,10"/>

        <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                   Text="Private keys are never exported and never reach profiles, manifests, logs or command lines. A key held on a hardware token may require a PIN and can prevent unattended runs."/>
    </StackPanel>
</ScrollViewer>
'@

    [xml]$xml = Expand-SharedStyle $xaml
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
        catch { $txtSignCertDetail.Text = 'Certificate enumeration failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
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
            $txtSignTestResult.Text = 'Test failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
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

function Get-WsusClassificationLabel {
    param([Parameter(Mandatory)][string]$Name)
    switch ($Name) {
        'SecurityUpdates' { return 'Security Updates' }
        'CriticalUpdates' { return 'Critical Updates' }
        'FeaturePacks'    { return 'Feature Packs' }
        'ServicePacks'    { return 'Service Packs' }
        'UpdateRollups'   { return 'Update Rollups' }
        default           { return $Name }
    }
}

function Show-WsusPasswordPrompt {
    # Returns the SecureString typed by the operator, or $null on cancel.
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][string]$Prompt)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Certificate Password" Width="440" SizeToContent="Height"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal" ResizeMode="NoResize"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <StackPanel Margin="20,16,20,16">
        <TextBlock x:Name="txtPrompt" TextWrapping="Wrap" FontSize="12" Margin="0,0,0,10"/>
        <PasswordBox x:Name="pwdValue" FontSize="13" Margin="0,0,0,14"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnOk" Content="OK" Margin="0,0,8,0" IsDefault="True" Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnCancel" Content="Cancel" IsCancel="True" Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </StackPanel>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner
    $dlg.FindName('txtPrompt').Text = $Prompt
    $pwdValue = $dlg.FindName('pwdValue')
    $state = @{ Value = $null }
    $dlg.FindName('btnOk').Add_Click({ $state.Value = $pwdValue.SecurePassword.Copy(); $dlg.Close() }.GetNewClosure())
    $dlg.FindName('btnCancel').Add_Click({ $dlg.Close() }.GetNewClosure())
    [void]$pwdValue.Focus()
    [void]$dlg.ShowDialog()
    return $state.Value
}

function Show-WsusPublishedUpdatesDialog {
    # Lists locally published updates on the server and applies approval,
    # decline and expiry to the selection.
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][hashtable]$Settings)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="WSUS Updates" Width="1240" Height="620" MinWidth="1040" MinHeight="420"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <DockPanel Grid.Row="0" Margin="16,12,16,8">
            <CheckBox x:Name="chkOtherPublishers" DockPanel.Dock="Right" Content="Show updates from other publishers" FontSize="12"
                      Controls:ControlsHelper.ContentCharacterCasing="Normal" VerticalAlignment="Center"
                      ToolTip="Lists every locally published update on the server, not only the ones AppPackager published."/>
            <TextBlock x:Name="txtSummary" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center"/>
        </DockPanel>
        <DataGrid RowStyle="{DynamicResource GridRowStyle}" CellStyle="{DynamicResource GridCellStyle}" Grid.Row="1" x:Name="gridUpdates" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended"
                  CanUserAddRows="False" HeadersVisibility="Column" Margin="16,0,16,8">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*" MinWidth="280"/>
                <DataGridTextColumn MinWidth="71" Header="Type" Binding="{Binding PackageType}" Width="71"/>
                <DataGridTextColumn MinWidth="97" Header="Created" Binding="{Binding CreatedText}" Width="125"/>
                <DataGridTextColumn MinWidth="138" Header="Approved For" Binding="{Binding ApprovedGroups}" Width="138"/>
                <DataGridTextColumn MinWidth="103" Header="Declined" Binding="{Binding Declined}" Width="103"/>
                <DataGridTextColumn MinWidth="121" Header="Superseded" Binding="{Binding Superseded}" Width="121"/>
                <DataGridTextColumn MinWidth="93" Header="Expired" Binding="{Binding Expired}" Width="93"/>
                <DataGridTextColumn MinWidth="116" Header="Package ID" Binding="{Binding PackageId}" Width="240"/>
            </DataGrid.Columns>
        </DataGrid>
        <Border Grid.Row="2" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <DockPanel Margin="16,10,16,10">
                <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Close" IsCancel="True"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnCatalog" DockPanel.Dock="Right" Content="Catalog Import" Margin="0,0,8,0"
                        Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                        ToolTip="Import Microsoft updates that do not synchronize automatically, by catalog update ID or link."/>
                <StackPanel Orientation="Horizontal">
                    <ComboBox ItemContainerStyle="{DynamicResource ThemedComboBoxItemStyle}" x:Name="cboGroup" IsEditable="True" Width="200" Height="30" FontSize="12" Margin="0,0,8,0"
                              ToolTip="The computer group that Approve uses."/>
                    <Button x:Name="btnApprove" Content="Approve for Group" Margin="0,0,8,0"
                            Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                    <Button x:Name="btnDecline" Content="Decline" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Decline the update so that no client gets it."/>
                    <Button x:Name="btnExpire" Content="Expire" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Expire the update on the server for good."/>
                    <Button x:Name="btnRemove" Content="Remove" Margin="0,0,8,0"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"
                            ToolTip="Decline the update and delete it from the WSUS database."/>
                    <Button x:Name="btnRefresh" Content="Refresh"
                            Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $gridUpdates = $dlg.FindName('gridUpdates')
    $txtSummary  = $dlg.FindName('txtSummary')
    $chkOther    = $dlg.FindName('chkOtherPublishers')
    $cboGroup    = $dlg.FindName('cboGroup')
    $cboGroup.Text = [string]$Settings.ApprovalGroup

    $load = {
        Set-DialogBusyStatus -Dialog $dlg -Control $txtSummary -Text ('Reading the updates from {0}...' -f [string]$Settings.ServerName)
        try {
            $rows = @(Get-WsusPublishedUpdates -Settings $Settings -IncludeOtherPublishers:($chkOther.IsChecked -eq $true))
            $gridUpdates.ItemsSource = $rows
            $txtSummary.Text = ('{0} update(s) on {1}.' -f $rows.Count, [string]$Settings.ServerName)
            if ($cboGroup.Items.Count -eq 0) {
                try { foreach ($name in @(Get-WsusComputerGroupNames -Settings $Settings)) { [void]$cboGroup.Items.Add($name) } }
                catch { $txtSummary.Text += ' Computer groups could not be listed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }
            }
        }
        catch {
            $gridUpdates.ItemsSource = @()
            $txtSummary.Text = 'Listing failed: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
        }
        finally { $dlg.Cursor = $null }
    }.GetNewClosure()

    $apply = {
        param([string]$Action)
        $selected = @($gridUpdates.SelectedItems)
        if ($selected.Count -eq 0) { $txtSummary.Text = 'Select one or more updates first.'; return }
        $group = ([string]$cboGroup.Text).Trim()
        if ($Action -eq 'Approve' -and -not $group) { $txtSummary.Text = 'Enter or choose a computer group to approve for.'; return }
        if ($Action -eq 'Expire') {
            $answer = Show-ThemedMessage -Owner $dlg -Title 'Expire Updates' -Buttons YesNo -Icon Warning `
                -Message ("Expire {0} update(s) on {1}?`n`nExpiry is permanent: clients stop seeing the update and it cannot be approved again. Declining is the reversible alternative." -f $selected.Count, [string]$Settings.ServerName)
            if ($answer -ne 'Yes') { return }
        }
        if ($Action -eq 'Remove') {
            $answer = Show-ThemedMessage -Owner $dlg -Title 'Remove Updates' -Buttons YesNo -Icon Warning `
                -Message ("Remove {0} update(s) from {1}?`n`nRemoval declines each update and deletes it from the WSUS database. It cannot be undone. Computers that installed an update keep the software.`n`nWith ConfigMgr, expire the update first and let the software update point synchronize, then remove it. WSUS refuses to remove an update that another update still references." -f $selected.Count, [string]$Settings.ServerName)
            if ($answer -ne 'Yes') { return }
        }
        $ids = [guid[]]@($selected | ForEach-Object { [guid][string]$_.PackageId })
        $outcomes = @()
        Set-DialogBusyStatus -Dialog $dlg -Control $txtSummary -Text ('{0}: working on {1} update(s)...' -f $Action, $selected.Count)
        try { $outcomes = @(Set-WsusPublishedUpdateState -Settings $Settings -PackageId $ids -Action $Action -GroupName $group) }
        catch { $outcomes = @([pscustomobject]@{ Title = ''; Ok = $false; Message = (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.') }) }
        finally { $dlg.Cursor = $null }
        & $load
        $failed = @($outcomes | Where-Object { -not $_.Ok } | ForEach-Object { if ($_.Title) { '{0}: {1}' -f $_.Title, $_.Message } else { [string]$_.Message } })
        $done = @($outcomes | Where-Object { $_.Ok }).Count
        $txtSummary.Text = ('{0}: {1} of {2} update(s) done.{3}' -f $Action, $done, $selected.Count, $(if ($failed.Count) { ' ' + ($failed -join ' | ') } else { '' }))
    }.GetNewClosure()

    $dlg.FindName('btnApprove').Add_Click({ & $apply 'Approve' }.GetNewClosure())
    $dlg.FindName('btnDecline').Add_Click({ & $apply 'Decline' }.GetNewClosure())
    $dlg.FindName('btnExpire').Add_Click({ & $apply 'Expire' }.GetNewClosure())
    $dlg.FindName('btnRemove').Add_Click({ & $apply 'Remove' }.GetNewClosure())
    $dlg.FindName('btnRefresh').Add_Click({ & $load }.GetNewClosure())
    $dlg.FindName('btnCatalog').Add_Click({ Show-WsusCatalogImportDialog -Owner $dlg -Settings $Settings }.GetNewClosure())
    $chkOther.Add_Click({ & $load }.GetNewClosure())
    $dlg.FindName('btnClose').Add_Click({ $dlg.Close() }.GetNewClosure())
    $dlg.Add_ContentRendered({ & $load }.GetNewClosure())
    [void]$dlg.ShowDialog()
}

function Show-WsusCatalogImportDialog {
    # Imports Microsoft Update Catalog updates into WSUS by update ID. The
    # server downloads the metadata itself; the import runs off the UI thread
    # because each ID waits on that download.
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][hashtable]$Settings)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Catalog Import" Width="780" Height="600" MinWidth="620" MinHeight="460"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False" GlowBrush="{DynamicResource MahApps.Brushes.Accent}" BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <!--SHARED_STYLES-->        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" TextWrapping="Wrap" FontSize="12" Margin="0,0,0,8"
                   Text="Paste update IDs or catalog links, one or more per line. The Copy button on a catalog update details page copies its update ID. The WSUS server downloads the update metadata itself, so it needs internet access; update files follow the server's Update files setting."/>
        <TextBox Grid.Row="1" x:Name="txtInput" AcceptsReturn="True" TextWrapping="NoWrap" FontFamily="Consolas" FontSize="12"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 Controls:TextBoxHelper.Watermark="12345678-90ab-cdef-1234-567890abcdef"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8">
            <Button x:Name="btnImport" Content="Import" Margin="0,0,8,0"
                    Style="{DynamicResource DialogAccentButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnLoadFile" Content="Load IDs from File..." Margin="0,0,8,0"
                    Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnOpenCatalog" Content="Open the Catalog"
                    Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
        <TextBox Grid.Row="3" x:Name="txtResults" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="12"
                 VerticalScrollBarVisibility="Auto"/>
        <DockPanel Grid.Row="4" Margin="0,10,0,0">
            <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Close" IsCancel="True"
                    Style="{DynamicResource DialogButton}" Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <TextBlock x:Name="txtStatus" FontSize="12" VerticalAlignment="Center" TextWrapping="Wrap"/>
        </DockPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$xml = Expand-SharedStyle $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogChromeFromOwner -Dialog $dlg -Owner $Owner

    $txtInput   = $dlg.FindName('txtInput')
    $txtResults = $dlg.FindName('txtResults')
    $txtStatus  = $dlg.FindName('txtStatus')
    $btnImport  = $dlg.FindName('btnImport')
    $modulePath = Join-Path $PSScriptRoot 'Packagers\AppPackagerWsus.psd1'
    $work = @{ Runspace = $null; PowerShell = $null; Timer = $null; Graveyard = @() }
    $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    $dlg.FindName('btnOpenCatalog').Add_Click({ Start-Process 'https://catalog.update.microsoft.com' }.GetNewClosure())
    $dlg.FindName('btnLoadFile').Add_Click({
        $open = New-Object Microsoft.Win32.OpenFileDialog
        $open.Title = 'Select a Text File of Update IDs'
        $open.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
        if ($open.ShowDialog()) { $txtInput.Text = [System.IO.File]::ReadAllText($open.FileName) }
    }.GetNewClosure())

    $btnImport.Add_Click({
        $ids = @(ConvertFrom-WsusCatalogInput -Text ([string]$txtInput.Text))
        if ($ids.Count -eq 0) { $txtStatus.Text = 'No update IDs found in the text.'; return }
        $btnImport.IsEnabled = $false
        $txtStatus.Text = ('Importing {0} update(s) into {1}...' -f $ids.Count, [string]$Settings.ServerName)
        try {
            if (-not $work.Runspace) { $work.Runspace = New-SuiteBgRunspace -ModulePath $modulePath }
            $ps = [powershell]::Create()
            $ps.Runspace = $work.Runspace
            [void]$ps.AddScript({
                param($Settings, $Ids, $Queue)
                foreach ($id in $Ids) {
                    $outcome = Import-WsusCatalogUpdate -Settings $Settings -UpdateId $id
                    $Queue.Enqueue(('{0}  {1}  {2}' -f $id, $(if ($outcome.Ok) { 'imported' } else { 'FAILED' }), [string]$outcome.Message))
                }
                $Queue.Enqueue('__done__')
            }).AddArgument($Settings).AddArgument([string[]]($ids | ForEach-Object { [string]$_ })).AddArgument($queue)
            $work.PowerShell = $ps
            [void]$ps.BeginInvoke()
        }
        catch {
            $txtStatus.Text = 'Import could not start: ' + (Get-UserErrorText -ErrorRecord $_ -Fallback 'The action failed. See the log file.')
            $btnImport.IsEnabled = $true
        }
    }.GetNewClosure())

    $work.Timer = New-Object System.Windows.Threading.DispatcherTimer
    $work.Timer.Interval = [TimeSpan]::FromMilliseconds(200)
    $work.Timer.Add_Tick({
        $line = $null
        while ($queue.TryDequeue([ref]$line)) {
            if ($line -eq '__done__') {
                $txtStatus.Text = 'Import finished. Server-side details: %ProgramFiles%\Update Services\LogFiles\SoftwareDistribution.log on the WSUS server.'
                $btnImport.IsEnabled = $true
                continue
            }
            $txtResults.AppendText($line + [Environment]::NewLine)
            $txtResults.ScrollToEnd()
        }
    }.GetNewClosure())
    $work.Timer.Start()

    $dlg.FindName('btnClose').Add_Click({ $dlg.Close() }.GetNewClosure())
    $dlg.Add_Closed({
        $work.Graveyard = @(Stop-SuiteBgWork -PowerShell $work.PowerShell -Timer $work.Timer -Graveyard $work.Graveyard)
        Close-SuiteBgRunspace -Runspace $work.Runspace
    }.GetNewClosure())
    [void]$dlg.ShowDialog()
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
    try { Initialize-SuiteFooter -Window $window -StatusName 'txtStatus' -OnOpenOptions { if (-not $script:ActionButtonsEnabled) { return }; Show-OptionsDialog -Owner $window; Update-ApFooter } | Out-Null }
    catch { Add-LogLine -Message 'The footer was not set up.' }
    Update-ApFooter
    Add-LogLine -Message ("Loading packagers from: {0}" -f $PackagersRoot)
    Invoke-RefreshGrid
    Add-LogLine -Message ("{0} packager(s) loaded. Ready." -f $script:PackagerData.Count)
    Update-SidebarForSystems
    Start-UpdateCheck
})

# Loaded runs before the first render, so a modal dialog opened there keeps
# the main window from painting until the dialog closes. The launch prompts
# wait for ContentRendered, which fires once, after the first render.
$window.Add_ContentRendered({
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
        Update-ApFooter
    }
})

$window.Add_Closing({
    param($closingSender, $closingArgs)
    # The conversion child would go on without a window, so closing waits for it.
    if ($script:CmImportRunning) {
        $closingArgs.Cancel = $true
        [void](Show-ThemedMessage -Owner $window -Title 'Conversion Running' `
            -Message 'A ConfigMgr conversion is running. Wait for it to finish, then close the window.' -Buttons OK -Icon Info)
        return
    }
    if ($script:UpdateCheckTimer) { $script:UpdateCheckTimer.Stop(); $script:UpdateCheckTimer = $null }
    if ($script:DiscoveryTimer) { $script:DiscoveryTimer.Stop(); $script:DiscoveryTimer = $null }
    Stop-DiscoveryRunspace -Job $script:DiscoveryJob
    $script:DiscoveryJob = $null

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
