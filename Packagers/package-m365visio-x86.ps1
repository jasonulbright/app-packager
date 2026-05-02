<#
Vendor: Microsoft
App: M365 Visio (x86)
CMName: M365 Visio
VendorUrl: https://www.microsoft.com/microsoft-365
CPE: cpe:2.3:a:microsoft:365_apps:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date
DownloadPageUrl: https://www.microsoft.com/en-us/microsoft-365

.SYNOPSIS
    Packages M365 Visio (x86) for MECM using the Office Deployment Tool.

.DESCRIPTION
    Downloads the latest Office Deployment Tool, uses it to download the offline
    source files for M365 Visio (x86) from the configured update channel
    (Monthly Enterprise or Current), stages content to a versioned local folder
    with file-based detection metadata, and creates an MECM Application.

    Detection uses VISIO.EXE file version >= packaged version in the
    Program Files (x86) install path.

    Supports two-phase operation:
      -StageOnly    Download ODT + Office source, generate wrappers + manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Local administrator
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [ValidateSet('MonthlyEnterprise','Current')]
    [string]$M365Channel = "MonthlyEnterprise",
    [ValidateSet('Managed','Online')]
    [string]$M365DeployMode = "Managed",
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
# Channel resolution from preference
$ChannelMap = @{
    'MonthlyEnterprise' = @{ Guid = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'; Name = 'MonthlyEnterprise'; Display = 'Monthly Enterprise Channel'; Tag = 'MEC' }
    'Current'           = @{ Guid = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'; Name = 'Current';           Display = 'Current Channel';              Tag = 'CC' }
}
$ch = $ChannelMap[$M365Channel]
if (-not $ch) { Write-Log "Invalid M365Channel: $M365Channel" -Level ERROR; exit 1 }
$ChannelGuid     = $ch.Guid
$ChannelName     = $ch.Name
$ProductId       = "VisioProRetail"
$Architecture    = "32"
$CdnBaseUrl      = "https://officecdn.microsoft.com/pr/$ChannelGuid/Office/Data"
$CdnCabUrl       = "$CdnBaseUrl/v32.cab"
$OdtSetupUrl     = "https://officecdn.microsoft.com/pr/wsus/setup.exe"

$VendorFolder    = "Microsoft"
$AppFolder       = "M365 Visio (x86) [$($ch.Tag)]"
$BaseDownloadRoot = Join-Path $DownloadRoot "M365Visio-x86-$($ch.Tag)"

$DetectionExe    = "VISIO.EXE"
$DetectionPath   = "{0}\Microsoft Office\root\Office16" -f ${env:ProgramFiles(x86)}
$ProductIds      = @('O365ProPlusRetail', $ProductId)

# --- Functions ---


function Get-M365VersionFromCdn {
    param([switch]$Quiet)

    Write-Log "CDN cab URL                  : $CdnCabUrl" -Quiet:$Quiet

    try {
        $tempDir = Join-Path $BaseDownloadRoot "_cdntemp"
        Initialize-Folder -Path $tempDir

        $cabPath = Join-Path $tempDir "v$Architecture.cab"
        Invoke-DownloadWithRetry -Url $CdnCabUrl -OutFile $cabPath -Quiet:$Quiet

        $expandOut = & "$env:SystemRoot\System32\expand.exe" $cabPath -F:VersionDescriptor.xml $tempDir 2>&1
        $vdPath = Join-Path $tempDir "VersionDescriptor.xml"
        if (-not (Test-Path -LiteralPath $vdPath)) {
            throw "VersionDescriptor.xml not found in cab file. expand.exe output: $expandOut"
        }

        [xml]$vdXml = Get-Content -LiteralPath $vdPath -Raw -ErrorAction Stop
        $version = $vdXml.Version.Available.I320Version
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Build version not found in VersionDescriptor.xml."
        }

        Write-Log "CDN version                  : $version" -Quiet:$Quiet
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $version
    }
    catch {
        Write-Log "Failed to get M365 version from CDN: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-OdtSetupExe {
    $setupExe = Join-Path $BaseDownloadRoot "setup.exe"

    Write-Log "ODT setup.exe URL            : $OdtSetupUrl"
    Write-Log "Downloading ODT setup.exe..."
    Invoke-DownloadWithRetry -Url $OdtSetupUrl -OutFile $setupExe

    if (-not (Test-Path -LiteralPath $setupExe)) {
        throw "ODT setup.exe download failed: $setupExe"
    }

    Write-Log "ODT setup.exe downloaded     : $setupExe"
    return $setupExe
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageM365Visio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "M365 Visio (x86) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $isOnline = ($M365DeployMode -eq 'Online')
    Write-Log "Deploy mode                  : $M365DeployMode"

    # --- Load preferences ---
    $prefs = Get-PackagerPreferences
    $companyName = if ($prefs -and $prefs.CompanyName) { $prefs.CompanyName } else { $null }
    $excludeApps = if ($prefs -and $null -ne $prefs.M365ExcludeApps) { @($prefs.M365ExcludeApps) } else { @('Groove','Lync','OneDrive','Teams','Bing') }

    if ($isOnline) {
        # -- Online mode: CDN-direct, no source download, no version pin --
        $version = "Latest"
        Write-Log "Version                      : Latest (CDN-direct, no precache)"
        Write-Log ""

        $localContentPath = Join-Path $BaseDownloadRoot $version  # CDN mode: version is 'Latest'
        Initialize-Folder -Path $localContentPath

        # Download ODT setup.exe only
        $setupExe = Get-OdtSetupExe
        $contentSetupExe = Join-Path $localContentPath "setup.exe"
        Copy-Item -LiteralPath $setupExe -Destination $contentSetupExe -Force -ErrorAction Stop
        Write-Log "Copied setup.exe to content  : $contentSetupExe"

        # Install XML: no SourcePath, no Version ï¿½ pulls latest from CDN
        $installXmlPath = Join-Path $localContentPath "install.xml"
        $installXml = New-OdtConfigXml -OfficeClientEdition $Architecture -ProductIds $ProductIds -Channel $ChannelName -CompanyName $companyName -ExcludeApps $excludeApps
        Set-Content -LiteralPath $installXmlPath -Value $installXml -Encoding ASCII -ErrorAction Stop
        Write-Log "Written install.xml          : $installXmlPath (online, no version pin)"

        # Detection: existence only ï¿½ WSUS/CDN handles updates
        $detection = @{
            Type         = "File"
            FilePath     = $DetectionPath
            FileName     = $DetectionExe
            PropertyType = "Existence"
            Is64Bit      = $false
        }

        $appName = "M365 Visio (x86) ($($ch.Tag))"
    }
    else {
        # -- Managed mode: pin version, precache source --
        $version = Get-M365VersionFromCdn
        if (-not $version) { throw "Could not determine M365 version from CDN." }

        Write-Log "Version                      : $version"
        Write-Log ""

        $localContentPath = Join-Path $BaseDownloadRoot $version
        Initialize-Folder -Path $localContentPath

        # Download ODT
        $setupExe = Get-OdtSetupExe

        # Download config ï¿½ includes SourcePath for offline caching
        $downloadXmlPath = Join-Path $BaseDownloadRoot "download.xml"
        $downloadXml = New-OdtConfigXml -OfficeClientEdition $Architecture -Version $version -ProductIds $ProductIds -SourcePath $localContentPath -Channel $ChannelName -CompanyName $companyName -ExcludeApps $excludeApps
        Set-Content -LiteralPath $downloadXmlPath -Value $downloadXml -Encoding ASCII -ErrorAction Stop
        Write-Log "Written download.xml         : $downloadXmlPath"

        # Download Office source files
        Write-Log ""
        Write-Log "Downloading Office source files (this may take several minutes)..."
        $dlProc = Start-Process -FilePath $setupExe -ArgumentList @('/download', $downloadXmlPath) -Wait -PassThru -NoNewWindow
        if ($dlProc.ExitCode -ne 0) {
            throw "ODT /download failed with exit code $($dlProc.ExitCode)."
        }
        Write-Log "Office source download complete."
        Write-Log ""

        # Copy setup.exe into content folder
        $contentSetupExe = Join-Path $localContentPath "setup.exe"
        if (-not (Test-Path -LiteralPath $contentSetupExe)) {
            Copy-Item -LiteralPath $setupExe -Destination $contentSetupExe -Force -ErrorAction Stop
            Write-Log "Copied setup.exe to content  : $contentSetupExe"
        }

        # Install XML: offline, version-pinned
        $installXmlPath = Join-Path $localContentPath "install.xml"
        $installXml = New-OdtConfigXml -OfficeClientEdition $Architecture -Version $version -ProductIds $ProductIds -Channel $ChannelName -CompanyName $companyName -ExcludeApps $excludeApps
        Set-Content -LiteralPath $installXmlPath -Value $installXml -Encoding ASCII -ErrorAction Stop
        Write-Log "Written install.xml          : $installXmlPath"

        # Detection: version-based
        $detection = @{
            Type          = "File"
            FilePath      = $DetectionPath
            FileName      = $DetectionExe
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
        }

        $appName = "M365 Visio - $version (x86) ($($ch.Tag))"
    }

    # --- Uninstall XML (same for both modes) ---
    $uninstallXmlPath = Join-Path $localContentPath "uninstall.xml"
    $uninstallXml = @(
        '<Configuration>',
        '  <Remove>',
        ('    <Product ID="{0}" />' -f $ProductId),
        '  </Remove>',
        '  <Display Level="None" AcceptEULA="TRUE" />',
        '</Configuration>'
    ) -join "`r`n"
    Set-Content -LiteralPath $uninstallXmlPath -Value $uninstallXml -Encoding ASCII -ErrorAction Stop
    Write-Log "Written uninstall.xml        : $uninstallXmlPath"

    # --- Generate content wrappers ---
    $installPs1 = @(
        '$setupPath = Join-Path $PSScriptRoot ''setup.exe''',
        '$configPath = Join-Path $PSScriptRoot ''install.xml''',
        '$proc = Start-Process -FilePath $setupPath -ArgumentList @(''/configure'', $configPath) -Wait -PassThru -NoNewWindow -WorkingDirectory $PSScriptRoot',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = @(
        '$setupPath = Join-Path $PSScriptRoot ''setup.exe''',
        '$configPath = Join-Path $PSScriptRoot ''uninstall.xml''',
        '$proc = Start-Process -FilePath $setupPath -ArgumentList @(''/configure'', $configPath) -Wait -PassThru -NoNewWindow -WorkingDirectory $PSScriptRoot',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $publisher = "Microsoft"

    Write-Log ""
    Write-Log "Detection path               : $DetectionPath"
    Write-Log "Detection file               : $DetectionExe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        DisplayName     = "M365 Visio (x86)"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = "setup.exe"
        InstallerType   = "ODT"
        InstallArgs     = "/configure install.xml"
        UninstallArgs   = "/configure uninstall.xml"
        RunningProcess  = @("VISIO")
        Detection       = $detection
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageM365Visio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "M365 Visio (x86) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    $items = Get-ChildItem -Path $localContentPath -ErrorAction Stop
    foreach ($item in $items) {
        if ($item.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $item.Name
        if ($item.PSIsContainer) {
            if (-not (Test-Path -LiteralPath $dest)) {
                Copy-Item -Path $item.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
                Write-Log "Copied directory to network  : $($item.Name)"
            }
            else {
                Write-Log "Directory exists on network  : $($item.Name)"
            }
        }
        else {
            if (-not (Test-Path -LiteralPath $dest)) {
                Copy-Item -LiteralPath $item.FullName -Destination $dest -Force -ErrorAction Stop
                Write-Log "Copied to network            : $($item.Name)"
            }
            else {
                Write-Log "Already on network           : $($item.Name)"
            }
        }
    }

    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
}


# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        Initialize-Folder -Path $BaseDownloadRoot
        $v = Get-M365VersionFromCdn -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "M365 Visio (x86) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "Channel                      : $($ch.Display)"
    Write-Log "Deploy mode                  : $M365DeployMode"
    Write-Log "Product                      : $ProductId"
    Write-Log "Architecture                 : x$Architecture"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageM365Visio
    }
    elseif ($PackageOnly) {
        Invoke-PackageM365Visio
    }
    else {
        Invoke-StageM365Visio
        Invoke-PackageM365Visio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
