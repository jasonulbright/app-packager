<#
Vendor: Zoom Video Communications
App: Zoom Workplace
CMName: Zoom Workplace
VendorUrl: https://zoom.us/download
CPE: cpe:2.3:a:zoom:zoom:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0061222
DownloadPageUrl: https://zoom.us/download
IconSource: Installer
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Zoom Workplace (x64) MSI for MECM.

.DESCRIPTION
    Downloads the full Zoom Workplace x64 MSI from the vendor's "latest" URL,
    stages content to a versioned local folder with ARP detection metadata
    derived from MSI properties, and creates an MECM Application with
    registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: Zoom's "latest" MSI URL always serves the current general release. The
    version is read from MSI properties after download.

    GetLatestVersionOnly follows the ZoomInstaller.exe redirect, which resolves
    to a CDN path carrying the current build number, and exits without
    downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Zoom\Zoom Workplace\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Zoom).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Zoom version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [ValidateSet('Nested','Flat')]
    [string]$ContentLayout = "Nested",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force -ErrorAction Stop
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$VersionProbeUrl = "https://zoom.us/client/latest/ZoomInstaller.exe"
$MsiDownloadUrl  = "https://zoom.us/client/latest/ZoomInstallerFull.msi"
$MsiFileName     = "ZoomInstallerFull.msi"

$VendorFolder = "Zoom"
$AppFolder    = "Zoom Workplace"

$BaseDownloadRoot = Join-Path $DownloadRoot "Zoom"
$ExtraInstallArg  = "ZoomAutoUpdate=Disabled"

# --- Functions ---


function Get-LatestZoomVersion {
    param([switch]$Quiet)

    Write-Log "Zoom version probe URL       : $VersionProbeUrl" -Quiet:$Quiet

    try {
        $effective = (curl.exe --silent --show-error --head --location `
            --write-out "%{url_effective}" --output NUL $VersionProbeUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to probe Zoom download redirect: $VersionProbeUrl" }

        if ($effective -notmatch '/prod/(?<ver>\d+\.\d+\.\d+\.\d+)/') {
            throw "Redirect target did not contain a build number: $effective"
        }
        $version = $Matches['ver']

        Write-Log "Latest Zoom version          : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Zoom version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageZoom {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zoom Workplace (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    # The MSI ProductVersion drops the vendor's third component
    # (6.6.11.23272 ships as ProductVersion 6.6.23272), so the redirect probe
    # supplies the release identity and the MSI property supplies the value
    # the client writes to the registry.
    $version = Get-LatestZoomVersion
    if (-not $version) { throw "Could not resolve Zoom version." }

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""
    Write-Log "Downloading Zoom Workplace MSI..."
    Invoke-DownloadWithRetry -Url $MsiDownloadUrl -OutFile $localMsi

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    # For standard MSI installs the ARP uninstall key name is the ProductCode
    # GUID, so detection needs no temp install cycle.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    # ZoomAutoUpdate=Disabled keeps the client on the packaged build so the
    # deployed version and the detection value stay in agreement.
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName -ExtraInstallArgs @($ExtraInstallArg)
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Zoom Video Communications" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName         = "Zoom Workplace"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart $ExtraInstallArg"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("Zoom")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageZoom {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zoom Workplace (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

    # --- MECM application ---
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
        $v = Get-LatestZoomVersion -Quiet
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
    Write-Log "Zoom Workplace (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "MsiDownloadUrl               : $MsiDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageZoom
    }
    elseif ($PackageOnly) {
        Invoke-PackageZoom
    }
    else {
        Invoke-StageZoom
        Invoke-PackageZoom
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-zoom'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
