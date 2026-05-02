<#
Vendor: The Document Foundation
App: LibreOffice
CMName: LibreOffice
VendorUrl: https://www.libreoffice.org/
CPE: cpe:2.3:a:libreoffice:libreoffice:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://wiki.documentfoundation.org/Releases
DownloadPageUrl: https://www.libreoffice.org/download/download-libreoffice/

.SYNOPSIS
    Packages LibreOffice (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest LibreOffice x64 MSI from the Document Foundation CDN,
    stages content to a versioned local folder with ARP detection metadata,
    and creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest version is resolved from the endoflife.date API. The MSI is
    approximately 350MB so staging takes longer than most packagers.

    The MSI install passes ADDLOCAL=ALL and enterprise-appropriate defaults
    for auto-update and desktop shortcuts.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\The Document Foundation\LibreOffice (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\LibreOffice).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Queries the endoflife.date API for the latest LibreOffice version, outputs
    the version string, and exits. No download or MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 60,
    [string]$LogPath,
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
$CdnStableUrl = "https://download.documentfoundation.org/libreoffice/stable/"

$VendorFolder = "The Document Foundation"
$AppFolder    = "LibreOffice (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "LibreOffice"

# --- Functions ---


function Get-LatestLibreOfficeVersion {
    param([switch]$Quiet)

    # Scrape the CDN stable directory listing for the actual latest downloadable
    # version. The endoflife.date API returns 4-part versions (e.g. 26.2.1.1)
    # that do not match the CDN folder names (e.g. 26.2.0).
    Write-Log "LibreOffice CDN stable       : $CdnStableUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $CdnStableUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query LibreOffice CDN." }

        # Extract version folder names (e.g., "26.2.0/") from href attributes
        $versions = [regex]::Matches($html, 'href="(\d+\.\d+\.\d+)/"') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object { [version]$_ } -Descending
        if (-not $versions -or $versions.Count -eq 0) {
            throw "No version folders found on LibreOffice CDN."
        }

        $latest = $versions[0]
        Write-Log "Latest LibreOffice version   : $latest" -Quiet:$Quiet
        return $latest
    }
    catch {
        Write-Log "Failed to get LibreOffice version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageLibreOffice {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "LibreOffice (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestLibreOfficeVersion
    if (-not $version) { throw "Could not resolve LibreOffice version." }

    # Build download URL - version parts needed for CDN path
    # Format: LibreOffice_{version}_Win_x86-64.msi
    # CDN: https://download.documentfoundation.org/libreoffice/stable/{version}/win/x86_64/
    $MsiFileName = "LibreOffice_${version}_Win_x86-64.msi"
    $downloadUrl = "https://download.documentfoundation.org/libreoffice/stable/${version}/win/x86_64/${MsiFileName}"

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading LibreOffice MSI (~350MB)..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

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
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
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
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers (custom install args) ---
    $installPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''/quiet'', ''/norestart'', ''ADDLOCAL=ALL'', ''CREATEDESKTOPLINK=0'', ''ISCHECKFORPRODUCTUPDATES=0'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $publisher = "The Document Foundation"
    $appName = "LibreOffice $productVersionRaw"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/quiet /norestart ADDLOCAL=ALL CREATEDESKTOPLINK=0 ISCHECKFORPRODUCTUPDATES=0"
        UninstallArgs   = "/quiet /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("soffice")
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

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageLibreOffice {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "LibreOffice (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $msiFiles = Get-ChildItem -Path $BaseDownloadRoot -Filter "LibreOffice_*_Win_x86-64.msi" -File
    if (-not $msiFiles -or $msiFiles.Count -eq 0) {
        throw "No staged LibreOffice MSI found - run Stage phase first."
    }
    $localMsi = $msiFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    $props = Get-MsiPropertyMap -MsiPath $localMsi.FullName
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $productVersion   = $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
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

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    $localFiles = Get-ChildItem -Path $localContentPath -File -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $f.Name
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $($f.Name)"
        }
        else {
            Write-Log "Already on network           : $($f.Name)"
        }
    }

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
        $v = Get-LatestLibreOfficeVersion -Quiet
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
    Write-Log "LibreOffice (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CdnStableUrl                 : $CdnStableUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageLibreOffice
    }
    elseif ($PackageOnly) {
        Invoke-PackageLibreOffice
    }
    else {
        Invoke-StageLibreOffice
        Invoke-PackageLibreOffice
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
