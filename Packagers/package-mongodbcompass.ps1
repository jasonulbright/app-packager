<#
Vendor: MongoDB
App: MongoDB Compass
CMName: MongoDB Compass
VendorUrl: https://www.mongodb.com/products/tools/compass
CPE: cpe:2.3:a:mongodb:compass:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/mongodb-js/compass/releases
DownloadPageUrl: https://www.mongodb.com/try/download/compass
UpdateCadenceDays: 30

.SYNOPSIS
    Packages MongoDB Compass (x64) MSI for MECM.

.DESCRIPTION
    Resolves the current stable Compass version from the vendor's "latest"
    download endpoint, downloads the matching x64 MSI, stages content to a
    versioned local folder with ARP detection metadata, and creates an MECM
    Application with registry-based detection.

    The latest endpoint redirects to the EXE build; the version is taken from
    that redirect and the MSI of the same version is downloaded instead,
    because only the MSI installs silently per-machine.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\MongoDB\MongoDB Compass\<Version>

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
    Runs only the Stage phase: download MSI, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available MongoDB Compass version string and exits.

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
$LatestFeedUrl  = "https://compass.mongodb.com/api/v2/download/latest/compass/stable/windows"
$MsiUrlTemplate = "https://downloads.mongodb.com/compass/mongodb-compass-{0}-win32-x64.msi"

$VendorFolder = "MongoDB"
$AppFolder    = "MongoDB Compass"

$BaseDownloadRoot = Join-Path $DownloadRoot "MongoDBCompass"

# --- Functions ---


function Assert-PayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file
        signature that every Windows Installer package carries.
    .DESCRIPTION
        The download host answers 200 with an HTML body for a version that was
        never published as an MSI, which would otherwise stage as a
        valid-looking MSI and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestCompassVersion {
    <#
    .SYNOPSIS
        Reads the Compass version from the latest-download redirect target.
    .DESCRIPTION
        The endpoint answers 302 with a Location naming the versioned build;
        following the redirect would pull the full installer, so only the
        headers are fetched.
    #>
    param([switch]$Quiet)

    Write-Log "Compass latest feed URL      : $LatestFeedUrl" -Quiet:$Quiet

    try {
        $headers = (curl.exe --fail --silent --show-error --head $LatestFeedUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the Compass latest endpoint: $LatestFeedUrl" }

        $m = [regex]::Match($headers, 'mongodb-compass-(?<ver>\d+(?:\.\d+)+)-win32-x64\.')
        if (-not $m.Success) {
            throw "Could not read a Compass version from the redirect target."
        }

        $version = $m.Groups['ver'].Value
        Write-Log "Latest Compass version       : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get MongoDB Compass version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCompass {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MongoDB Compass (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestCompassVersion
    if (-not $version) { throw "Could not resolve MongoDB Compass version." }

    $msiFileName = "mongodb-compass-$version-win32-x64.msi"
    $downloadUrl = $MsiUrlTemplate -f $version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-PayloadIsMsi -Path $localMsi

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

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    # For a standard MSI install the ARP uninstall key name is the ProductCode
    # GUID, so detection needs no temp install cycle.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "MongoDB, Inc." }

    $appName = "MongoDB Compass $version"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName         = $appName
        DisplayName     = "MongoDB Compass"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("MongoDBCompass")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
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

function Invoke-PackageCompass {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MongoDB Compass (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

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
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
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
        $v = Get-LatestCompassVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("MongoDB Compass GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MongoDB Compass (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LatestFeedUrl                : $LatestFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCompass
    }
    elseif ($PackageOnly) {
        Invoke-PackageCompass
    }
    else {
        Invoke-StageCompass
        Invoke-PackageCompass
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-mongodbcompass'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
