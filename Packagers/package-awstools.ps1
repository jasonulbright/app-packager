<#
Vendor: Amazon
App: Amazon AWS Tools for Windows
CMName: AWS Tools for Windows
VendorUrl: https://aws.amazon.com/powershell/
CPE: cpe:2.3:a:amazon:aws_tools_for_windows:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/aws/aws-tools-for-powershell/blob/main/changelogs/CHANGELOG.ALL.md
DownloadPageUrl: https://aws.amazon.com/powershell/
UpdateCadenceDays: 180

.SYNOPSIS
    Packages AWS Tools for Windows (x64) MSI for MECM.

.DESCRIPTION
    Downloads the AWS Tools for Windows MSI from the vendor's static
    distribution URL, stages content to a versioned local folder with ARP
    detection metadata, and creates an MECM Application with registry-based
    detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The distribution URL always serves the current bundle and publishes no
    version metadata, so the version is read from the MSI's ProductVersion
    property. GetLatestVersionOnly reuses an already-downloaded copy when one
    is present rather than re-transferring the ~200 MB payload.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Amazon\AWS Tools for Windows\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AWS Tools for Windows).
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
    Outputs only the latest available AWS Tools for Windows version string and exits.

.PARAMETER ForceRefresh
    Ignores the cached installer and re-downloads it before reading the version.

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
    [switch]$ForceRefresh,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$MsiDownloadUrl = "https://sdk-for-net.amazonwebservices.com/latest/AWSToolsAndSDKForNet.msi"
$MsiFileName    = "AWSToolsAndSDKForNet.msi"

$VendorFolder = "Amazon"
$AppFolder    = "AWS Tools for Windows"

$BaseDownloadRoot = Join-Path $DownloadRoot "AWS Tools for Windows"

# --- Functions ---


function Assert-PayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file
        signature that every Windows Installer package carries.
    .DESCRIPTION
        A CDN edge can answer 200 with an error document; that body would
        otherwise stage as a valid-looking MSI.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $expected = @(0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $buffer = New-Object byte[] 8
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 8) } finally { $stream.Dispose() }

    if ($read -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($buffer[$i] -ne $expected[$i]) {
            throw "Downloaded payload is not a Windows Installer package (no OLE compound-file header): $Path"
        }
    }
}


function Get-AwsToolsInstaller {
    <#
    .SYNOPSIS
        Returns the path to a verified local copy of the installer, downloading
        it when the cache is empty or a refresh was requested.
    #>
    param(
        [switch]$Force,
        [switch]$Quiet
    )

    Initialize-Folder -Path $BaseDownloadRoot
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName

    if ($Force -and (Test-Path -LiteralPath $localMsi)) {
        Remove-Item -LiteralPath $localMsi -Force -ErrorAction Stop
    }

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $MsiDownloadUrl" -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $MsiDownloadUrl -OutFile $localMsi -Quiet:$Quiet
    }
    else {
        Write-Log "Local MSI exists. Skipping download." -Quiet:$Quiet
    }

    Assert-PayloadIsMsi -Path $localMsi
    return $localMsi
}


function Get-LatestAwsToolsVersion {
    param(
        [switch]$Force,
        [switch]$Quiet
    )

    try {
        $localMsi = Get-AwsToolsInstaller -Force:$Force -Quiet:$Quiet
        $props = Get-MsiPropertyMap -MsiPath $localMsi
        $version = $props["ProductVersion"]
        if ([string]::IsNullOrWhiteSpace($version)) { throw "MSI ProductVersion missing." }

        Write-Log "Latest AWS Tools version     : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get AWS Tools version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAwsTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Tools for Windows (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    # Staging always refreshes: the rolling URL is the only version signal, so
    # a cached copy could otherwise stage a superseded bundle.
    $localMsi = Get-AwsToolsInstaller -Force
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName    = $props["ProductName"]
    $productVersion = $props["ProductVersion"]
    $manufacturer   = $props["Manufacturer"]
    $productCode    = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))    { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersion"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
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
    # For a standard MSI install the ARP uninstall key name is the ProductCode
    # GUID, so detection needs no temp install/uninstall cycle.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersion"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrappers = New-MsiWrapperContent -MsiFileName $MsiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Amazon Web Services" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "AWS Tools for Windows"
        Publisher       = $publisher
        SoftwareVersion = $productVersion
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersion
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

function Invoke-PackageAwsTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Tools for Windows (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    if (-not (Test-Path -LiteralPath $localMsi)) {
        throw "Local MSI not found - run Stage phase first: $localMsi"
    }

    $props = Get-MsiPropertyMap -MsiPath $localMsi
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $localContentPath = Join-Path $BaseDownloadRoot $props["ProductVersion"]
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

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
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
        $v = Get-LatestAwsToolsVersion -Force:$ForceRefresh -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("AWS Tools GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Tools for Windows (x64) Auto-Packager starting"
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
        Invoke-StageAwsTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageAwsTools
    }
    else {
        Invoke-StageAwsTools
        Invoke-PackageAwsTools
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-awstools'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
