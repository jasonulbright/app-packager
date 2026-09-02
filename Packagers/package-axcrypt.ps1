<#
Vendor: AxCrypt
App: AxCrypt
CMName: AxCrypt
VendorUrl: https://www.axcrypt.net/
CPE: cpe:2.3:a:axcrypt:axcrypt:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.axcrypt.net/download/
DownloadPageUrl: https://www.axcrypt.net/download/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages AxCrypt (x64) MSI for MECM.

.DESCRIPTION
    Downloads the current AxCrypt x64 MSI from the vendor distribution host,
    stages content to a versioned local folder with ARP detection metadata, and
    creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: The vendor MSI URL is unversioned and always serves the current
    release. The authoritative version is read from MSI properties after
    download; the download page is only used for the no-download version query.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\AxCrypt\AxCrypt\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AxCrypt).
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
    Outputs only the latest available AxCrypt version string and exits.

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
$DownloadPageUrl = "https://www.axcrypt.net/download/"
$MsiDownloadUrl  = "https://downloader.axcrypt.net/windows/axcrypt-3-appsetup-win_x64.msi"

$VendorFolder = "AxCrypt"
$AppFolder    = "AxCrypt"

$BaseDownloadRoot = Join-Path $DownloadRoot "AxCrypt"
$MsiFileName      = "axcrypt-3-appsetup-win_x64.msi"

# --- Functions ---


function Assert-AxCryptPayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-document
        signature that every MSI carries.
    .DESCRIPTION
        The distribution host answers 200 with an HTML page when a download key
        no longer resolves, which would otherwise stage as a valid-looking MSI.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $expected = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    $buffer = New-Object byte[] $expected.Length

    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, $buffer.Length) }
    finally { $stream.Dispose() }

    if ($read -lt $expected.Length) {
        throw "Downloaded payload is too small to be an MSI: $Path"
    }
    for ($i = 0; $i -lt $expected.Length; $i++) {
        if ($buffer[$i] -ne $expected[$i]) {
            throw "Downloaded payload is not an MSI (no OLE compound-document header): $Path"
        }
    }
}


function Get-AxCryptPublishedVersion {
    <#
    .SYNOPSIS
        Returns the Windows version shown on the vendor download page.
    .DESCRIPTION
        The page lists one version per platform; the match is anchored on the
        Windows card so a macOS or mobile version is never picked up.
        Returns $null when the page layout no longer yields a match.
    #>
    param([switch]$Quiet)

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch AxCrypt download page: $DownloadPageUrl" }

        $m = [regex]::Match($html, '(?s)platform-icon-windows.*?Version\s+(?<ver>\d+(?:\.\d+){1,3})')
        if (-not $m.Success) { return $null }

        Write-Log "Published Windows version    : $($m.Groups['ver'].Value)" -Quiet:$Quiet
        return $m.Groups['ver'].Value
    }
    catch {
        Write-Log "Failed to read AxCrypt download page: $($_.Exception.Message)" -Level WARN -Quiet:$Quiet
        return $null
    }
}


function Get-AxCryptMsiVersion {
    <#
    .SYNOPSIS
        Downloads the current MSI and returns its ProductVersion.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName

    Invoke-DownloadWithRetry -Url $MsiDownloadUrl -OutFile $localMsi -Quiet:$Quiet
    Assert-AxCryptPayloadIsMsi -Path $localMsi

    $props = Get-MsiPropertyMap -MsiPath $localMsi
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from AxCrypt MSI."
    }
    return $props["ProductVersion"]
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAxCrypt {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AxCrypt (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    # The MSI URL is unversioned, so a cached file from an earlier release
    # would shadow the current one.
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log "Download URL                 : $MsiDownloadUrl"
    Write-Log ""
    Write-Log "Downloading AxCrypt MSI..."
    Invoke-DownloadWithRetry -Url $MsiDownloadUrl -OutFile $localMsi

    Assert-AxCryptPayloadIsMsi -Path $localMsi

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
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "AxCrypt AB" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "AxCrypt $productVersionRaw"
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("AxCrypt")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
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

function Invoke-PackageAxCrypt {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AxCrypt (x64) - PACKAGE phase"
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

        $v = Get-AxCryptPublishedVersion -Quiet
        if (-not $v) { $v = Get-AxCryptMsiVersion -Quiet }
        if (-not $v) { exit 1 }

        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("AxCrypt GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AxCrypt (x64) Auto-Packager starting"
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
        Invoke-StageAxCrypt
    }
    elseif ($PackageOnly) {
        Invoke-PackageAxCrypt
    }
    else {
        Invoke-StageAxCrypt
        Invoke-PackageAxCrypt
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-axcrypt'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
