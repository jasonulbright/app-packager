<#
Vendor: Devolutions
App: Remote Desktop Manager
CMName: Remote Desktop Manager
VendorUrl: https://devolutions.net/
CPE: cpe:2.3:a:devolutions:remote_desktop_manager:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://devolutions.net/remote-desktop-manager/home/whatsnew/
DownloadPageUrl: https://devolutions.net/remote-desktop-manager/home/download/
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Remote Desktop Manager (x64 MSI, Free edition entitlement) for MECM.

.DESCRIPTION
    Resolves the current release from the vendor version feed at
    devolutions.net/productinfo.htm, downloads the x64 MSI, stages content to a
    versioned local folder with ARP detection metadata, and creates an MECM
    Application with registry-based detection.

    Free and paid editions ship the same installer; the edition follows from the
    license applied after install, so the feed's RDMFreemsi and RDMmsi entries
    point at the same file. The x64-specific key is used so the payload matches
    the architecture this packager targets.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Remote Desktop Manager version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$ProductInfoUrl = "https://devolutions.net/productinfo.htm"
$FeedKey        = "RDMmsiX64"

$VendorFolder = "Devolutions"
$AppFolder    = "Remote Desktop Manager"

$BaseDownloadRoot = Join-Path $DownloadRoot "Remote Desktop Manager"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        The vendor CDN answers 200 with an HTML body for withdrawn builds, which
        would otherwise stage as a valid-looking MSI and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestRdmRelease {
    <#
    .SYNOPSIS
        Returns the current stable release version, MSI URL and published hash.
    .DESCRIPTION
        The feed is a flat key=value list that carries beta and LTS builds under
        their own key prefixes, so the key is matched whole: a substring match on
        RDMmsiX64 would also accept RDMmsiX64Beta.
    #>
    param([switch]$Quiet)

    Write-Log "Version feed                 : $ProductInfoUrl" -Quiet:$Quiet

    try {
        $lines = curl.exe -L --fail --silent --show-error $ProductInfoUrl
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Devolutions product info feed: $ProductInfoUrl" }

        $map = @{}
        foreach ($line in $lines) {
            if ($line -match '^\s*(?<key>[A-Za-z0-9_.]+)\s*=\s*(?<value>.+?)\s*$') {
                $map[$matches['key']] = $matches['value']
            }
        }

        $version = $map["$FeedKey.Version"]
        $url     = $map["$FeedKey.Url"]
        $hash    = $map["$FeedKey.hash"]

        if ([string]::IsNullOrWhiteSpace($version)) { throw "Feed key $FeedKey.Version not found." }
        if ([string]::IsNullOrWhiteSpace($url))     { throw "Feed key $FeedKey.Url not found." }

        Write-Log "Latest RDM version           : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = Split-Path -Leaf ([uri]$url).AbsolutePath
            DownloadUrl = $url
            Sha256      = $hash
        }
    }
    catch {
        Write-Log "Failed to get Remote Desktop Manager version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageRdm {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Remote Desktop Manager (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestRdmRelease
    if (-not $releaseInfo) { throw "Could not resolve Remote Desktop Manager version." }

    $version     = $releaseInfo.Version
    $msiFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-MsiPayload -Path $localMsi

    # The feed publishes a SHA256 per build; a mismatch means the CDN served a
    # different payload than the feed describes.
    if (-not [string]::IsNullOrWhiteSpace($releaseInfo.Sha256)) {
        $actual = (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $releaseInfo.Sha256.ToUpperInvariant()) {
            throw "Downloaded MSI hash does not match the vendor feed value for $version."
        }
        Write-Log "Vendor SHA256 verified       : $actual"
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
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
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
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Devolutions" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Remote Desktop Manager"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("RemoteDesktopManager")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageRdm {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Remote Desktop Manager (x64) - PACKAGE phase"
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

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

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
        $info = Get-LatestRdmRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Remote Desktop Manager GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Remote Desktop Manager (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ProductInfoUrl               : $ProductInfoUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageRdm
    }
    elseif ($PackageOnly) {
        Invoke-PackageRdm
    }
    else {
        Invoke-StageRdm
        Invoke-PackageRdm
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-remotedesktopmanager'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
