<#
Vendor: UrBackup
App: UrBackup Client
CMName: UrBackup Client
VendorUrl: https://www.urbackup.org/
CPE: cpe:2.3:a:urbackup:urbackup:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.urbackup.org/changelog.html
DownloadPageUrl: https://www.urbackup.org/download.html
IconSource: Installer
UpdateCadenceDays: 180

.SYNOPSIS
    Packages the UrBackup Client (x64 MSI) for MECM.

.DESCRIPTION
    Resolves the newest stable x64 client MSI from the vendor download page,
    stages content to a versioned local folder with ARP detection metadata,
    and creates an MECM Application with registry-based detection.

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
    Content is staged under: <FileServerPath>\Applications\UrBackup\UrBackup Client\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\UrBackup Client).
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
    Outputs only the latest available UrBackup Client version string and exits.

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
$DownloadPageUrl = "https://www.urbackup.org/download.html"

$VendorFolder = "UrBackup"
$AppFolder    = "UrBackup Client"

$BaseDownloadRoot = Join-Path $DownloadRoot "UrBackup Client"
$MsiFileName      = "urbackup-client-x64.msi"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        A mirror that answers 200 with an HTML error body would otherwise stage
        as a valid-looking MSI and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestUrBackupClient {
    <#
    .SYNOPSIS
        Returns the newest stable UrBackup Client version and its x64 MSI URL.
    .DESCRIPTION
        The download page lists release candidates and a tray-less build beside
        the stable client. The version segment must be followed immediately by
        the x64 marker, which excludes both "2.5.33RC3" directories and the
        "(No tray)" filenames. Links are protocol-relative.
    #>
    param([switch]$Quiet)

    Write-Log "UrBackup download page       : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch UrBackup download page: $DownloadPageUrl" }

        $rx = [regex]'href\s*=\s*"(?<href>[^"]*?/Client/(?<ver>\d+\.\d+\.\d+)/UrBackup%20Client%20\d+\.\d+\.\d+%28x64%29\.msi)"'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate a stable x64 client MSI link on the download page."
        }

        $candidates = foreach ($m in $rxMatches) {
            [pscustomobject]@{
                Href    = $m.Groups['href'].Value
                Version = $m.Groups['ver'].Value
            }
        }

        $best = $candidates | Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1

        $url = $best.Href
        if ($url -like '//*')      { $url = "https:$url" }
        elseif ($url -notlike 'http*') { $url = ([uri]::new([uri]$DownloadPageUrl, $url)).AbsoluteUri }

        Write-Log "Latest UrBackup Client       : $($best.Version)" -Quiet:$Quiet
        Write-Log "Resolved MSI URL             : $url" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get UrBackup Client version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageUrBackupClient {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UrBackup Client (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestUrBackupClient
    if (-not $releaseInfo) { throw "Could not resolve UrBackup Client version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""
    Write-Log "Downloading MSI..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localMsi

    Assert-MsiPayload -Path $localMsi

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
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
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
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "UrBackup" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "UrBackup Client"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("UrBackupClient", "UrBackupClientBackend")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
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

function Invoke-PackageUrBackupClient {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UrBackup Client (x64) - PACKAGE phase"
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
        $info = Get-LatestUrBackupClient -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("UrBackup Client GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UrBackup Client (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageUrBackupClient
    }
    elseif ($PackageOnly) {
        Invoke-PackageUrBackupClient
    }
    else {
        Invoke-StageUrBackupClient
        Invoke-PackageUrBackupClient
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-urbackupclient'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
