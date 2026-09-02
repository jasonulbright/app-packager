<#
Vendor: WireGuard
App: WireGuard
CMName: WireGuard
VendorUrl: https://www.wireguard.com/
CPE: cpe:2.3:a:wireguard:wireguard:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://git.zx2c4.com/wireguard-windows/log/
DownloadPageUrl: https://download.wireguard.com/windows-client/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages WireGuard for Windows (amd64) MSI for MECM.

.DESCRIPTION
    Resolves the newest amd64 MSI from the official windows-client download
    index, stages content to a versioned local folder with ARP detection derived
    from the MSI properties, and creates an MECM Application with registry-based
    detection.

    The index also carries wireguard-installer.exe (a downloader stub) and the
    arm64/x86 MSIs; the pattern is anchored so only the amd64 MSI is taken.

    The install passes DO_NOT_LAUNCH=1 so the MSI does not start the tunnel UI
    in the SYSTEM context after a silent deployment.

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
    Content is staged under: <FileServerPath>\Applications\WireGuard\WireGuard\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\WireGuard).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 5

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available WireGuard version string and exits.

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
    [int]$EstimatedRuntimeMins = 5,
    [int]$MaximumRuntimeMins = 15,
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
$DownloadIndexUrl = "https://download.wireguard.com/windows-client/"

$VendorFolder = "WireGuard"
$AppFolder    = "WireGuard"

$BaseDownloadRoot = Join-Path $DownloadRoot "WireGuard"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        A CDN error page answers 200 with HTML, which would otherwise stage as a
        valid-looking MSI.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestWireGuardRelease {
    <#
    .SYNOPSIS
        Returns the newest amd64 MSI version and its download URL.
    .DESCRIPTION
        The index lists every published MSI rather than a "latest" link, so the
        candidates are ordered by parsed version: a lexical sort would rank
        1.1 above 1.10.
    #>
    param([switch]$Quiet)

    Write-Log "Download index URL           : $DownloadIndexUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadIndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the WireGuard download index." }

        $rx = [regex]'href\s*=\s*"(?<file>wireguard-amd64-(?<ver>\d+(?:\.\d+)+)\.msi)"'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any amd64 MSI links on the download index."
        }

        $candidates = foreach ($m in $rxMatches) {
            [pscustomobject]@{
                FileName = $m.Groups["file"].Value
                Version  = $m.Groups["ver"].Value
                Sortable = [version]$m.Groups["ver"].Value
            }
        }

        $best = $candidates | Sort-Object Sortable -Descending | Select-Object -First 1
        $url = ([uri]::new([uri]$DownloadIndexUrl, $best.FileName)).AbsoluteUri

        Write-Log "Latest WireGuard version     : $($best.Version)" -Quiet:$Quiet
        Write-Log "Installer filename           : $($best.FileName)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            FileName    = $best.FileName
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get WireGuard version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWireGuard {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WireGuard (amd64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestWireGuardRelease
    if (-not $releaseInfo) { throw "Could not resolve WireGuard version." }

    $version     = $releaseInfo.Version
    $msiFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

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

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    # For a per-machine MSI install the ARP uninstall key name is the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName -ExtraInstallArgs @('DO_NOT_LAUNCH=1')
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "WireGuard" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "WireGuard"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart DO_NOT_LAUNCH=1"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("wireguard")
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

function Invoke-PackageWireGuard {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WireGuard (amd64) - PACKAGE phase"
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
        $info = Get-LatestWireGuardRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("WireGuard GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WireGuard (amd64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadIndexUrl             : $DownloadIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageWireGuard
    }
    elseif ($PackageOnly) {
        Invoke-PackageWireGuard
    }
    else {
        Invoke-StageWireGuard
        Invoke-PackageWireGuard
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-wireguard'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
