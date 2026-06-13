<#
Vendor: NVIDIA
App: NVIDIA Graphics Driver - GeForce Game Ready (x64)
CMName: NVIDIA Graphics Driver - GeForce Game Ready
VendorUrl: https://www.nvidia.com/Download/index.aspx
CPE: cpe:2.3:a:nvidia:gpu_display_driver:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.nvidia.com/en-us/drivers/drivers-faq/
DownloadPageUrl: https://www.nvidia.com/Download/index.aspx
UpdateCadenceDays: 30

.SYNOPSIS
    Packages the latest NVIDIA GeForce Game Ready DCH driver (x64) for MECM.

.DESCRIPTION
    Queries NVIDIA's AjaxDriverService.php JSON endpoint with pinned psid/pfid
    for the GeForce RTX 50 Series flagship (covers all current Maxwell+
    GeForce GTX/RTX cards via the unified DCH driver), downloads the latest
    Game Ready WHQL installer, stages content to a versioned local folder,
    and creates an MECM Application with ARP registry-based detection on
    the constant NVIDIA Display.Driver uninstall GUID.

    Supports two-phase operation:
      -StageOnly    Resolve latest version, download installer, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The psid/pfid pinning answers the typical "combo box ambiguity" of the
    NVIDIA download page: one DCH driver covers the whole current GeForce
    line, so the specific flagship pfid is just a key into the version
    lookup -- the resulting installer is identical regardless of which
    current-series pfid you pin.

    GetLatestVersionOnly issues a single JSON call (no installer download)
    and exits.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under:
      <FileServerPath>\Applications\NVIDIA\NVIDIA Graphics Driver - GeForce Game Ready\<Version>

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
    Queries the NVIDIA driver API for the current Game Ready version, outputs
    the version string, and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - Local administrator
    - Write access to FileServerPath
    - Outbound HTTPS to gfwsl.geforce.com and us.download.nvidia.com
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
# psid 131 = GeForce RTX 50 Series; pfid 1066 = RTX 5090.
# The DCH driver is unified across all current GeForce GPUs, so this pfid is
# a stable lookup key for "the latest Game Ready driver for current cards"
# rather than a per-GPU selector. Update only if NVIDIA retires PSID 131.
$NvidiaApiBaseUrl = "https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php"
$NvidiaPsid       = 131
$NvidiaPfid       = 1066
$NvidiaOsId       = 57       # Windows 10 64-bit (DCH driver covers Win10 + Win11 in one package)
$NvidiaLangId     = 1033     # English - United States
$NvidiaDch        = 1        # DCH driver (required for Windows 10 1809+)
$NvidiaUpCrd      = 0        # Game Ready branch (1 = Production Branch / Enterprise)

# Constant ARP uninstall GUID for the NVIDIA Display.Driver component on
# DCH installs. Same GUID for Game Ready and RTX Enterprise -- the package
# name differs only by AppFolder/AppName, so the two MECM apps coexist on
# the share without colliding.
$NvidiaDisplayDriverArpKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}_Display.Driver"

$VendorFolder = "NVIDIA"
$AppFolder    = "NVIDIA Graphics Driver - GeForce Game Ready"

$BaseDownloadRoot = Join-Path $DownloadRoot "NvidiaGeForce"

# --- Functions ---


function Resolve-NvidiaGeForceLatest {
    <#
    .SYNOPSIS
        Calls AjaxDriverService.php and returns @{ Version; DownloadUrl; InstallerFileName }.
    #>
    param([switch]$Quiet)

    $url = "{0}?func=DriverManualLookup&psid={1}&pfid={2}&osID={3}&languageCode={4}&beta=0&isWHQL=1&dltype=-1&dch={5}&upCRD={6}&qnf=0&sort1=0&numberOfResults=10" -f `
        $NvidiaApiBaseUrl, $NvidiaPsid, $NvidiaPfid, $NvidiaOsId, $NvidiaLangId, $NvidiaDch, $NvidiaUpCrd

    Write-Log "NVIDIA driver API URL        : $url" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $url) -join ''
        if ($LASTEXITCODE -ne 0) { throw "AjaxDriverService.php returned exit $LASTEXITCODE" }

        $data = ConvertFrom-Json $json
        if (-not $data.IDS -or @($data.IDS).Count -eq 0) {
            throw "NVIDIA driver lookup returned no results (psid=$NvidiaPsid, pfid=$NvidiaPfid, osID=$NvidiaOsId, upCRD=$NvidiaUpCrd)."
        }

        $latest = $data.IDS[0].downloadInfo
        $version = [string]$latest.Version
        $downloadUrl = [string]$latest.DownloadURL

        if ([string]::IsNullOrWhiteSpace($version)) { throw "downloadInfo.Version missing in API response." }
        if ([string]::IsNullOrWhiteSpace($downloadUrl)) { throw "downloadInfo.DownloadURL missing in API response." }

        $installerFileName = [System.IO.Path]::GetFileName($downloadUrl)

        Write-Log "Latest GeForce driver version: $version"   -Quiet:$Quiet
        Write-Log "Driver name                  : $($latest.Name)" -Quiet:$Quiet
        Write-Log "Release date                 : $($latest.ReleaseDateTime)" -Quiet:$Quiet
        Write-Log "Download URL                 : $downloadUrl" -Quiet:$Quiet
        Write-Log "Installer file               : $installerFileName" -Quiet:$Quiet

        return @{
            Version           = $version
            DownloadUrl       = $downloadUrl
            InstallerFileName = $installerFileName
        }
    }
    catch {
        Write-Log "Failed to resolve NVIDIA GeForce driver: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageNvidiaGeForce {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVIDIA Graphics Driver (GeForce Game Ready) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Resolve-NvidiaGeForceLatest
    if (-not $release) { throw "Could not resolve latest NVIDIA GeForce driver." }

    $version       = $release.Version
    $installerName = $release.InstallerFileName
    $downloadUrl   = $release.DownloadUrl

    # --- Download ---
    $localInstaller = Join-Path $BaseDownloadRoot $installerName
    Write-Log "Local installer path         : $localInstaller"
    Write-Log ""
    Write-Log "Downloading installer (~1 GB, this can take several minutes)..."
    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedInstaller = Join-Path $localContentPath $installerName
    if (-not (Test-Path -LiteralPath $stagedInstaller)) {
        Copy-Item -LiteralPath $localInstaller -Destination $stagedInstaller -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedInstaller"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # NVIDIA DCH installer silent switches:
    #   -s         silent
    #   -noreboot  do not auto-reboot (let MECM handle)
    #   -clean     clean install: removes prior driver settings + profiles
    # Uninstall:
    #   -uninstall -s -noreboot
    $installPs1 = @"
`$exePath = Join-Path `$PSScriptRoot '$installerName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('-s','-noreboot','-clean') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallPs1 = @"
`$exePath = Join-Path `$PSScriptRoot '$installerName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('-uninstall','-s','-noreboot') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Stage manifest ---
    $publisher = "NVIDIA Corporation"
    $appName   = "NVIDIA Graphics Driver - GeForce Game Ready $version"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = $appName
        Publisher        = $publisher
        SoftwareVersion  = $version
        InstallerFile    = $installerName
        InstallerType    = "EXE"
        InstallArgs      = "-s -noreboot -clean"
        UninstallCommand = $installerName
        UninstallArgs    = "-uninstall -s -noreboot"
        RunningProcess   = @("nvcontainer","NVDisplay.Container","nvsphelper64","NVIDIA Web Helper")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $NvidiaDisplayDriverArpKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
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

function Invoke-PackageNvidiaGeForce {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVIDIA Graphics Driver (GeForce Game Ready) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    $release = Resolve-NvidiaGeForceLatest -Quiet
    if (-not $release) { throw "Could not resolve latest NVIDIA GeForce driver for manifest lookup." }

    $localContentPath = Join-Path $BaseDownloadRoot $release.Version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "Stage manifest not found - run Stage phase first: $manifestPath"
    }

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot     = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

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
        $release = Resolve-NvidiaGeForceLatest -Quiet
        if (-not $release) { exit 1 }
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}


# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVIDIA GeForce Game Ready Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "NvidiaApiBaseUrl             : $NvidiaApiBaseUrl"
    Write-Log "Pinned psid/pfid             : $NvidiaPsid / $NvidiaPfid (Game Ready branch)"
    Write-Log ""

    if ($StageOnly)       { Invoke-StageNvidiaGeForce }
    elseif ($PackageOnly) { Invoke-PackageNvidiaGeForce }
    else                  { Invoke-StageNvidiaGeForce; Invoke-PackageNvidiaGeForce }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-nvidia-geforce'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
