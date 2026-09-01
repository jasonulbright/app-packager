<#
Vendor: Raspberry Pi Ltd
App: Raspberry Pi Imager
CMName: Raspberry Pi Imager
VendorUrl: https://www.raspberrypi.com/software/
ReleaseNotesUrl: https://github.com/raspberrypi/rpi-imager/releases
DownloadPageUrl: https://github.com/raspberrypi/rpi-imager/releases
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Raspberry Pi Imager (x64) for MECM.

.DESCRIPTION
    Downloads the latest Raspberry Pi Imager Windows installer from GitHub
    releases, stages content to a versioned local folder with ARP detection
    metadata, and creates an MECM Application with registry-based detection.

    The installer is an Inno Setup package that requires elevation and always
    installs per-machine into Program Files.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection, write manifest
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
    Outputs only the latest available Raspberry Pi Imager version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - Target clients need Windows 10 build 15063 or later; the installer
      enforces this as a minimum-version condition.
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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$GitHubApiUrl = "https://api.github.com/repos/raspberrypi/rpi-imager/releases/latest"

$VendorFolder = "Raspberry Pi Ltd"
$AppFolder    = "Raspberry Pi Imager"

$BaseDownloadRoot = Join-Path $DownloadRoot "Raspberry Pi Imager"

# Inno Setup appends _is1 to the package AppId to form the ARP key name.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{46C75BF3-E267-4834-AE1D-BEB77E05BFA1}_is1"

$InstallDir      = "C:\Program Files\Raspberry Pi Ltd\Imager"
$UninstallerPath = Join-Path $InstallDir "unins000.exe"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        A release-asset redirect that answers 200 with an HTML error body would
        otherwise stage as a valid-looking installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestRpiImagerRelease {
    <#
    .SYNOPSIS
        Returns the newest Raspberry Pi Imager release and its Windows installer asset.
    .DESCRIPTION
        The release carries Debian, macOS and source assets alongside the
        Windows installer, so the asset filter pins the imager-*.exe name.
        The tag string is carried through unchanged as well: the installer
        stamps it verbatim as its version, leading "v" included, and that is
        what the ARP entry records.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        $tag = [string]$release.tag_name
        if ([string]::IsNullOrWhiteSpace($tag)) {
            throw "Could not parse version from GitHub release tag."
        }
        $version = $tag -replace '^v', ''

        $asset = $release.assets |
            Where-Object { $_.name -match '^imager-.+\.exe$' } |
            Select-Object -First 1
        if (-not $asset) { throw "No Windows installer asset found in release $version." }

        Write-Log "Latest Imager version        : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version        = $version
            DisplayVersion = $tag
            FileName       = $asset.name
            DownloadUrl    = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get Raspberry Pi Imager version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageRpiImager {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Raspberry Pi Imager (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestRpiImagerRelease
    if (-not $releaseInfo) { throw "Could not resolve Raspberry Pi Imager version." }

    $version     = $releaseInfo.Version
    $exeFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $exeFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $exeFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

    # The installer stamps its Inno AppVersion into the file's ProductVersion,
    # and Inno writes that same string to the ARP DisplayVersion value.
    $stampedVersion = (Get-Item -LiteralPath $localExe).VersionInfo.ProductVersion
    if ($stampedVersion) { $stampedVersion = $stampedVersion.Trim() }
    if ([string]::IsNullOrWhiteSpace($stampedVersion)) { $stampedVersion = $releaseInfo.DisplayVersion }

    Write-Log "Installer ProductVersion     : $stampedVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $exeFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedExe"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $stampedVersion"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-ExeWrapperContent -InstallerFileName $exeFileName `
        -InstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-'" `
        -UninstallCommand $UninstallerPath `
        -UninstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Raspberry Pi Imager"
        Publisher       = "Raspberry Pi Ltd"
        SoftwareVersion = $version
        InstallerFile   = $exeFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("rpi-imager")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $stampedVersion
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

function Invoke-PackageRpiImager {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Raspberry Pi Imager (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

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
        $info = Get-LatestRpiImagerRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Raspberry Pi Imager GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Raspberry Pi Imager (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "GitHubApiUrl                 : $GitHubApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageRpiImager
    }
    elseif ($PackageOnly) {
        Invoke-PackageRpiImager
    }
    else {
        Invoke-StageRpiImager
        Invoke-PackageRpiImager
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-rpiimager'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
