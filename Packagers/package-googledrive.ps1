<#
Vendor: Google
App: Google Drive for desktop
CMName: Google Drive
VendorUrl: https://www.google.com/drive/download/
CPE: cpe:2.3:a:google:drive_for_desktop:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://support.google.com/a/answer/7577057
DownloadPageUrl: https://support.google.com/a/answer/7577057
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Google Drive for desktop (x64) for MECM.

.DESCRIPTION
    Downloads GoogleDriveSetup.exe from Google's evergreen download URL, stages
    content to a versioned local folder with ARP detection metadata, and creates
    an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, read the installer file version, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: Google publishes no version endpoint for Drive for desktop, so the
    version is the installer's own file version. The installer is ~270 MB, so a
    cached copy is refreshed conditionally on the published Last-Modified date
    rather than re-fetched on every version check.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Google\Google Drive\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\GoogleDrive).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers and
    stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Google Drive version string and exits.

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
$SetupDownloadUrl = "https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"
$SetupFileName    = "GoogleDriveSetup.exe"

$VendorFolder = "Google"
$AppFolder    = "Google Drive"

$BaseDownloadRoot = Join-Path $DownloadRoot "GoogleDrive"

# The installer registers Add/Remove Programs under a fixed product GUID
# (an x64 process, so the entry is in the native registry view).
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6BBAE539-2232-434A-A4E5-9A33560C6283}"

# --- Functions ---


function Assert-DrivePayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        An edge node that answers 200 with an HTML error body would otherwise
        stage as a valid-looking EXE and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-GoogleDriveInstaller {
    <#
    .SYNOPSIS
        Returns the local installer path and its file version, refreshing the
        cached copy only when the published file is newer.
    .DESCRIPTION
        The download is ~270 MB and the URL is evergreen, so an unconditional
        fetch on every version check would move a quarter of a gigabyte to learn
        nothing. The cached file keeps the server's Last-Modified timestamp so
        the conditional request stays accurate across runs.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot
    $localExe = Join-Path $BaseDownloadRoot $SetupFileName

    if (Test-Path -LiteralPath $localExe) {
        Write-Log "Cached installer present; refreshing only if newer." -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $SetupDownloadUrl -OutFile $localExe -Quiet:$Quiet `
            -ExtraCurlArgs @('-R', '-z', $localExe)
    }
    else {
        Write-Log "Downloading installer        : $SetupDownloadUrl" -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $SetupDownloadUrl -OutFile $localExe -Quiet:$Quiet -ExtraCurlArgs @('-R')
    }

    Assert-DrivePayloadIsExecutable -Path $localExe

    $version = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Installer carries no file version: $localExe"
    }
    $version = $version.Trim()

    Write-Log "Google Drive version         : $version" -Quiet:$Quiet

    return [pscustomobject]@{
        Path    = $localExe
        Version = $version
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGoogleDrive {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Google Drive (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $installer = Get-GoogleDriveInstaller
    $version   = $installer.Version

    Write-Log "Local installer path         : $($installer.Path)"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $SetupFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $installer.Path -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedExe"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # --silent installs per-machine without launching the client; the shortcut
    # switches keep the install from writing desktop and Office entry points
    # that a managed image does not want.
    $installArgs = "'--silent', '--desktop_shortcut=false', '--gsuite_shortcuts=false'"
    $wrappers = New-ExeWrapperContent -InstallerFileName $SetupFileName `
        -InstallArgs $installArgs `
        -UninstallCommand 'unused'

    # The uninstaller lives in a version-numbered directory under the install
    # root, so the ARP UninstallString is the only value that names the copy
    # that is actually installed.
    $uninstallContent = @'
$key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6BBAE539-2232-434A-A4E5-9A33560C6283}'
if (-not (Test-Path -LiteralPath $key)) { exit 0 }
$cmd = (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue).UninstallString
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = ($cmd -split '\s+--')[0].Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('--silent', '--force_stop') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Google Drive"
        Publisher       = "Google LLC"
        SoftwareVersion = $version
        InstallerFile   = $SetupFileName
        InstallerType   = "EXE"
        InstallArgs     = "--silent --desktop_shortcut=false --gsuite_shortcuts=false"
        UninstallArgs   = "--silent --force_stop"
        RunningProcess  = @("GoogleDriveFS")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
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

function Invoke-PackageGoogleDrive {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Google Drive (x64) - PACKAGE phase"
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
        $info = Get-GoogleDriveInstaller -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Google Drive GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Google Drive (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "SetupDownloadUrl             : $SetupDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGoogleDrive
    }
    elseif ($PackageOnly) {
        Invoke-PackageGoogleDrive
    }
    else {
        Invoke-StageGoogleDrive
        Invoke-PackageGoogleDrive
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-googledrive'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
