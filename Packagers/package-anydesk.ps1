<#
Vendor: AnyDesk Software GmbH
App: AnyDesk
CMName: AnyDesk
VendorUrl: https://anydesk.com/
CPE: cpe:2.3:a:anydesk:anydesk:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://anydesk.com/en/changelog/windows
DownloadPageUrl: https://anydesk.com/en/downloads/windows
UpdateCadenceDays: 60

.SYNOPSIS
    Packages AnyDesk (x86 payload, 64-bit capable host) for MECM.

.DESCRIPTION
    Downloads the current AnyDesk.exe from the vendor download endpoint, reads
    the version from the binary's FileVersion resource, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection on the installed AnyDesk.exe.

    AnyDesk ships a single self-contained EXE that acts as both the installer
    and the installed program; the same binary performs the uninstall.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\AnyDesk\AnyDesk\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AnyDesk).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available AnyDesk version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$DownloadUrl = "https://download.anydesk.com/AnyDesk.exe"

$VendorFolder = "AnyDesk"
$AppFolder    = "AnyDesk"

$BaseDownloadRoot  = Join-Path $DownloadRoot "AnyDesk"
$InstallerFileName = "AnyDesk.exe"

# AnyDesk installs to the 32-bit Program Files tree regardless of host bitness.
$InstallDir = "${env:ProgramFiles(x86)}\AnyDesk"

# --- Functions ---


function Get-AnyDeskDownload {
    <#
    .SYNOPSIS
        Downloads AnyDesk.exe and returns its path and FileVersion.
    .DESCRIPTION
        The vendor publishes no version feed for the Windows build; the
        authoritative version is the FileVersion resource of the shipped
        binary, so the download is the version check.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName

    Write-Log "Download URL                 : $DownloadUrl" -Quiet:$Quiet
    Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe -Quiet:$Quiet

    $fileVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($fileVersion)) {
        throw "AnyDesk.exe carries no FileVersion resource: $localExe"
    }
    $fileVersion = $fileVersion.Trim()

    Write-Log "AnyDesk FileVersion          : $fileVersion" -Quiet:$Quiet

    return [pscustomobject]@{
        Path    = $localExe
        Version = $fileVersion
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAnyDesk {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyDesk - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download and read version ---
    $download = Get-AnyDeskDownload
    $version  = $download.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $download.Path -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installArgs = "'--install', '{0}', '--start-with-win', '--create-shortcuts', '--create-desktop-icon', '--silent'" -f $InstallDir

    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand 'unused'

    # The installed AnyDesk.exe removes itself; the staged installer copy is
    # not present on the client at uninstall time.
    $uninstallContent = @'
$exePath = Join-Path ${env:ProgramFiles(x86)} 'AnyDesk\AnyDesk.exe'
if (-not (Test-Path -LiteralPath $exePath)) { exit 0 }
$proc = Start-Process -FilePath $exePath -ArgumentList @('--remove', '--silent') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "AnyDesk"
    $publisher = "AnyDesk Software GmbH"

    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : AnyDesk.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "--install `"$InstallDir`" --start-with-win --create-shortcuts --create-desktop-icon --silent"
        UninstallArgs   = "--remove --silent"
        RunningProcess  = @("AnyDesk")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "AnyDesk.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
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

function Invoke-PackageAnyDesk {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyDesk - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $download = Get-AnyDeskDownload -Quiet
        if (-not $download) { exit 1 }
        Write-Output $download.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("AnyDesk GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyDesk Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadUrl                  : $DownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAnyDesk
    }
    elseif ($PackageOnly) {
        Invoke-PackageAnyDesk
    }
    else {
        Invoke-StageAnyDesk
        Invoke-PackageAnyDesk
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-anydesk'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
