<#
Vendor: Acro Software Inc.
App: CutePDF Writer
CMName: CutePDF Writer
VendorUrl: https://www.cutepdf.com/
CPE: cpe:2.3:a:acrosoftware:cutepdf_writer:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.cutepdf.com/Products/CutePDF/writer.asp
DownloadPageUrl: https://www.cutepdf.com/products/cutepdf/writer.asp
UpdateCadenceDays: 180

.SYNOPSIS
    Packages CutePDF Writer (Inno Setup EXE) for MECM.

.DESCRIPTION
    Downloads the current CuteWriter.exe from the vendor's static download URL,
    reads its ProductVersion, stages content to a versioned local folder, and
    creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, read EXE version, generate wrappers and manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: The vendor publishes no version feed and the download URL always
    serves the current build, so the version is read from the downloaded EXE's
    version resource. GetLatestVersionOnly therefore downloads the installer.

    The Inno Setup uninstaller is registered under a fixed ARP key rather than
    a ProductCode, and the staged installer accepts /uninstall, so the uninstall
    wrapper re-runs the staged EXE instead of calling unins000.exe from a path
    that varies by build.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Acro Software\CutePDF Writer\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\CutePDFWriter).
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
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Downloads the installer, outputs its ProductVersion, and exits.

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
# The vendor's download host does not serve this path over TLS; the installer
# is version-checked after download, so a tampered payload cannot reach the
# staged folder without changing the recorded version and file hash.
$ExeDownloadUrl   = "http://www.cutepdf.com/download/CuteWriter.exe"
$CacheFileName    = "CuteWriter.exe"

$VendorFolder = "Acro Software"
$AppFolder    = "CutePDF Writer"

$BaseDownloadRoot = Join-Path $DownloadRoot "CutePDFWriter"

$DetectionRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CutePDF Writer Installation"

# --- Functions ---


function Get-CutePdfWriterExeVersion {
    param([Parameter(Mandatory)][string]$ExePath)

    $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
    $version = $info.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) { $version = $info.FileVersion }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Could not read a ProductVersion or FileVersion from $ExePath."
    }

    return $version.Trim()
}


function Get-CutePdfWriterUninstallContent {
    param([Parameter(Mandatory)][string]$InstallerFileName)

    $escaped = $InstallerFileName -replace "'", "''"
    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $escaped),
        'if (-not (Test-Path -LiteralPath $exePath)) { exit 1 }',
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/SP-'', ''/VERYSILENT'', ''/SUPPRESSMSGBOXES'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCutePdfWriter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CutePDF Writer - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $CacheFileName
    Write-Log "Download URL                 : $ExeDownloadUrl"
    Write-Log "Local EXE path               : $localExe"
    Write-Log ""
    Write-Log "Downloading CutePDF Writer installer..."
    Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe

    # --- Version ---
    $version = Get-CutePdfWriterExeVersion -ExePath $localExe
    $installerFileName = "CuteWriter-$version.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'" `
        -UninstallCommand 'unused'

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content (Get-CutePdfWriterUninstallContent -InstallerFileName $installerFileName)

    # --- Write stage manifest ---
    Write-Log "Detection key                : HKLM\$DetectionRegistryKey"
    Write-Log "Detection value              : DisplayVersion >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "CutePDF Writer $version"
        Publisher       = "Acro Software Inc."
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallArgs   = "/uninstall /SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @()
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $DetectionRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
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

function Invoke-PackageCutePdfWriter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CutePDF Writer - PACKAGE phase"
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
        Initialize-Folder -Path $BaseDownloadRoot

        $localExe = Join-Path $BaseDownloadRoot $CacheFileName
        Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe -Quiet

        Write-Output (Get-CutePdfWriterExeVersion -ExePath $localExe)
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("CutePDF Writer GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CutePDF Writer Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ExeDownloadUrl               : $ExeDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCutePdfWriter
    }
    elseif ($PackageOnly) {
        Invoke-PackageCutePdfWriter
    }
    else {
        Invoke-StageCutePdfWriter
        Invoke-PackageCutePdfWriter
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-cutepdfwriter'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
