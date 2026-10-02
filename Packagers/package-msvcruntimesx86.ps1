<#
Vendor: Microsoft
App: Microsoft Visual C++ v14 Redistributable (x86)
CMName: Microsoft Visual C++ v14 Redistributable (x86)
VendorUrl: https://learn.microsoft.com/cpp/windows/latest-supported-vc-redist
CPE: cpe:2.3:a:microsoft:visual_c%2b%2b_redistributable:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist
DownloadPageUrl: https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist
IconSource: None
WsusSupport: Yes

.SYNOPSIS
    Packages Microsoft Visual C++ v14 Redistributable (x86) for ConfigMgr.

.DESCRIPTION
    Downloads the latest vc_redist.x86.exe from Microsoft's permalink URL,
    reads the version from VersionInfo, stages it to a versioned local folder,
    and creates a ConfigMgr Application with file version detection.
    Detection requires %SystemRoot%\SysWOW64\vcruntime140.dll at or above the
    packaged version. The runtime DLL carries the redistributable version, so
    the same detection also finds an older installed copy. The detection
    reads SysWOW64 directly, so it requires 64-bit Windows.

    The x64 runtime is a separate application (package-msvcruntimesx64.ps1).
    package-msvcruntimes.ps1 installs both architectures as one application.

    NOTE: The aka.ms permalink URL always serves the current release. The
    installer is always re-downloaded to ensure the latest version is packaged.

    GetLatestVersionOnly downloads only the installer to a local staging
    folder, reads the version from VersionInfo, outputs the short version
    string, and exits.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\VC++ v14 Redistributable x86\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\MsvcRedistX86).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file version detection.

.PARAMETER GetLatestVersionOnly
    Downloads the installer, reads the version from VersionInfo, outputs the
    short version string, and exits. No ConfigMgr changes are made.

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
$Url      = "https://aka.ms/vc14/vc_redist.x86.exe"
$FileName = "vc_redist.x86.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "VC++ v14 Redistributable x86"

$BaseDownloadRoot = Join-Path $DownloadRoot "MsvcRedistX86"

# SysWOW64 with the 64-bit flag, not System32 with the 32-bit flag: a 32-bit
# ConfigMgr file clause falls back to the 64-bit System32, where the x64
# runtime carries the same file name.
$DetectionFolder = "%SystemRoot%\SysWOW64"
$DetectionFile   = "vcruntime140.dll"

# --- Functions ---


function Get-ExeFileVersion {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Quiet
    )

    $vi = (Get-Item -LiteralPath $Path).VersionInfo
    if (-not $vi) { throw "Could not read VersionInfo from: $Path" }

    $fv = $vi.FileVersion
    $pv = $vi.ProductVersion

    if (-not [string]::IsNullOrWhiteSpace($fv)) { $fv = $fv.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($pv)) { $pv = $pv.Trim() }

    Write-Log "EXE FileVersion              : $fv" -Quiet:$Quiet
    Write-Log "EXE ProductVersion           : $pv" -Quiet:$Quiet

    if ($fv -match '^\d+\.\d+\.\d+\.\d+$') { return $fv }
    if ($pv -match '^\d+\.\d+\.\d+\.\d+$') { return $pv }

    throw "Could not determine quad version from VersionInfo for: $Path"
}

function Get-ShortVersionFromQuad {
    param([Parameter(Mandatory)][string]$QuadVersion)

    $parts = $QuadVersion -split '\.'
    if ($parts.Count -lt 3) { throw "Unexpected version format: $QuadVersion" }
    return ("{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2])
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageMsvcRedistX86 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC v14 Redistributable (x86) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # Always re-download (permalink URL serves the latest release)
    $localExe = Join-Path $BaseDownloadRoot $FileName

    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $Url -OutFile $localExe

    $quadVersion  = Get-ExeFileVersion -Path $localExe
    $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion

    Write-Log ""
    Write-Log "Version (short)              : $shortVersion"
    Write-Log "Version (quad)               : $quadVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $shortVersion
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $FileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $FileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/install'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $FileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent `
        -InstallBatExitCode '3010' `
        -UninstallBatExitCode '3010'

    # --- Write stage manifest ---
    $appName   = "Microsoft Visual C++ v14 Redistributable (x86) - $shortVersion"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection                    : $DetectionFolder\$DetectionFile >= $quadVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName               = $appName
        Publisher             = $publisher
        SoftwareVersion       = $shortVersion
        InstallerFile         = $FileName
        InstallerType         = "EXE"
        InstallArgs           = "/install /quiet /norestart"
        UninstallArgs         = "/uninstall /quiet /norestart"
        RunningProcess        = @()
        PostExecutionBehavior = "ForceReboot"
        Detection             = @{
            Type          = "File"
            FilePath      = $DetectionFolder
            FileName      = $DetectionFile
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $quadVersion
            Is64Bit       = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $shortVersion -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageMsvcRedistX86 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC v14 Redistributable (x86) - PACKAGE phase"
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
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
    Write-Log "PostExecutionBehavior        : $($manifest.PostExecutionBehavior)"
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

    # --- ConfigMgr application ---
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

        $localExe = Join-Path $BaseDownloadRoot $FileName

        Invoke-DownloadWithRetry -Url $Url -OutFile $localExe -Quiet

        $quadVersion = Get-ExeFileVersion -Path $localExe -Quiet
        $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion

        Write-Output $shortVersion
        exit 0
    }
    catch {
        Write-Log $_.Exception.Message -Level ERROR
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC v14 Redistributable (x86) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "Url                          : $Url"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageMsvcRedistX86
    }
    elseif ($PackageOnly) {
        Invoke-PackageMsvcRedistX86
    }
    else {
        Invoke-StageMsvcRedistX86
        Invoke-PackageMsvcRedistX86
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-msvcruntimesx86'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
