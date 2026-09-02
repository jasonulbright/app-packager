<#
Vendor: Calibrite
App: Calibrite PROFILER
CMName: Calibrite PROFILER
VendorUrl: https://calibrite.com/
CPE: cpe:2.3:a:calibrite:profiler:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/LUMESCA/calibrite-profiler-releases/releases
DownloadPageUrl: https://calibrite.com/us/software-downloads/
UpdateCadenceDays: 120

.SYNOPSIS
    Packages Calibrite PROFILER (x64) for MECM.

.DESCRIPTION
    Reads the current release from the vendor's public release repository,
    downloads the matching electron-builder NSIS installer, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection on the installed application executable.

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
    Content is staged under: <FileServerPath>\Applications\Calibrite\Calibrite PROFILER\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Calibrite PROFILER).
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
    Outputs only the latest available Calibrite PROFILER version string and exits.

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
$DownloadPageUrl = "https://calibrite.com/us/software-downloads/"
$ReleaseApiUrl   = "https://api.github.com/repos/LUMESCA/calibrite-profiler-releases/releases/latest"

$VendorFolder = "Calibrite"
$AppFolder    = "Calibrite PROFILER"

$BaseDownloadRoot = Join-Path $DownloadRoot "Calibrite PROFILER"

# electron-builder installs the machine-wide x64 build under the 64-bit
# Program Files using the product name verbatim, including its lower-case "calibrite".
$InstallDir      = "{0}\calibrite PROFILER" -f $env:ProgramFiles
$DetectionFile   = "calibrite PROFILER.exe"
$UninstallerPath = "{0}\Uninstall calibrite PROFILER.exe" -f $InstallDir

# --- Functions ---


function Assert-CalibritePayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A release asset that has been renamed or withdrawn answers with an HTML
        error page, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Resolve-CalibriteRelease {
    <#
    .SYNOPSIS
        Returns the version and Windows installer URL of the current release.
    .DESCRIPTION
        Each release carries macOS artifacts alongside the Windows installer, so
        the asset is selected by the Setup-<version>.exe filename rather than by
        position. The blockmap sibling shares the .exe prefix and is excluded.
    #>
    param([switch]$Quiet)

    Write-Log "Release API URL              : $ReleaseApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error -H "Accept: application/vnd.github+json" $ReleaseApiUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Calibrite PROFILER release metadata." }

        $release = $json | ConvertFrom-Json
        if (-not $release -or -not $release.assets) { throw "Release metadata contained no assets." }

        $asset = $release.assets |
            Where-Object { $_.name -match '^calibrite-PROFILER-Setup-\d+(\.\d+)+\.exe$' } |
            Select-Object -First 1

        if (-not $asset) { throw "Release $($release.tag_name) has no Windows Setup asset." }

        $version = [regex]::Match($asset.name, '(?<ver>\d+(?:\.\d+)+)').Groups['ver'].Value
        if ([string]::IsNullOrWhiteSpace($version)) { throw "Could not parse a version from asset name: $($asset.name)" }

        Write-Log "Latest Calibrite version     : $version" -Quiet:$Quiet
        Write-Log "Resolved installer URL       : $($asset.browser_download_url)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version  = $version
            Url      = $asset.browser_download_url
            FileName = $asset.name
        }
    }
    catch {
        Write-Log "Failed to resolve Calibrite PROFILER release: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCalibrite {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Calibrite PROFILER (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Resolve-CalibriteRelease
    if (-not $release) { throw "Could not resolve Calibrite PROFILER release." }

    $version           = $release.Version
    $installerFileName = $release.FileName
    $downloadUrl       = $release.Url

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-CalibritePayloadIsExecutable -Path $localExe

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
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand $UninstallerPath -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : $DetectionFile"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Calibrite PROFILER"
        Publisher       = "Calibrite LLC"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("calibrite PROFILER")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = $DetectionFile
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageCalibrite {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Calibrite PROFILER (x64) - PACKAGE phase"
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
        $r = Resolve-CalibriteRelease -Quiet
        if (-not $r) { exit 1 }
        Write-Output $r.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Calibrite PROFILER GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Calibrite PROFILER (x64) Auto-Packager starting"
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
        Invoke-StageCalibrite
    }
    elseif ($PackageOnly) {
        Invoke-PackageCalibrite
    }
    else {
        Invoke-StageCalibrite
        Invoke-PackageCalibrite
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-calibrite'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
