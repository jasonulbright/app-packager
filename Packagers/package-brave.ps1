<#
Vendor: Brave Software
App: Brave Browser
CMName: Brave Browser
VendorUrl: https://brave.com/
CPE: cpe:2.3:a:brave:brave:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/brave/brave-browser/releases
DownloadPageUrl: https://brave.com/download/
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Brave Browser (x64) for MECM.

.DESCRIPTION
    Resolves the latest Brave release from the brave-browser GitHub releases API,
    downloads the standalone silent system-level installer, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    based detection on brave.exe.

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
    Content is staged under: <FileServerPath>\Applications\Brave Software\Brave Browser\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Brave).
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
    Outputs only the latest available Brave version string and exits.

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
$ReleaseApiUrl = "https://api.github.com/repos/brave/brave-browser/releases/latest"
$DownloadBase  = "https://github.com/brave/brave-browser/releases/download"

$VendorFolder = "Brave Software"
$AppFolder    = "Brave Browser"

$BaseDownloadRoot = Join-Path $DownloadRoot "Brave"
$InstallArgsLine   = "--install --silent --system-level"
$UninstallArgsLine = "--uninstall --silent --system-level"
$DetectionPath     = "C:\Program Files\BraveSoftware\Brave-Browser\Application"

# --- Functions ---


function Get-LatestBraveVersion {
    param([switch]$Quiet)

    Write-Log "Brave release API URL        : $ReleaseApiUrl" -Quiet:$Quiet

    try {
        # The GitHub API rejects requests without a User-Agent header.
        $jsonText = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" $ReleaseApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Brave release API: $ReleaseApiUrl" }

        $json = ConvertFrom-Json $jsonText
        $tag = [string]$json.tag_name
        if ([string]::IsNullOrWhiteSpace($tag)) { throw "tag_name field was empty." }

        $version = $tag -replace '^v', ''
        if ($version -notmatch '^\d+(\.\d+)+$') {
            throw "Release tag '$tag' did not yield a numeric version."
        }

        Write-Log "Latest Brave version         : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Brave version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBrave {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Brave Browser (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestBraveVersion
    if (-not $version) { throw "Could not resolve Brave version." }

    $installerFileName = "BraveBrowserStandaloneSilentSetup-$version.exe"
    $downloadUrl = "$DownloadBase/v$version/BraveBrowserStandaloneSilentSetup.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localInstaller = Join-Path $BaseDownloadRoot $installerFileName
    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedInstaller = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedInstaller)) {
        Copy-Item -LiteralPath $localInstaller -Destination $stagedInstaller -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedInstaller"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # The standalone setup handles both directions; uninstall runs the staged
    # copy from the deployment type content folder rather than a fixed path,
    # because the installed product keeps its setup under a versioned
    # Installer directory that moves with every update.
    $installWrapper = (
        ('$installer = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        'if (-not (Test-Path -LiteralPath $installer)) { Write-Error "Missing Brave installer"; exit 2 }',
        '$proc = Start-Process -FilePath $installer -ArgumentList @(''--install'', ''--silent'', ''--system-level'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallWrapper = (
        ('$installer = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        'if (-not (Test-Path -LiteralPath $installer)) { exit 0 }',
        '$proc = Start-Process -FilePath $installer -ArgumentList @(''--uninstall'', ''--silent'', ''--system-level'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installWrapper `
        -UninstallPs1Content $uninstallWrapper

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $DetectionPath"
    Write-Log "Detection file               : brave.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName          = "Brave Browser"
        Publisher        = "Brave Software"
        SoftwareVersion  = $version
        DisplayName      = "Brave Browser"
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = $InstallArgsLine
        UninstallArgs    = $UninstallArgsLine
        UninstallCommand = $installerFileName
        RunningProcess   = @("brave")
        Detection        = @{
            Type          = "File"
            FilePath      = $DetectionPath
            FileName      = "brave.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
        }
    }
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageBrave {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Brave Browser (x64) - PACKAGE phase"
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
        $v = Get-LatestBraveVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Brave Browser (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseApiUrl                : $ReleaseApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageBrave
    }
    elseif ($PackageOnly) {
        Invoke-PackageBrave
    }
    else {
        Invoke-StageBrave
        Invoke-PackageBrave
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-brave'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
