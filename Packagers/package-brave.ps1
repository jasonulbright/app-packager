<#
Vendor: Brave Software
App: Brave Browser
CMName: Brave Browser
VendorUrl: https://brave.com/
CPE: cpe:2.3:a:brave:brave:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/brave/brave-browser/releases
DownloadPageUrl: https://brave.com/download/
IconSource: Installer
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Brave Browser (x64) for ConfigMgr.

.DESCRIPTION
    Resolves the latest Brave release from the brave-browser GitHub releases API,
    downloads the raw x64 installer from the vendor endpoint, stages content to a
    versioned local folder, and creates a ConfigMgr Application with file-version
    based detection on brave.exe.

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
    Content is staged under: <FileServerPath>\Applications\Brave Software\Brave Browser\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Brave).
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
    create ConfigMgr application with file-based detection.

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
# The GitHub release carries only the Omaha wrappers, which refuse a
# system-level install; the raw installer that honors --system-level is
# published unversioned at this endpoint, so the downloaded file's own
# version is the packaged version.
$DownloadUrl   = "https://referrals.brave.com/latest/brave_installer-x64.exe"

$VendorFolder = "Brave Software"
$AppFolder    = "Brave Browser"

$BaseDownloadRoot = Join-Path $DownloadRoot "Brave"
$InstallArgsLine   = "--install --silent --system-level --do-not-launch-chrome"
$UninstallArgsLine = "--uninstall --system-level --force-uninstall"
$InstallRoot       = "C:\Program Files\BraveSoftware\Brave-Browser\Application"
$DetectionPath     = $InstallRoot

# --- Functions ---


function Get-LatestBraveVersion {
    param([switch]$Quiet)

    Write-Log "Brave release API URL        : $ReleaseApiUrl" -Quiet:$Quiet

    try {
        # The GitHub API rejects requests without a User-Agent header.
        $jsonText = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" @(Get-GitHubApiCurlArgs) $ReleaseApiUrl) -join ''
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

    # --- Download ---
    # The endpoint serves the current release only, so the cached copy is
    # refreshed when the server reports a newer file and the version comes
    # from the binary itself.
    $localInstaller = Join-Path $BaseDownloadRoot "brave_installer-x64.exe"
    Write-Log "Download URL                 : $DownloadUrl"
    Invoke-CachedDownload -Url $DownloadUrl -OutFile $localInstaller

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localInstaller).FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    # The binary reports <Chromium major>.<Brave version>; brave.exe carries
    # the same value, so the detector compares against the full string while
    # the package is versioned the way the vendor numbers releases.
    if ($fileVersion -notmatch '^\d+\.(\d+\.\d+\.\d+)$') {
        throw "Installer file version '$fileVersion' is not in the expected <chromium>.<brave> form."
    }
    $version = $Matches[1]
    $latest = Get-LatestBraveVersion -Quiet
    if ($latest -and $latest -ne $version) {
        Write-Log ("Downloaded installer is {0}; the release feed lists {1}. Packaging the downloaded build." -f $version, $latest) -Level WARN
    }

    $installerFileName = "brave_installer-x64-$version.exe"
    Write-Log "Version                      : $version"
    Write-Log "Installer file version       : $fileVersion"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

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
    # Without --do-not-launch-chrome the installer starts the browser under
    # the installing account. Uninstall runs the setup the product keeps
    # under its versioned Installer folder; 19 is its success code and 20
    # asks for a reboot.
    $installWrapper = (
        ('$installer = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        'if (-not (Test-Path -LiteralPath $installer)) { Write-Error "Missing Brave installer"; exit 2 }',
        '$proc = Start-Process -FilePath $installer -ArgumentList @(''--install'', ''--silent'', ''--system-level'', ''--do-not-launch-chrome'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallWrapper = (
        ('$setup = Get-ChildItem -Path ''{0}\*\Installer\setup.exe'' -ErrorAction SilentlyContinue | Sort-Object -Property FullName -Descending | Select-Object -First 1' -f $InstallRoot),
        'if (-not $setup) { exit 0 }',
        '$proc = Start-Process -FilePath $setup.FullName -ArgumentList @(''--uninstall'', ''--system-level'', ''--force-uninstall'') -Wait -PassThru -NoNewWindow',
        'if ($proc.ExitCode -eq 19) { exit 0 }',
        'if ($proc.ExitCode -eq 20) { exit 3010 }',
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
        UninstallCommand = ("{0}\{1}\Installer\setup.exe" -f $InstallRoot, $fileVersion)
        RunningProcess   = @("brave")
        Detection        = @{
            Type          = "File"
            FilePath      = $DetectionPath
            FileName      = "brave.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $fileVersion
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
