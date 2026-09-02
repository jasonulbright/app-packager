<#
Vendor: Opera Software
App: Opera Browser
CMName: Opera Browser
VendorUrl: https://www.opera.com/
CPE: cpe:2.3:a:opera:opera:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://blogs.opera.com/desktop/
DownloadPageUrl: https://www.opera.com/download
UpdateCadenceDays: 21

.SYNOPSIS
    Packages Opera Browser (x64) for MECM.

.DESCRIPTION
    Resolves the latest Opera desktop release from the vendor's public release
    directory listing, downloads the x64 NSIS setup, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    based detection on opera.exe.

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
    Content is staged under: <FileServerPath>\Applications\Opera Software\Opera Browser\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Opera).
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
    Outputs only the latest available Opera version string and exits.

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
$ReleaseListingUrl = "https://get.geo.opera.com/pub/opera/desktop/"

$VendorFolder = "Opera Software"
$AppFolder    = "Opera Browser"

$BaseDownloadRoot = Join-Path $DownloadRoot "Opera"
$InstallArgsLine   = "/silent /allusers=1 /launchopera=0 /setdefaultbrowser=0"
$UninstallArgsLine = "--uninstall --system-level --force-uninstall --silent"
$DetectionPath     = "C:\Program Files\Opera"

# --- Functions ---


function Get-LatestOperaVersion {
    param([switch]$Quiet)

    Write-Log "Opera release listing URL    : $ReleaseListingUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $ReleaseListingUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Opera release listing: $ReleaseListingUrl" }

        $rx = [regex]'href="(?<ver>\d+\.\d+\.\d+\.\d+)/"'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any version folders in the release listing."
        }

        # The listing is sorted as text, so 99.x sorts after 135.x; compare as
        # versions to pick the real newest build.
        $version = ($rxMatches | ForEach-Object { [version]$_.Groups['ver'].Value } |
                    Sort-Object -Descending | Select-Object -First 1).ToString()

        Write-Log "Latest Opera version         : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Opera version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageOpera {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Opera Browser (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestOperaVersion
    if (-not $version) { throw "Could not resolve Opera version." }

    $installerFileName = "Opera_${version}_Setup_x64.exe"
    $downloadUrl = "$ReleaseListingUrl$version/win/$installerFileName"

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
    $installWrapper = (
        ('$installer = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        'if (-not (Test-Path -LiteralPath $installer)) { Write-Error "Missing Opera installer"; exit 2 }',
        '$proc = Start-Process -FilePath $installer -ArgumentList @(''/silent'', ''/allusers=1'', ''/launchopera=0'', ''/setdefaultbrowser=0'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Opera keeps its uninstaller in a per-build subfolder under the install
    # path, so the path cannot be hard-coded; locate installer.exe at runtime.
    $uninstallWrapper = (
        '$installer = Get-ChildItem -LiteralPath ''C:\Program Files\Opera'' -Recurse -Filter ''installer.exe'' -ErrorAction SilentlyContinue | Select-Object -First 1',
        'if (-not $installer) { exit 0 }',
        '$proc = Start-Process -FilePath $installer.FullName -ArgumentList @(''--uninstall'', ''--system-level'', ''--force-uninstall'', ''--silent'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installWrapper `
        -UninstallPs1Content $uninstallWrapper

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $DetectionPath"
    Write-Log "Detection file               : opera.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName          = "Opera Browser"
        Publisher        = "Opera Software"
        SoftwareVersion  = $version
        DisplayName      = "Opera Browser"
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = $InstallArgsLine
        UninstallArgs    = $UninstallArgsLine
        UninstallCommand = "$DetectionPath\installer.exe"
        RunningProcess   = @("opera")
        Detection        = @{
            Type          = "File"
            FilePath      = $DetectionPath
            FileName      = "opera.exe"
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

function Invoke-PackageOpera {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Opera Browser (x64) - PACKAGE phase"
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
        $v = Get-LatestOperaVersion -Quiet
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
    Write-Log "Opera Browser (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseListingUrl            : $ReleaseListingUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageOpera
    }
    elseif ($PackageOnly) {
        Invoke-PackageOpera
    }
    else {
        Invoke-StageOpera
        Invoke-PackageOpera
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-opera'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
