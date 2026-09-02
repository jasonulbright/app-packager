<#
Vendor: Rainmeter
App: Rainmeter
CMName: Rainmeter
VendorUrl: https://www.rainmeter.net/
CPE: cpe:2.3:a:rainmeter:rainmeter:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/rainmeter/rainmeter/releases
DownloadPageUrl: https://www.rainmeter.net/
UpdateCadenceDays: 180

.SYNOPSIS
    Packages Rainmeter (x64) for MECM.

.DESCRIPTION
    Resolves the latest Rainmeter installer from the GitHub releases API, stages
    content to a versioned local folder, and creates an MECM Application with
    file-existence detection.

    The vendor ships a single installer that carries both architectures and
    selects one through /VERSION.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
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
    Outputs only the latest available Rainmeter version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$GitHubApiUrl = "https://api.github.com/repos/rainmeter/rainmeter/releases/latest"

$VendorFolder = "Rainmeter"
$AppFolder    = "Rainmeter"

$BaseDownloadRoot = Join-Path $DownloadRoot "Rainmeter"
$InstallPath      = "C:\Program Files\Rainmeter"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        A mirror that answers 200 with an HTML error body would otherwise stage
        as a valid-looking installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2) { throw "Downloaded payload is too small to be an executable: $Path" }
    if ($bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not an executable (no MZ header): $Path"
    }
}


function Get-LatestRainmeterRelease {
    <#
    .SYNOPSIS
        Returns the latest Rainmeter version and its installer URL.
    .DESCRIPTION
        The release tag carries a build revision (v4.5.26.3894) that the
        installer file name and the product itself do not, so the version is
        taken from the asset name.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        # The API rejects requests without a user agent.
        $json = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" $GitHubApiUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        $asset = $release.assets |
            Where-Object { $_.name -match '^Rainmeter-(?<ver>\d+(\.\d+)+)\.exe$' } |
            Select-Object -First 1
        if (-not $asset) { throw "No Rainmeter installer asset found in the latest release." }

        $version = [regex]::Match($asset.name, '^Rainmeter-(?<ver>\d+(\.\d+)+)\.exe$').Groups['ver'].Value

        Write-Log "Latest Rainmeter version     : $version" -Quiet:$Quiet
        return @{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get Rainmeter version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageRainmeter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Rainmeter (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestRainmeterRelease
    if (-not $releaseInfo) { throw "Could not resolve Rainmeter version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName
    $downloadUrl       = $releaseInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Rainmeter..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

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
    # /AUTOSTARTUP=0 keeps the installer out of the deploying account's Startup
    # folder; /RESTART=0 stops it relaunching a running instance under SYSTEM.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/S', '/VERSION=64', '/AUTOSTARTUP=0', '/RESTART=0'" `
        -UninstallCommand "$InstallPath\uninstall.exe" `
        -UninstallArgs "'/S'"

    # The skin host runs from the tray and holds its own files; the uninstaller
    # leaves them behind while it runs. Absent uninstaller means the product is
    # not present, so removal exits clean.
    $customUninstall = (
        'Get-Process Rainmeter -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        'Start-Sleep -Seconds 2',
        ('$uninstaller = ''{0}\uninstall.exe''' -f $InstallPath),
        'if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : Rainmeter.exe"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Rainmeter"
        Publisher       = "Rainmeter"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S /VERSION=64 /AUTOSTARTUP=0 /RESTART=0"
        UninstallArgs   = "/S"
        RunningProcess  = @("Rainmeter")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallPath
            FileName     = "Rainmeter.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
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

function Invoke-PackageRainmeter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Rainmeter (x64) - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $info = Get-LatestRainmeterRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Rainmeter GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Rainmeter (x64) Auto-Packager starting"
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
        Invoke-StageRainmeter
    }
    elseif ($PackageOnly) {
        Invoke-PackageRainmeter
    }
    else {
        Invoke-StageRainmeter
        Invoke-PackageRainmeter
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-rainmeter'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
