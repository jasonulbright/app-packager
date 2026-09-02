<#
Vendor: Karakun AG
App: OpenWebStart
CMName: OpenWebStart
VendorUrl: https://openwebstart.com/
CPE: cpe:2.3:a:karakun:openwebstart:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/karakun/OpenWebStart/releases
DownloadPageUrl: https://openwebstart.com/download/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages OpenWebStart (x64) for MECM.

.DESCRIPTION
    Downloads the latest OpenWebStart Windows x64 installer from the official
    GitHub releases API, stages content together with an install4j response
    file that pins an all-users install, and creates an MECM Application with
    file-version-based detection on the installed javaws.exe.

    Supports two-phase operation:
      -StageOnly    Download, generate response file and content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Karakun AG\OpenWebStart\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\OpenWebStart).
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
    Outputs only the latest available OpenWebStart version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/karakun/OpenWebStart/releases/latest"

$VendorFolder = "Karakun AG"
$AppFolder    = "OpenWebStart"

$BaseDownloadRoot = Join-Path $DownloadRoot "OpenWebStart"

$InstallDir       = "{0}\OpenWebStart" -f $env:ProgramFiles
$ResponseFileName = "ows-response.varfile"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A release-asset redirect that lands on an error page answers 200 with
        HTML, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestOpenWebStartRelease {
    <#
    .SYNOPSIS
        Returns the newest OpenWebStart release version and its x64 installer URL.
    .DESCRIPTION
        Asset names carry the version with underscores rather than dots
        (OpenWebStart_windows-x64_1_14_0.exe), so the version comes from the tag
        and the asset is matched on shape.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub API URL               : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error -A "PowerShell" $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch GitHub release info: $GitHubApiUrl" }

        $release = ConvertFrom-Json $json

        $version = $release.tag_name -replace '^[vV]', ''
        if ($version -notmatch '^\d+(\.\d+)+$') {
            throw "Unrecognized OpenWebStart release tag: $($release.tag_name)"
        }

        $asset = $release.assets |
            Where-Object { $_.name -match '^OpenWebStart_windows-x64_.*\.exe$' } |
            Select-Object -First 1

        if (-not $asset) {
            throw "Could not find a Windows x64 installer asset in the latest GitHub release."
        }

        Write-Log "Latest OpenWebStart version  : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $asset.browser_download_url
            FileName    = $asset.name
        }
    }
    catch {
        Write-Log "Failed to get OpenWebStart version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function New-OpenWebStartResponseFile {
    <#
    .SYNOPSIS
        Writes the install4j response file that pins an all-users install.
    .DESCRIPTION
        Unattended install4j runs otherwise fall back to the interactive
        defaults, which install per user under the local profile. The target
        directory is not set here: the -dir argument on the command line takes
        precedence over the response file.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $lines = @(
        'sys.adminRights$Boolean=true'
        'userMode$Integer=1'
    )
    Set-Content -LiteralPath $Path -Value $lines -Encoding ASCII -ErrorAction Stop
    Write-Log "Wrote response file          : $Path"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageOpenWebStart {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "OpenWebStart (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestOpenWebStartRelease
    if (-not $releaseInfo) { throw "Could not resolve OpenWebStart version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', 'PowerShell')
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

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

    New-OpenWebStartResponseFile -Path (Join-Path $localContentPath $ResponseFileName)

    # --- Generate content wrappers ---
    # install4j: -q is unattended, -dir pins the target so the uninstaller path
    # stays predictable across upgrades, and the response file forces the
    # all-users mode that a silent run would otherwise decide on its own.
    $uninstallExe = Join-Path $InstallDir "uninstall.exe"
    $installArgs = "'-q', '-varfile', (Join-Path `$PSScriptRoot '$ResponseFileName'), '-dir', '$InstallDir'"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand $uninstallExe -UninstallArgs "'-q'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    # install4j names its uninstall key after a generated installer id, so
    # detection reads the installed launcher's file version instead.
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : javaws.exe"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "OpenWebStart"
        Publisher        = "Karakun AG"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "-q -varfile $ResponseFileName -dir `"$InstallDir`""
        UninstallCommand = $uninstallExe
        UninstallArgs    = "-q"
        RunningProcess   = @("javaws", "itw-settings")
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "javaws.exe"
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

function Invoke-PackageOpenWebStart {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "OpenWebStart (x64) - PACKAGE phase"
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
        $info = Get-LatestOpenWebStartRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("OpenWebStart GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "OpenWebStart (x64) Auto-Packager starting"
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
        Invoke-StageOpenWebStart
    }
    elseif ($PackageOnly) {
        Invoke-PackageOpenWebStart
    }
    else {
        Invoke-StageOpenWebStart
        Invoke-PackageOpenWebStart
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-openwebstart'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
