<#
Vendor: Unity Technologies
App: Unity Hub
CMName: Unity Hub
VendorUrl: https://unity.com/
CPE: cpe:2.3:a:unity:unity_hub:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://unity.com/unity-hub/release-notes
DownloadPageUrl: https://unity.com/download
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Unity Hub (x64) for MECM.

.DESCRIPTION
    Resolves the current Unity Hub release from the vendor update feed at
    public-cdn.cloud.unity3d.com, downloads the machine-wide NSIS installer,
    stages content to a versioned local folder, and creates an MECM Application
    with file-version-based detection on the installed Unity Hub.exe.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    Unity Hub is a launcher: Unity Editor versions are downloaded separately by
    the signed-in user and are subject to Unity's own licensing terms.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Unity Technologies\Unity Hub\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\UnityHub).
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
    Outputs only the latest available Unity Hub version string and exits.

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
$HubCdnRoot   = "https://public-cdn.cloud.unity3d.com/hub/prod/"
$LatestYmlUrl = $HubCdnRoot + "latest.yml"

$VendorFolder = "Unity Technologies"
$AppFolder    = "Unity Hub"

$BaseDownloadRoot = Join-Path $DownloadRoot "UnityHub"

$InstallDir = "{0}\Unity Hub" -f $env:ProgramFiles

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A CDN edge that answers 200 with an HTML error body would otherwise
        stage as a valid-looking EXE and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestUnityHubRelease {
    <#
    .SYNOPSIS
        Returns the current Unity Hub version and its x64 installer URL.
    .DESCRIPTION
        The feed is an electron-updater YAML document whose 'path' entry is
        relative to the feed's own directory. It is parsed line-wise rather
        than with a YAML reader so the packager keeps no extra dependency;
        both 'version' and 'path' must be present or the release is rejected.
    #>
    param([switch]$Quiet)

    Write-Log "Unity Hub update feed        : $LatestYmlUrl" -Quiet:$Quiet

    try {
        $lines = curl.exe -L --fail --silent --show-error $LatestYmlUrl
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Unity Hub update feed: $LatestYmlUrl" }

        $text = ($lines -join "`n")

        $verMatch = [regex]::Match($text, '(?m)^version:\s*(?<ver>\d+(\.\d+)+)\s*$')
        if (-not $verMatch.Success) { throw "Could not parse a version from the Unity Hub update feed." }
        $version = $verMatch.Groups['ver'].Value

        $pathMatch = [regex]::Match($text, '(?m)^path:\s*(?<path>\S+\.exe)\s*$')
        if (-not $pathMatch.Success) { throw "Could not parse an installer path from the Unity Hub update feed." }
        $relPath = $pathMatch.Groups['path'].Value

        $fileName = Split-Path -Leaf $relPath
        if ($fileName -notmatch 'x64\.exe$') {
            throw "Unity Hub update feed does not point at an x64 installer: $relPath"
        }

        Write-Log "Latest Unity Hub version     : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $HubCdnRoot + $relPath
            FileName    = $fileName
        }
    }
    catch {
        Write-Log "Failed to get Unity Hub version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageUnityHub {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Unity Hub (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestUnityHubRelease
    if (-not $releaseInfo) { throw "Could not resolve Unity Hub version." }

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
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = (Get-Item -LiteralPath $localExe).VersionInfo.FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    if ($fileVersion -and $fileVersion -ne $version) {
        throw "Installer file version ($fileVersion) does not match the feed version ($version)."
    }

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
    $uninstallExe = Join-Path $InstallDir "Uninstall Unity Hub.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand $uninstallExe -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    # The electron-builder payload registers its ARP entry under a generated
    # key name, so detection reads the installed binary's file version instead
    # of an uninstall key whose name is not stated by the vendor.
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : Unity Hub.exe"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Unity Hub"
        Publisher        = "Unity Technologies"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/S"
        UninstallArgs    = "/S"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("Unity Hub")
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "Unity Hub.exe"
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

function Invoke-PackageUnityHub {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Unity Hub (x64) - PACKAGE phase"
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
        $info = Get-LatestUnityHubRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Unity Hub GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Unity Hub (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LatestYmlUrl                 : $LatestYmlUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageUnityHub
    }
    elseif ($PackageOnly) {
        Invoke-PackageUnityHub
    }
    else {
        Invoke-StageUnityHub
        Invoke-PackageUnityHub
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-unityhub'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
