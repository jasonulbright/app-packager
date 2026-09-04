<#
Vendor: Bambu Lab
App: Bambu Studio
CMName: Bambu Studio
VendorUrl: https://bambulab.com/en/download/studio
CPE: cpe:2.3:a:bambulab:bambu_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/bambulab/BambuStudio/releases
DownloadPageUrl: https://bambulab.com/en/download/studio
IconSource: Installer
UpdateCadenceDays: 45

.SYNOPSIS
    Packages Bambu Studio (x64) for MECM.

.DESCRIPTION
    Downloads the latest Bambu Studio Windows installer from the official
    GitHub releases API, stages content to a versioned local folder, and
    creates an MECM Application with file-version-based detection on the
    installed bambu-studio.exe.

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
    Content is staged under: <FileServerPath>\Applications\Bambu Lab\Bambu Studio\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\BambuStudio).
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
    Outputs only the latest available Bambu Studio version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/bambulab/BambuStudio/releases/latest"

$VendorFolder = "Bambu Lab"
$AppFolder    = "Bambu Studio"

$BaseDownloadRoot = Join-Path $DownloadRoot "BambuStudio"

$InstallDir = "{0}\Bambu Studio" -f $env:ProgramFiles

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


function ConvertTo-BambuStudioVersion {
    <#
    .SYNOPSIS
        Normalizes a Bambu Studio release tag to a comparable version string.
    .DESCRIPTION
        Tags are zero-padded per segment (v02.08.02.61). The zero padding makes
        string comparison order releases wrongly and does not match the file
        version detection compares against, so each segment is reduced to its
        numeric value.
    #>
    param([Parameter(Mandatory)][string]$Tag)

    $raw = $Tag -replace '^v', ''
    $parts = $raw -split '\.'
    if ($parts.Count -lt 2) { throw "Unrecognized Bambu Studio release tag: $Tag" }

    $numeric = foreach ($p in $parts) {
        if ($p -notmatch '^\d+$') { throw "Unrecognized Bambu Studio release tag: $Tag" }
        [string][int]$p
    }

    return ($numeric -join '.')
}


function Get-LatestBambuStudioRelease {
    param([switch]$Quiet)

    Write-Log "GitHub API URL               : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error -A "PowerShell" @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch GitHub release info: $GitHubApiUrl" }

        $release = ConvertFrom-Json $json
        $version = ConvertTo-BambuStudioVersion -Tag $release.tag_name

        # Asset names carry a build timestamp and have changed prefix across
        # releases, so the Windows installer is matched on shape, not on a
        # literal name.
        $asset = $release.assets |
            Where-Object { $_.name -match '^Bambu_Studio_win.*\.exe$' } |
            Select-Object -First 1

        if (-not $asset) {
            throw "Could not find a Windows installer asset in the latest GitHub release."
        }

        Write-Log "Latest Bambu Studio version  : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $asset.browser_download_url
            FileName    = $asset.name
        }
    }
    catch {
        Write-Log "Failed to get Bambu Studio version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBambuStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bambu Studio (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestBambuStudioRelease
    if (-not $releaseInfo) { throw "Could not resolve Bambu Studio version." }

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

    # --- Generate content wrappers ---
    $uninstallExe = Join-Path $InstallDir "Uninstall.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand $uninstallExe -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    # The installer writes its ARP entry through the 32-bit registry view of a
    # 64-bit payload, so detection reads the installed binary's file version
    # instead of guessing which uninstall hive the entry lands in.
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : bambu-studio.exe"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Bambu Studio"
        Publisher       = "Bambu Lab"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        UninstallCommand = $uninstallExe
        RunningProcess  = @("bambu-studio")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "bambu-studio.exe"
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

function Invoke-PackageBambuStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bambu Studio (x64) - PACKAGE phase"
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
        $info = Get-LatestBambuStudioRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Bambu Studio GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bambu Studio (x64) Auto-Packager starting"
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
        Invoke-StageBambuStudio
    }
    elseif ($PackageOnly) {
        Invoke-PackageBambuStudio
    }
    else {
        Invoke-StageBambuStudio
        Invoke-PackageBambuStudio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-bambustudio'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
