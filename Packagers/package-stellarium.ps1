<#
Vendor: Stellarium
App: Stellarium
CMName: Stellarium
VendorUrl: https://stellarium.org/
CPE: cpe:2.3:a:stellarium:stellarium:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/Stellarium/stellarium/releases
DownloadPageUrl: https://stellarium.org/
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Stellarium (x64, Qt6) for MECM.

.DESCRIPTION
    Resolves the newest release from the Stellarium/stellarium GitHub releases
    API, downloads the Qt6 win64 Inno Setup installer, stages content to a
    versioned local folder with ARP detection metadata, and creates an MECM
    Application with registry-based detection.

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
    Content is staged under: <FileServerPath>\Applications\Stellarium\Stellarium\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Stellarium).
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
    Outputs only the latest available Stellarium version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/Stellarium/stellarium/releases/latest"

$VendorFolder = "Stellarium"
$AppFolder    = "Stellarium"

$BaseDownloadRoot = Join-Path $DownloadRoot "Stellarium"

# The installer is Inno Setup with a fixed AppId, so the ARP key name and the
# per-machine install directory are stable across releases; /ALLUSERS selects
# the per-machine scope the installer otherwise asks for.
$InstallPath    = "{0}\Stellarium" -f $env:ProgramFiles
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Stellarium_is1"

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
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestStellariumRelease {
    <#
    .SYNOPSIS
        Returns the newest Stellarium version and its Qt6 win64 installer URL.
    .DESCRIPTION
        A release carries Qt5 and Qt6 Windows builds plus an ARM64 build; the
        asset is matched on the qt6-win64 suffix so neither the legacy Qt5 build
        nor the ARM64 build can be staged as the x64 payload.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the GitHub releases API." }

        $release = ConvertFrom-Json $json
        $version = $release.tag_name -replace '^v', ''
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse a version from the GitHub release tag."
        }

        $asset = $release.assets |
            Where-Object { $_.name -match '^stellarium-.+-qt6-win64\.exe$' } |
            Select-Object -First 1
        if (-not $asset) { throw "No qt6-win64 installer asset found in the latest release." }

        Write-Log "Latest Stellarium version    : $version" -Quiet:$Quiet
        Write-Log "Installer asset              : $($asset.name)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get Stellarium version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageStellarium {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Stellarium (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestStellariumRelease
    if (-not $releaseInfo) { throw "Could not resolve Stellarium version." }

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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/ALLUSERS','/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART'" `
        -UninstallCommand ("{0}\unins000.exe" -f $InstallPath) `
        -UninstallArgs "'/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # The installer writes DisplayVersion under a fixed ARP key name rather than
    # a ProductCode GUID; a version comparison keeps detection true after an
    # in-place upgrade that lands a higher build.
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log "Uninstall command            : $InstallPath\unins000.exe /VERYSILENT"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Stellarium"
        Publisher       = "Stellarium team"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/ALLUSERS /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("stellarium")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
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

function Invoke-PackageStellarium {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Stellarium (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

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
        $info = Get-LatestStellariumRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Stellarium GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Stellarium (x64) Auto-Packager starting"
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
        Invoke-StageStellarium
    }
    elseif ($PackageOnly) {
        Invoke-PackageStellarium
    }
    else {
        Invoke-StageStellarium
        Invoke-PackageStellarium
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-stellarium'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
