<#
Vendor: Kai Kramer
App: KeyStore Explorer
CMName: KeyStore Explorer
VendorUrl: https://keystore-explorer.org/
CPE: cpe:2.3:a:keystore-explorer:keystore_explorer:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/kaikramer/keystore-explorer/releases
DownloadPageUrl: https://keystore-explorer.org/downloads.html
IconSource: Installer
UpdateCadenceDays: 180

.SYNOPSIS
    Packages KeyStore Explorer (x64) for MECM.

.DESCRIPTION
    Resolves the newest release from the GitHub releases API, downloads the
    bundled-JRE Windows installer, stages content to a versioned local folder,
    and creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers and stage manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Kai Kramer\KeyStore Explorer\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\KeyStoreExplorer).
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
    Outputs only the latest available KeyStore Explorer version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/kaikramer/keystore-explorer/releases/latest"

$VendorFolder = "Kai Kramer"
$AppFolder    = "KeyStore Explorer"

$BaseDownloadRoot = Join-Path $DownloadRoot "KeyStoreExplorer"
$InstallPath      = "{0}\KeyStore Explorer" -f $env:ProgramFiles

# The installer is an Inno Setup package with a fixed AppId; the ARP key is
# that GUID with the _is1 suffix Inno appends.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A771FBEB-F7E1-4443-9181-AFD57F7BFF45}_is1"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A release-asset redirect that lands on an HTML error page would
        otherwise stage as a valid-looking installer and fail at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestKeyStoreExplorerRelease {
    <#
    .SYNOPSIS
        Returns the newest KeyStore Explorer version and its installer URL.
    .DESCRIPTION
        Asset names strip the dots from the version (5.7.0 becomes kse-570), so
        the version is taken from the release tag and the asset is matched by
        name. The plain setup asset is chosen over the -no-jre variant because
        the product will not start without a JRE the package does not otherwise
        guarantee.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        # The API rejects requests without a user agent.
        $json = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json

        $asset = $release.assets |
            Where-Object { $_.name -match '^kse-\d+-setup\.exe$' } |
            Select-Object -First 1
        if (-not $asset) { throw "Could not locate a Windows setup asset in the latest release." }

        $version = $release.tag_name -replace '^v', ''
        if ($version -notmatch '^\d+(\.\d+)+$') {
            throw "Could not parse a version from release tag '$($release.tag_name)'."
        }

        Write-Log "Latest KeyStore Explorer ver : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get KeyStore Explorer version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageKeyStoreExplorer {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "KeyStore Explorer (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestKeyStoreExplorerRelease
    if (-not $releaseInfo) { throw "Could not resolve KeyStore Explorer version." }

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
        Write-Log "Downloading KeyStore Explorer..."
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
    # UninstallFilesDir puts the generated uninstaller in an "uninstall"
    # subfolder rather than the application root.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/SP-'" `
        -UninstallCommand "$InstallPath\uninstall\unins000.exe" `
        -UninstallArgs "'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/SP-'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "KeyStore Explorer"
        Publisher       = "Kai Kramer"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-"
        UninstallArgs   = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP-"
        RunningProcess  = @("kse")
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

function Invoke-PackageKeyStoreExplorer {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "KeyStore Explorer (x64) - PACKAGE phase"
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
        $info = Get-LatestKeyStoreExplorerRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("KeyStore Explorer GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "KeyStore Explorer (x64) Auto-Packager starting"
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
        Invoke-StageKeyStoreExplorer
    }
    elseif ($PackageOnly) {
        Invoke-PackageKeyStoreExplorer
    }
    else {
        Invoke-StageKeyStoreExplorer
        Invoke-PackageKeyStoreExplorer
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-keystoreexplorer'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
