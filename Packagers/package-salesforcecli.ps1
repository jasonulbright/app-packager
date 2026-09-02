<#
Vendor: Salesforce
App: Salesforce CLI
CMName: Salesforce CLI
VendorUrl: https://developer.salesforce.com/tools/salesforcecli
ReleaseNotesUrl: https://github.com/salesforcecli/cli/releases
DownloadPageUrl: https://developer.salesforce.com/tools/salesforcecli
IconSource: Installer
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Salesforce CLI (sf, x64) for MECM.

.DESCRIPTION
    Reads the version from the vendor's stable-channel build manifest, downloads
    the matching x64 installer from the same channel path, stages content to a
    versioned local folder, and creates an MECM Application with registry-based
    detection on the installer's ARP entry.

    The installer adds <InstallDir>\bin to the PATH of the account that runs it,
    so a machine-context deployment leaves the sf command off interactive users'
    PATH until it is added by another mechanism.

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
    Content is staged under: <FileServerPath>\Applications\Salesforce\Salesforce CLI\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\SalesforceCLI).
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
    Outputs only the latest available Salesforce CLI version string and exits.

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
$ChannelRootUrl   = "https://developer.salesforce.com/media/salesforce-cli/sf/channels/stable"
$BuildManifestUrl = "$ChannelRootUrl/sf-win32-x64-buildmanifest"
$InstallerUrl     = "$ChannelRootUrl/sf-x64.exe"

$VendorFolder = "Salesforce"
$AppFolder    = "Salesforce CLI"

$BaseDownloadRoot  = Join-Path $DownloadRoot "SalesforceCLI"
$InstallerFileName = "sf-x64.exe"

$InstallDir = "{0}\sf" -f $env:ProgramFiles

# The installer is a 32-bit NSIS executable that writes its ARP entry without
# switching the registry view, so the values land under WOW6432Node.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\sf"

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


function Get-LatestSalesforceCliVersion {
    param([switch]$Quiet)

    Write-Log "Build manifest URL           : $BuildManifestUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $BuildManifestUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the stable-channel build manifest." }

        $manifest = ConvertFrom-Json $json
        $version = [string]$manifest.version

        if ([string]::IsNullOrWhiteSpace($version) -or $version -notmatch '^\d+\.\d+\.\d+') {
            throw "Build manifest did not carry a usable version value."
        }

        Write-Log "Latest Salesforce CLI version: $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Salesforce CLI version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSalesforceCli {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Salesforce CLI (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestSalesforceCliVersion
    if (-not $version) { throw "Could not resolve Salesforce CLI version." }

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Download ---
    # The channel URL always serves the current stable build, so a cached file
    # from an earlier version would stage under the new version folder.
    $localExe = Join-Path $BaseDownloadRoot ("sf-x64-{0}.exe" -f $version)
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $InstallerUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

    # The installer binary is stamped with the same version the channel
    # manifest reports; a mismatch means the channel moved mid-download.
    $stampedVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ($stampedVersion) {
        $stampedVersion = $stampedVersion.Trim()
        if ($stampedVersion -notlike "$version*") {
            throw "Installer file version ($stampedVersion) does not match the manifest version ($version)."
        }
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $uninstallExe = Join-Path $InstallDir "uninstall.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand $uninstallExe `
        -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log "Install directory            : $InstallDir"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Salesforce CLI"
        Publisher        = "Salesforce"
        SoftwareVersion  = $version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/S"
        UninstallCommand = $uninstallExe
        UninstallArgs    = "/S"
        RunningProcess   = @("sf", "node")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $false
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

function Invoke-PackageSalesforceCli {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Salesforce CLI (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $version = Get-LatestSalesforceCliVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Salesforce CLI GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Salesforce CLI (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ChannelRootUrl               : $ChannelRootUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSalesforceCli
    }
    elseif ($PackageOnly) {
        Invoke-PackageSalesforceCli
    }
    else {
        Invoke-StageSalesforceCli
        Invoke-PackageSalesforceCli
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-salesforcecli'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
