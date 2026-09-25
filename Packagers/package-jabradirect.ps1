<#
Vendor: GN Audio A/S
App: Jabra Direct
CMName: Jabra Direct
VendorUrl: https://www.jabra.com/software-and-services/jabra-direct
CPE: cpe:2.3:a:jabra:direct:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.jabra.com/support/release-notes/release-note-jabra-direct
DownloadPageUrl: https://www.jabra.com/software-and-services/jabra-direct
IconSource: Installer
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Jabra Direct for ConfigMgr.

.DESCRIPTION
    Reads the current version from the vendor release-notes page, downloads
    JabraDirectSetup.exe from the vendor's fixed download URL, stages content
    to a versioned local folder, and creates a ConfigMgr Application with
    registry-based detection on the WiX Burn bundle ARP entry.

    The download URL carries no version, so the staged payload's file version
    is checked against the release-notes version before content is accepted.

    JabraDirectSetup.exe is a WiX Burn bundle: /install /quiet /norestart
    installs silently for all users. Uninstall runs the cached bundle named by
    the ARP entry with /uninstall /quiet /norestart.

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
    Content is staged under: <FileServerPath>\Applications\GN Audio\Jabra Direct\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\JabraDirect).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Jabra Direct version string and exits.

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
    [int]$MaximumRuntimeMins = 45,
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
$ReleaseNotesUrl = "https://www.jabra.com/support/release-notes/release-note-jabra-direct"
$DownloadUrl     = "https://jabraxpressonlineprdstor.blob.core.windows.net/jdo/JabraDirectSetup.exe"

$VendorFolder = "GN Audio"
$AppFolder    = "Jabra Direct"

$BaseDownloadRoot = Join-Path $DownloadRoot "JabraDirect"

$InstallerFileName = "JabraDirectSetup.exe"

# Burn registers the bundle under its BundleId. The bundle stub is 32-bit, so
# the entry lands in the 32-bit view of the machine uninstall hive.
$BundleId       = "{FF8111EB-E2A2-4A6B-830C-4F676D20FB39}"
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$BundleId"

# --- Functions ---


function Get-JabraDirectVersionFromPage {
    <#
    .SYNOPSIS
        Returns the newest release version printed on the release-notes page,
        or nothing when the page carries no version marker.
    .DESCRIPTION
        The page lists releases newest first, each with a release-version
        element; the first one is the build the download URL serves.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html)

    $m = [regex]::Match($Html, 'data-testid="release-version"[^>]*>\s*(?<ver>\d+\.\d+\.\d+)\s*<')
    if (-not $m.Success) { return $null }
    return $m.Groups['ver'].Value
}


function Get-LatestJabraDirectVersion {
    param([switch]$Quiet)

    Write-Log "Release notes URL            : $ReleaseNotesUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the release-notes page: $ReleaseNotesUrl" }

        $version = Get-JabraDirectVersionFromPage -Html $html
        if (-not $version) { throw "Release-notes page carries no release-version marker." }

        Write-Log "Latest Jabra Direct version  : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Jabra Direct version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Assert-PayloadVersion {
    <#
    .SYNOPSIS
        Throws when the payload's file version differs from the expected one.
    .DESCRIPTION
        The download URL is unversioned; a release-notes page updated before
        the blob, or a cached older payload, would otherwise stage content
        whose recorded version does not match the installed DisplayVersion.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Expected
    )

    $fileVersion = [string](Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.FileVersion
    $fileVersion = $fileVersion.Trim()
    if ($fileVersion -ne $Expected) {
        throw "Payload file version '$fileVersion' does not match the release-notes version '$Expected'."
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageJabraDirect {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Jabra Direct - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestJabraDirectVersion
    if (-not $version) { throw "Could not resolve Jabra Direct version." }

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $DownloadUrl"
    Write-Log ""

    # --- Download ---
    # The payload name carries no version, so the local copy is versioned to
    # keep an older cached build from shadowing the new one.
    $localExe = Join-Path $BaseDownloadRoot ("JabraDirectSetup-{0}.exe" -f $version)
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', 'Mozilla/5.0')
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe
    Assert-PayloadVersion -Path $localExe -Expected $version
    Assert-ArpDetectionKey -InstallerPath $localExe -ExpectedKey $ArpRegistryKey -Is64BitView $false

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
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs "'/install', '/quiet', '/norestart'" `
        -UninstallCommand 'unused'

    # The bundle caches itself under a per-build folder, so the ARP
    # QuietUninstallString is the only value that names the right copy.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\__BUNDLEID__',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\__BUNDLEID__'
)
$hit = $null
foreach ($key in $keys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $hit = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($hit) { break }
}
if (-not $hit) { exit 0 }
$cmd = if ($hit.QuietUninstallString) { $hit.QuietUninstallString } else { $hit.UninstallString }
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = ($cmd -split '\s+/')[0].Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/uninstall', '/quiet', '/norestart') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@
    $uninstallContent = $uninstallContent.Replace('__BUNDLEID__', $BundleId)

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Jabra Direct"
        Publisher       = "GN Audio A/S"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/install /quiet /norestart"
        UninstallArgs   = "/uninstall /quiet /norestart"
        RunningProcess  = @("jabra-direct")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
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

function Invoke-PackageJabraDirect {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Jabra Direct - PACKAGE phase"
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
        $version = Get-LatestJabraDirectVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Jabra Direct GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Jabra Direct Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadUrl                  : $DownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageJabraDirect
    }
    elseif ($PackageOnly) {
        Invoke-PackageJabraDirect
    }
    else {
        Invoke-StageJabraDirect
        Invoke-PackageJabraDirect
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-jabradirect'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
