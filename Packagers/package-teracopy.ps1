<#
Vendor: Code Sector
App: TeraCopy
CMName: TeraCopy
VendorUrl: https://www.codesector.com/
CPE: cpe:2.3:a:codesector:teracopy:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://blog.codesector.com/category/teracopy/
DownloadPageUrl: https://www.codesector.com/downloads
UpdateCadenceDays: 180

.SYNOPSIS
    Packages TeraCopy (x64) for MECM.

.DESCRIPTION
    Resolves the newest build from the vendor's update feed
    (codesector.com/updates/teracopy.txt), which carries the product version,
    a versioned download URL and the payload SHA256. Downloads the installer,
    verifies its hash, stages content to a versioned local folder, and creates
    an MECM Application with script-based detection on the ARP entry.

    The payload is an Advanced Installer bootstrapper wrapping an MSI, so the
    silent switches are /exenoui /qn and ALLUSERS=1 forces the per-machine
    install of a package that also supports a per-user install.

    The free edition is licensed for home use only; business and organizational
    use requires a purchased TeraCopy Pro license. A license entitlement must be
    in place before this payload is deployed in a business environment.

    Supports two-phase operation:
      -StageOnly    Download, verify hash, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Code Sector\TeraCopy\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TeraCopy).
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
    create MECM application with script-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available TeraCopy version string and exits.

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
$UpdateFeedUrl = "https://codesector.com/updates/teracopy.txt"

$VendorFolder = "Code Sector"
$AppFolder    = "TeraCopy"

$BaseDownloadRoot = Join-Path $DownloadRoot "TeraCopy"

$InstallerFileName = "teracopy-setup.exe"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The CDN answers a missing build path with a 200 HTML page, which would
        otherwise stage as a valid-looking EXE and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestTeraCopyRelease {
    <#
    .SYNOPSIS
        Returns the newest TeraCopy version, download URL and payload SHA256.
    .DESCRIPTION
        The update feed is the same INI-shaped file the shipped updater reads,
        so it carries the exact build the vendor is serving plus its hash. The
        download page links an unversioned filename that resolves to the same
        build; the feed is preferred because it makes the payload verifiable.
    #>
    param([switch]$Quiet)

    Write-Log "TeraCopy update feed         : $UpdateFeedUrl" -Quiet:$Quiet

    try {
        $feed = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $UpdateFeedUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch TeraCopy update feed: $UpdateFeedUrl" }

        $version = [regex]::Match($feed, '(?m)^\s*ProductVersion\s*=\s*(?<v>\d+(\.\d+)+)\s*$').Groups['v'].Value
        $url     = [regex]::Match($feed, '(?m)^\s*URL\s*=\s*(?<u>\S+)\s*$').Groups['u'].Value
        $sha256  = [regex]::Match($feed, '(?m)^\s*SHA256\s*=\s*(?<h>[0-9A-Fa-f]{64})\s*$').Groups['h'].Value

        if ([string]::IsNullOrWhiteSpace($version)) { throw "Update feed did not carry a ProductVersion." }
        if ([string]::IsNullOrWhiteSpace($url))     { throw "Update feed did not carry a download URL." }
        if ($url -notmatch '\.exe($|\?)')           { throw "Update feed download URL is not an EXE: $url" }

        Write-Log "Latest TeraCopy version      : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $InstallerFileName
            DownloadUrl = $url
            Sha256      = $sha256
        }
    }
    catch {
        Write-Log "Failed to get TeraCopy version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeraCopy {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeraCopy (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestTeraCopyRelease
    if (-not $releaseInfo) { throw "Could not resolve TeraCopy version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Download ---
    # The staged filename is fixed, so a cached copy from an earlier release
    # would shadow the new one; the payload is re-fetched every stage.
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe

    Assert-PayloadIsExecutable -Path $localExe

    if ($releaseInfo.Sha256) {
        $actual = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $releaseInfo.Sha256.ToUpperInvariant()) {
            throw "Downloaded payload SHA256 '$actual' does not match the vendor feed value '$($releaseInfo.Sha256)'."
        }
        Write-Log "Payload SHA256 verified      : $actual"
    }

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -ne $version) {
        throw "Downloaded installer reports version '$fileVersion' but the update feed announced '$version'; refusing to stage a mismatched payload."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # The bootstrapper registers the ARP entry under the wrapped MSI's
    # ProductCode, which changes between builds, so uninstall resolves the
    # entry by DisplayName and hands the ProductCode back to msiexec.
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/exenoui', '/qn', 'ALLUSERS=1') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallScript = @'
$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    foreach ($sub in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
        $entry = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction SilentlyContinue
        if (-not $entry -or [string]$entry.DisplayName -notlike 'TeraCopy*') { continue }
        $code = [regex]::Match([string]$entry.UninstallString, '\{[0-9A-Fa-f-]{36}\}').Value
        if (-not $code) { $code = [regex]::Match($sub.PSChildName, '\{[0-9A-Fa-f-]{36}\}').Value }
        if (-not $code) { continue }
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList @('/x', $code, '/qn', '/norestart') -Wait -PassThru -NoNewWindow
        exit $proc.ExitCode
    }
}
Write-Error 'TeraCopy uninstall entry not found.'
exit 1
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # The ARP key name is the wrapped MSI's ProductCode and changes between
    # builds, so detection matches on DisplayName and compares DisplayVersion.
    $detectionScript = @"
`$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
`$wanted = [version]'$version'
foreach (`$root in `$roots) {
    if (-not (Test-Path -LiteralPath `$root)) { continue }
    foreach (`$sub in Get-ChildItem -LiteralPath `$root -ErrorAction SilentlyContinue) {
        `$entry = Get-ItemProperty -LiteralPath `$sub.PSPath -ErrorAction SilentlyContinue
        if (-not `$entry -or [string]`$entry.DisplayName -notlike 'TeraCopy*') { continue }
        `$found = `$null
        if (-not [version]::TryParse([string]`$entry.DisplayVersion, [ref]`$found)) { continue }
        if (`$found -ge `$wanted) { Write-Output 'Installed'; exit 0 }
    }
}
exit 0
"@

    Write-Log ""
    Write-Log "Detection                    : ARP DisplayName 'TeraCopy*' with DisplayVersion >= $version"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "TeraCopy"
        Publisher       = "Code Sector"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/exenoui /qn ALLUSERS=1"
        UninstallArgs   = ""
        RunningProcess  = @("TeraCopy")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
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

function Invoke-PackageTeraCopy {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeraCopy (x64) - PACKAGE phase"
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
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
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
        $info = Get-LatestTeraCopyRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("TeraCopy GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeraCopy (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "UpdateFeedUrl                : $UpdateFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTeraCopy
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeraCopy
    }
    else {
        Invoke-StageTeraCopy
        Invoke-PackageTeraCopy
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-teracopy'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
