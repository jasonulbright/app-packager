<#
Vendor: Piriform Software Ltd.
App: CCleaner
CMName: CCleaner
VendorUrl: https://www.ccleaner.com/
CPE: cpe:2.3:a:piriform:ccleaner:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.ccleaner.com/ccleaner/version-history
DownloadPageUrl: https://www.ccleaner.com/ccleaner/download
UpdateCadenceDays: 30

.SYNOPSIS
    Packages CCleaner Free for MECM.

.DESCRIPTION
    Reads the current version from the vendor version-history page, downloads
    the matching full Free installer from the vendor distribution endpoint,
    stages content to a versioned local folder, and creates an MECM Application
    with ARP registry detection.

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
    Content is staged under: <FileServerPath>\Applications\Piriform\CCleaner\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\CCleaner).
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
    Outputs only the latest available CCleaner version string and exits.

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
$VersionHistoryUrl = "https://www.ccleaner.com/ccleaner/version-history"
$DownloadUrl       = "https://bits.avcdn.net/productfamily_CCLEANER7/insttype_FREE/platform_WIN/installertype_FULL/build_RELEASE"

$VendorFolder = "Piriform"
$AppFolder    = "CCleaner"

$BaseDownloadRoot = Join-Path $DownloadRoot "CCleaner"

$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CCleaner"

# --- Functions ---


function Get-LatestCCleanerVersion {
    <#
    .SYNOPSIS
        Returns the highest version listed on the vendor version-history page.
    .DESCRIPTION
        The page lists every supported branch (7.x alongside 6.x), so the
        highest [version] wins rather than the first match in document order.
    #>
    param([switch]$Quiet)

    Write-Log "Version history URL          : $VersionHistoryUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $VersionHistoryUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch CCleaner version history: $VersionHistoryUrl" }

        $rx = [regex]'v(?<ver>\d+\.\d+\.\d+)'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any version strings on the version history page."
        }

        $best = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        if ([string]::IsNullOrWhiteSpace($best)) { throw "Version selection produced an empty result." }

        Write-Log "Latest CCleaner version      : $best" -Quiet:$Quiet
        return $best
    }
    catch {
        Write-Log "Failed to get CCleaner version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCCleaner {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CCleaner - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestCCleanerVersion
    if (-not $version) { throw "Could not resolve CCleaner version." }

    $installerFileName = "CCleaner-$version-Setup.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $DownloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # The distribution endpoint answers 200 with an HTML error body rather
    # than a redirect when a build is unavailable.
    $header = Get-Content -LiteralPath $localExe -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($header.Count -lt 2 -or $header[0] -ne 0x4D -or $header[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $localExe"
    }

    $binaryVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if (-not [string]::IsNullOrWhiteSpace($binaryVersion)) {
        Write-Log "Installer FileVersion        : $binaryVersion"
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
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand 'unused'

    # The uninstaller path is not fixed across CCleaner builds; the ARP
    # UninstallString is the only value guaranteed to point at the copy that
    # installed this client.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CCleaner',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CCleaner'
)
$cmd = $null
foreach ($k in $keys) {
    $p = Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue
    if ($p -and $p.UninstallString) { $cmd = $p.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "CCleaner"
    $publisher = "Piriform Software Ltd."

    Write-Log ""
    Write-Log "Detection key                : $ArpRegistryKey"
    Write-Log "Detection value              : DisplayVersion >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("CCleaner", "CCleaner64")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
            DisplayName         = $appName
            DisplayVersion      = $version
            Is64Bit             = $true
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

function Invoke-PackageCCleaner {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CCleaner - PACKAGE phase"
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

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
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
        $v = Get-LatestCCleanerVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("CCleaner GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CCleaner Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionHistoryUrl            : $VersionHistoryUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCCleaner
    }
    elseif ($PackageOnly) {
        Invoke-PackageCCleaner
    }
    else {
        Invoke-StageCCleaner
        Invoke-PackageCCleaner
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-ccleaner'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
