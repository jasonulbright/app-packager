<#
Vendor: PostgreSQL Global Development Group
App: PostgreSQL 17 (x64)
CMName: PostgreSQL 17
VendorUrl: https://www.postgresql.org/
CPE: cpe:2.3:a:postgresql:postgresql:17.*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.postgresql.org/docs/17/release.html
DownloadPageUrl: https://www.postgresql.org/download/windows/

.SYNOPSIS
    Packages PostgreSQL 17 (x64) for MECM.

.DESCRIPTION
    Downloads the latest PostgreSQL 17.x EDB installer from the EnterpriseDB CDN,
    stages content to a versioned local folder with file-based version detection
    metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The EDB installer is a BitRock InstallBuilder EXE (not MSI). Default install
    creates the PostgreSQL service, data directory, and superuser account.

    IMPORTANT: The install wrapper includes a default superuser password placeholder
    (P0stgres!MECM). Edit install.ps1 in the staged content folder before deploying
    to production, or change the password post-install.

    GetLatestVersionOnly queries the endoflife.date API for the latest 17.x release.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\PostgreSQL\PostgreSQL 17\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\PostgreSQL17).
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
    Queries the endoflife.date API for the latest PostgreSQL 17.x version,
    outputs the version string, and exits. No MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$MajorVersion          = 17
$VersionApiUrl         = "https://endoflife.date/api/postgresql.json"
$CdnBaseUrl            = "https://get.enterprisedb.com/postgresql"
$InstallerFileNamePattern = "postgresql-{0}-1-windows-x64.exe"

$VendorFolder = "PostgreSQL"
$AppFolder    = "PostgreSQL $MajorVersion"

$BaseDownloadRoot = Join-Path $DownloadRoot "PostgreSQL$MajorVersion"

# --- Functions ---


function Get-LatestPostgreSQLVersion {
    param([switch]$Quiet)

    Write-Log "Version API URL              : $VersionApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $VersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query endoflife.date API." }

        $cycles = ConvertFrom-Json $json
        $entry = $cycles | Where-Object { $_.cycle -eq "$MajorVersion" } | Select-Object -First 1
        if (-not $entry) { throw "No release found for PostgreSQL $MajorVersion." }

        $version = $entry.latest
        if ([string]::IsNullOrWhiteSpace($version)) { throw "Empty version in API response." }

        Write-Log "Latest PostgreSQL $MajorVersion version  : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get PostgreSQL version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePostgreSQL {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PostgreSQL $MajorVersion (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestPostgreSQLVersion
    if (-not $version) { throw "Could not determine latest PostgreSQL $MajorVersion version." }

    # --- Download ---
    $installerFileName = $InstallerFileNamePattern -f $version
    $downloadUrl = "$CdnBaseUrl/$installerFileName"
    $localExe = Join-Path $BaseDownloadRoot $installerFileName

    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Local installer path         : $localExe"
    Write-Log ""

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading PostgreSQL $MajorVersion installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
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
    # BitRock InstallBuilder -- uses --mode unattended (not MSI)
    # IMPORTANT: Change the superpassword below before deploying to production.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'--mode', 'unattended', '--unattendedmodeui', 'none', '--superpassword', 'P0stgres!MECM', '--serverport', '5432', '--install_runtimes', '0', '--disable-components', 'stackbuilder'" `
        -UninstallCommand "C:\Program Files\PostgreSQL\$MajorVersion\uninstall-postgresql.exe" `
        -UninstallArgs "'--mode', 'unattended'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # EnterpriseDB installer writes ARP DisplayVersion as "<ver>-1" (build
    # suffix). Use BeginsWith on the short version so we're agnostic to the
    # trailing suffix. File-based Version detection on postgres.exe was
    # observed to fail intermittently in MECM even though FileVersion
    # metadata reads correctly - registry ARP is more reliable.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PostgreSQL $MajorVersion"

    $appName   = "PostgreSQL $MajorVersion $version"
    $publisher = "PostgreSQL Global Development Group"

    Write-Log "ARP Registry Key             : $arpRegistryKey"
    Write-Log "ARP DisplayVersion (starts)  : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "--mode unattended --unattendedmodeui none --superpassword P0stgres!MECM --serverport 5432 --install_runtimes 0 --disable-components stackbuilder"
        UninstallArgs   = "--mode unattended"
        RunningProcess  = @("postgres", "pgAdmin4")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            Operator            = "BeginsWith"
            ExpectedValue       = $version
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

function Invoke-PackagePostgreSQL {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PostgreSQL $MajorVersion (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

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

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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
        $v = Get-LatestPostgreSQLVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PostgreSQL $MajorVersion (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "Version API URL              : $VersionApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePostgreSQL
    }
    elseif ($PackageOnly) {
        Invoke-PackagePostgreSQL
    }
    else {
        Invoke-StagePostgreSQL
        Invoke-PackagePostgreSQL
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
