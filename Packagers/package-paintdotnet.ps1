<#
Vendor: dotPDN LLC
App: Paint.NET
CMName: Paint.NET
VendorUrl: https://www.getpaint.net/
CPE: cpe:2.3:a:dotpdn:paint.net:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.getpaint.net/roadmap.html
DownloadPageUrl: https://www.getpaint.net/download.html
IconSource: Installer
WsusSupport: Yes
ArpName: Paint.NET*

.SYNOPSIS
    Packages Paint.NET (x64) MSI for ConfigMgr.

.DESCRIPTION
    Downloads the latest Paint.NET x64 MSI ZIP from GitHub releases, extracts
    the MSI, stages content to a versioned local folder with ARP detection
    metadata, and creates a ConfigMgr Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download ZIP, extract MSI, derive ARP detection, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    The MSI install passes CHECKFORUPDATES=0 and DESKTOPSHORTCUT=0 for
    enterprise-appropriate defaults.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\dotPDN LLC\Paint.NET (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (for example, <DownloadRoot>\PaintDotNet).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download ZIP, extract MSI, derive ARP detection
    from MSI properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Queries the GitHub releases API for the latest Paint.NET version, outputs
    the version string, and exits. No download or ConfigMgr changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (ConfigurationManager PowerShell module available)
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
$GitHubApiUrl = "https://api.github.com/repos/paintdotnet/release/releases?per_page=30"

$VendorFolder = "dotPDN LLC"
$AppFolder    = "Paint.NET (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "PaintDotNet"

# --- Functions ---


function Get-PaintDotNetRelease {
    # The newest release that carries the x64 MSI zip. The asset list decides; the file name is never built from the version.
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $selected = Select-GitHubReleaseWithAsset -Releases @(ConvertFrom-Json $json) -AssetPattern '^paint\.net\..+\.winmsi\.x64\.zip$'
        if (-not $selected) { throw "No Paint.NET release carries an x64 MSI zip." }
        if ([string]::IsNullOrWhiteSpace($selected.Version)) { throw "Could not parse version from GitHub release." }
        if ($selected.SkippedNewer.Count -gt 0) {
            Write-Log ("Paint.NET {0} publishes no x64 MSI zip; packaging {1}." -f ($selected.SkippedNewer -join ', '), $selected.Version) -Level WARN -Quiet:$Quiet
        }

        Write-Log "Latest Paint.NET version     : $($selected.Version)" -Quiet:$Quiet
        return $selected
    }
    catch {
        Write-Log "Failed to get Paint.NET version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-LatestPaintDotNetVersion {
    param([switch]$Quiet)
    $selected = Get-PaintDotNetRelease -Quiet:$Quiet
    if (-not $selected) { return $null }
    return $selected.Version
}

# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePaintDotNet {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Paint.NET (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-PaintDotNetRelease
    if (-not $release) { throw "Could not resolve Paint.NET version." }
    $version = $release.Version

    # The name and the URL come from the asset of the release, not from the version.
    $zipFileName = $release.AssetName
    $downloadUrl = $release.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ZIP ---
    $localZip = Join-Path $BaseDownloadRoot $zipFileName
    Write-Log "Local ZIP path               : $localZip"

    if (-not (Test-Path -LiteralPath $localZip)) {
        Write-Log "Downloading Paint.NET ZIP..."
        Assert-DownloadAvailable -Url $downloadUrl -What 'The installer file'
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localZip
    }
    else {
        Write-Log "Local ZIP exists. Skipping download."
    }

    # --- Extract MSI from ZIP ---
    $extractDir = Join-Path $BaseDownloadRoot "_extracted"
    if (Test-Path -LiteralPath $extractDir) {
        Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive -LiteralPath $localZip -DestinationPath $extractDir -Force -ErrorAction Stop

    $msiFile = Get-ChildItem -Path $extractDir -Filter "*.msi" -Recurse -File | Select-Object -First 1
    if (-not $msiFile) { throw "No MSI found inside Paint.NET ZIP." }
    $MsiFileName = $msiFile.Name

    Write-Log "Extracted MSI                : $($msiFile.FullName)"
    Write-Log ""

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $msiFile.FullName

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $msiFile.FullName -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers (custom install args for Paint.NET) ---
    $installPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''/quiet'', ''/norestart'', ''CHECKFORUPDATES=0'', ''DESKTOPSHORTCUT=0'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $publisher = "dotPDN LLC"
    $appName = $productName

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/quiet /norestart CHECKFORUPDATES=0 DESKTOPSHORTCUT=0"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("paintdotnet")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackagePaintDotNet {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Paint.NET (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from staged MSI ---
    Initialize-Folder -Path $BaseDownloadRoot

    # Find MSI in extracted dir or versioned folders
    $extractDir = Join-Path $BaseDownloadRoot "_extracted"
    $msiFile = $null
    if (Test-Path -LiteralPath $extractDir) {
        $msiFile = Get-ChildItem -Path $extractDir -Filter "*.msi" -Recurse -File | Select-Object -First 1
    }
    if (-not $msiFile) {
        throw "No extracted MSI found - run Stage phase first."
    }

    $props = Get-MsiPropertyMap -MsiPath $msiFile.FullName
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $productVersion   = $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
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
        $v = Get-LatestPaintDotNetVersion -Quiet
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
    Write-Log "Paint.NET (x64) Auto-Packager starting"
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
        Invoke-StagePaintDotNet
    }
    elseif ($PackageOnly) {
        Invoke-PackagePaintDotNet
    }
    else {
        Invoke-StagePaintDotNet
        Invoke-PackagePaintDotNet
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-paintdotnet'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
