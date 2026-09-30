<#
Vendor: Slido
App: Slido for Windows
CMName: Slido for Windows
VendorUrl: https://www.slido.com/powerpoint-polling
ReleaseNotesUrl: https://community.slido.com/powerpoint-244/slido-for-powerpoint-on-windows-changelog-2503
DownloadPageUrl: https://www.slido.com/powerpoint-polling
IconSource: None

.SYNOPSIS
    Packages Slido for Windows (Slido for PowerPoint, admin MSI, x64) for ConfigMgr.

.DESCRIPTION
    Reads the latest version of the Slido admin installer from the Slido
    package endpoint, with the vendor download redirect as the fallback,
    downloads the per-machine admin MSI, stages content to a versioned local
    folder, and creates a ConfigMgr Application with file version detection.

    The admin MSI installs the PowerPoint add-in for all users; one x64
    package serves 32-bit and 64-bit Office. The install sets
    DISABLE_UPDATE_CHECKS=1, so the in-app updater does not change the
    version outside deployments. Per-user "Basic" installs (SlidoSetup EXE)
    are a separate product that this package does not update.

    Detection requires Slido.exe in the install folder at or above the
    packaged version; its file version carries the full build number.

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
    Content is staged under: <FileServerPath>\Applications\Slido\Slido for Windows\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Slido).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Slido for Windows version string and exits.

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
$LatestApiUrl = "https://api.slido.com/eu1/api/v0.5/switcher/packages/powerpoint-win-admin64/latest"
$RedirectUrl  = "https://www.slido.com/api/download?application=powerpoint-win-admin64"

$VendorFolder = "Slido"
$AppFolder    = "Slido for Windows"

$BaseDownloadRoot = Join-Path $DownloadRoot "Slido"

# The x64 admin MSI installs to ProgramFilesFolder, the 32-bit folder.
$InstallFolder = "C:\Program Files (x86)\Slido\Slido for Windows"

# --- Functions ---


function Get-LatestSlidoRelease {
    <#
    .SYNOPSIS
        Returns the latest admin MSI as a PSCustomObject with Version (four
        parts), FileName, and DownloadUrl.
    #>
    param([switch]$Quiet)

    $version = $null
    $downloadUrl = $null

    Write-Log "Slido package endpoint       : $LatestApiUrl" -Quiet:$Quiet
    try {
        $json = (curl.exe -L --fail --silent --show-error $LatestApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query $LatestApiUrl" }
        $package = ConvertFrom-Json $json
        $version = [string]$package.version
        $downloadUrl = [string]$package.publicUrl
    }
    catch {
        Write-Log "Package endpoint failed: $($_.Exception.Message). Reading the download redirect." -Level WARN -Quiet:$Quiet
    }

    if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$' -or $downloadUrl -notmatch '\.msi$') {
        # The redirect target carries "<timestamp>_<four-part version>/SlidoAdmin_x64_v<x.y.z>.msi".
        $location = (curl.exe --silent --show-error -o NUL -w "%{redirect_url}" $RedirectUrl) -join ''
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($location)) { throw "Failed to resolve the Slido download redirect: $RedirectUrl" }
        $m = [regex]::Match($location, '_(\d+\.\d+\.\d+\.\d+)/[^/]+\.msi$')
        if (-not $m.Success) { throw "The Slido download redirect names no versioned MSI: $location" }
        $version = $m.Groups[1].Value
        $downloadUrl = $location
    }

    Write-Log "Latest Slido version         : $version" -Quiet:$Quiet

    return [PSCustomObject]@{
        Version     = $version
        FileName    = [System.IO.Path]::GetFileName(([uri]$downloadUrl).AbsolutePath)
        DownloadUrl = $downloadUrl
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSlido {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Slido for Windows (admin, x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-LatestSlidoRelease
    $version     = $release.Version
    $msiFileName = $release.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local installer path         : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading Slido admin MSI..."
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi
    $productCode    = $props["ProductCode"]
    $productVersion = $props["ProductVersion"]
    if ([string]::IsNullOrWhiteSpace($productCode)) { throw "MSI ProductCode missing." }
    Write-Log "MSI ProductName              : $($props['ProductName'])"
    Write-Log "MSI ProductVersion           : $productVersion"
    Write-Log "MSI ProductCode              : $productCode"
    $shortVersion = ($version -split '\.')[0..2] -join '.'
    if ($productVersion -ne $shortVersion) {
        throw "The MSI reports version $productVersion, but the release is $version."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName -ExtraInstallArgs @('DISABLE_UPDATE_CHECKS=1')
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $appName   = "Slido for Windows $version"
    $publisher = "Slido"

    Write-Log ""
    Write-Log "Detection                    : $InstallFolder\Slido.exe >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart DISABLE_UPDATE_CHECKS=1"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("Slido", "POWERPNT")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallFolder
            FileName      = "Slido.exe"
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

function Invoke-PackageSlido {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Slido for Windows (admin, x64) - PACKAGE phase"
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
    Write-Log "Detection File               : $($manifest.Detection.FilePath)\$($manifest.Detection.FileName)"
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
        $rel = Get-LatestSlidoRelease -Quiet
        Write-Output $rel.Version
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
    Write-Log "Slido for Windows Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LatestApiUrl                 : $LatestApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSlido
    }
    elseif ($PackageOnly) {
        Invoke-PackageSlido
    }
    else {
        Invoke-StageSlido
        Invoke-PackageSlido
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-slido'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
