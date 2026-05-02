<#
Vendor: Microsoft
App: Microsoft Edge
CMName: Microsoft Edge
VendorUrl: https://www.microsoft.com/edge/business/download
CPE: cpe:2.3:a:microsoft:edge_chromium:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel
DownloadPageUrl: https://www.microsoft.com/en-us/edge/business/download

.SYNOPSIS
    Packages Microsoft Edge MSI for MECM.

.DESCRIPTION
    Downloads the latest Microsoft Edge MSI from the official
    Microsoft link, stages content to a versioned local folder with compound
    file-based detection metadata, and creates an MECM Application with
    file-version detection (OR connector).
    Detection uses msedge.exe version in the primary install path OR the staged
    EdgeUpdate path (OR connector).

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\Chredge\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Edge).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download MSI, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with compound file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Edge version string and exits.

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
$EdgeStableMSIUrl = "https://go.microsoft.com/fwlink/?LinkID=2093437"
$EdgeVersionUrl   = "https://edgeupdates.microsoft.com/api/products?view=enterprise"

$VendorFolder = "Microsoft"
$AppFolder    = "Chredge"
$MsiFileName  = "MicrosoftEdgeEnterpriseX64.msi"

$BaseDownloadRoot = Join-Path $DownloadRoot "Edge"

# --- Functions ---


function Get-LatestEdgeVersion {
    param([switch]$Quiet)

    Write-Log "Edge version API             : $EdgeVersionUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $EdgeVersionUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Edge version info: $EdgeVersionUrl" }

        $response = ConvertFrom-Json $json
        $stableChannel = $response | Where-Object { $_.Product -eq "Stable" }
        $latestVersion = $stableChannel.Releases |
            Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
            Sort-Object { [version]$_.ProductVersion } -Descending |
            Select-Object -First 1 -ExpandProperty ProductVersion

        if (-not $latestVersion) {
            throw "Could not determine latest Microsoft Edge Stable version."
        }

        Write-Log "Latest Edge version          : $latestVersion" -Quiet:$Quiet
        return $latestVersion
    }
    catch {
        Write-Log "Failed to get Edge version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageEdge {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Edge - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestEdgeVersion
    if (-not $version) { throw "Could not resolve Edge version." }

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $MsiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $EdgeStableMSIUrl"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $EdgeStableMSIUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrappers = New-MsiWrapperContent -MsiFileName $MsiFileName

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    $detectionPath1 = "C:\Program Files (x86)\Microsoft\Edge\Application"
    $detectionPath2 = "C:\Program Files (x86)\Microsoft\EdgeUpdate\Install"

    $appName   = "Microsoft Edge - $version"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection (OR) path 1        : $detectionPath1\msedge.exe"
    Write-Log "Detection (OR) path 2        : $detectionPath2\msedge.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        RunningProcess  = @("msedge")
        Detection       = @{
            Type      = "Compound"
            Connector = "Or"
            Clauses   = @(
                @{
                    Type          = "File"
                    FilePath      = $detectionPath1
                    FileName      = "msedge.exe"
                    PropertyType  = "Version"
                    Operator      = "GreaterEquals"
                    ExpectedValue = $version
                    Is64Bit       = $false
                },
                @{
                    Type          = "File"
                    FilePath      = $detectionPath2
                    FileName      = "msedge.exe"
                    PropertyType  = "Version"
                    Operator      = "GreaterEquals"
                    ExpectedValue = $version
                    Is64Bit       = $false
                }
            )
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

function Invoke-PackageEdge {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Edge - PACKAGE phase"
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
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
    Write-Log "Detection Connector          : $($manifest.Detection.Connector)"
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
        $v = Get-LatestEdgeVersion -Quiet
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
    Write-Log "Microsoft Edge Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "EdgeVersionUrl               : $EdgeVersionUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageEdge
    }
    elseif ($PackageOnly) {
        Invoke-PackageEdge
    }
    else {
        Invoke-StageEdge
        Invoke-PackageEdge
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
