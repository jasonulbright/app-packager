<#
Vendor: Bruno Software Inc
App: Bruno
CMName: Bruno
VendorUrl: https://www.usebruno.com/
CPE: cpe:2.3:a:usebruno:bruno:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/usebruno/bruno/releases
DownloadPageUrl: https://www.usebruno.com/downloads
SecurityAdvisoryUrl: https://github.com/usebruno/bruno/security/advisories
IconSource: Installer
UpdateCadenceDays: 14
WsusSupport: Yes
ArpName: Bruno
ArpArch: x64

.SYNOPSIS
    Packages Bruno (open source API client) for ConfigMgr.

.DESCRIPTION
    Downloads the latest Bruno x64 MSI from GitHub releases, verifies it
    against the SHA-256 digest that GitHub publishes for the release asset,
    stages content to a versioned local folder with ARP detection metadata
    read from the MSI, and creates a ConfigMgr Application with
    registry-based detection. The MSI installs per machine to
    C:\Program Files\Bruno.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 10

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 25

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Queries the GitHub releases API for the latest Bruno version, outputs the
    version string, and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
    [int]$EstimatedRuntimeMins = 10,
    [int]$MaximumRuntimeMins = 25,
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
$GitHubApiUrl = "https://api.github.com/repos/usebruno/bruno/releases/latest"

$VendorFolder = "Bruno Software"
$AppFolder    = "Bruno"

$BaseDownloadRoot = Join-Path $DownloadRoot "Bruno"

# --- Functions ---


function Get-LatestBrunoRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error --max-time 60 @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        $version = $release.tag_name -replace '^v', ''
        if ($version -notmatch '^\d+(\.\d+){1,3}$') {
            throw "Could not parse a version from the GitHub release tag '$($release.tag_name)'."
        }

        # bruno_<version>_x64_win.msi. The arm64 build has its own file name.
        $asset = $release.assets | Where-Object { $_.name -match '_x64_win\.msi$' } | Select-Object -First 1
        if (-not $asset) { throw "No x64 MSI asset found in release." }

        $sha256 = $null
        if ($asset.digest -match '^sha256:(?<h>[0-9A-Fa-f]{64})$') { $sha256 = $Matches['h'].ToUpperInvariant() }

        Write-Log "Latest Bruno version       : $version" -Quiet:$Quiet
        return @{ Version = $version; FileName = $asset.name; DownloadUrl = $asset.browser_download_url; Sha256 = $sha256 }
    }
    catch {
        Write-Log "Failed to get Bruno version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBruno {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bruno - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestBrunoRelease
    if (-not $releaseInfo) { throw "Could not resolve Bruno version." }

    $version      = $releaseInfo.Version
    $MsiFileName  = $releaseInfo.FileName
    $downloadUrl  = $releaseInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "MSI filename                 : $MsiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    # A cached file counts only when it matches the release digest. A mismatch
    # removes it and downloads again once.
    $digestKnown = -not [string]::IsNullOrWhiteSpace($releaseInfo.Sha256)
    if (Test-Path -LiteralPath $localMsi) {
        if ($digestKnown -and (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash -ne $releaseInfo.Sha256) {
            Write-Log "Local MSI does not match the release digest. Downloading it again." -Level WARN
            Remove-Item -LiteralPath $localMsi -Force -ErrorAction Stop
        }
        else {
            Write-Log "Local MSI exists. Skipping download."
        }
    }
    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading Bruno MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }

    # --- Verify against the digest GitHub publishes for the asset ---
    if (-not $digestKnown) {
        Write-Log "GitHub published no SHA-256 digest for $MsiFileName. The file is not verified." -Level WARN
    }
    else {
        $actual = (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $releaseInfo.Sha256) {
            Remove-Item -LiteralPath $localMsi -Force -ErrorAction SilentlyContinue
            throw "SHA256 mismatch for $MsiFileName (release: $($releaseInfo.Sha256), file: $actual). The download was removed."
        }
        Write-Log "SHA256 matches the release   : $actual"
    }

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productVersionRaw = $props["ProductVersion"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $($props['ProductName'])"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $($props['Manufacturer'])"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
    Write-Log "Copied MSI to staged folder  : $stagedMsi"

    # --- Derive ARP detection from MSI properties ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $props['Manufacturer']
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Anoop M D" }
    $appName = "Bruno $productVersionRaw"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("Bruno")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $productVersionRaw -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageBruno {
    param([string]$StagedContentPath)

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bruno - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    if ([string]::IsNullOrWhiteSpace($StagedContentPath)) {
        $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
        if (-not (Test-Path -LiteralPath $versionFile)) {
            throw "Version marker not found - run Stage phase first: $versionFile"
        }
        $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()
        $localContentPath = Join-Path $BaseDownloadRoot $version
    }
    else {
        $localContentPath = $StagedContentPath
    }
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
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
        $info = Get-LatestBrunoRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
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
    Write-Log "Bruno Auto-Packager starting"
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
        Invoke-StageBruno
    }
    elseif ($PackageOnly) {
        Invoke-PackageBruno
    }
    else {
        $stagedPath = Invoke-StageBruno | Select-Object -Last 1
        Invoke-PackageBruno -StagedContentPath ([string]$stagedPath)
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-bruno'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
