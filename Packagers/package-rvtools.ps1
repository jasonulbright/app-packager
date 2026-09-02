<#
Vendor: Dell
App: RVTools
CMName: RVTools
VendorUrl: https://www.robware.net/
CPE: cpe:2.3:a:robware:rvtools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.dell.com/support/kbdoc/en-us/000325532
DownloadPageUrl: https://www.dell.com/support/kbdoc/en-us/000325532
UpdateCadenceDays: 120

.SYNOPSIS
    Packages RVTools MSI for MECM.

.DESCRIPTION
    Downloads the latest RVTools MSI from the Dell download host, stages
    content to a versioned local folder with ARP detection metadata, and
    creates an MECM Application with registry-based detection.

    The download host serves the MSI at a fixed, predictable path but does not
    publish a directory index, and the vendor knowledge-base article that names
    the current version is not machine-readable. Version discovery therefore
    walks the download host itself from a known-good floor version upward until
    the next candidate stops resolving.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available RVTools version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - Target clients need .NET Framework 4.7.2 or later; the MSI enforces this
      as a launch condition and fails the install when it is absent.
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
$DownloadUrlFormat = "https://downloads.dell.com/rvtools/rvtools{0}.msi"

# Floor for the version walk: the oldest release the walk is allowed to settle
# on. Raising it as releases age keeps the probe count down; lowering it below
# a version the host has retired makes discovery fail outright.
$FloorVersion   = "4.8.1"
$MaxVersionHops = 12

$VendorFolder = "Dell"
$AppFolder    = "RVTools"

$BaseDownloadRoot = Join-Path $DownloadRoot "RVTools"

# The package is a dual-purpose installer whose default install scope follows
# the invoking user; ALLUSERS=1 pins it to per-machine under the SYSTEM context.
$PerMachineProperties = @('ALLUSERS=1')

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        The download host fronts an error page on the same origin; a body that
        answers 200 with HTML would otherwise stage as a valid-looking MSI and
        fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Test-RVToolsVersionPublished {
    <#
    .SYNOPSIS
        Returns $true when the download host serves an installer for the version.
    .DESCRIPTION
        Uses a header-only request: the payload is several megabytes and the
        walk issues one request per candidate.
    #>
    param([Parameter(Mandatory)][string]$Version)

    $url = [string]::Format($DownloadUrlFormat, $Version)
    $code = (curl.exe -L --silent --show-error --head --output NUL --write-out '%{http_code}' $url) -join ''
    return ($code.Trim() -eq '200')
}


function Get-LatestRVToolsRelease {
    <#
    .SYNOPSIS
        Returns the newest published RVTools version and its MSI URL.
    .DESCRIPTION
        Walks upward from the floor version, preferring the next patch, then
        the first patch of the next minor, then the first patch of the next
        major. The hop cap bounds the request count if the host ever starts
        answering 200 for every path.
    #>
    param([switch]$Quiet)

    Write-Log "RVTools download host        : downloads.dell.com/rvtools" -Quiet:$Quiet

    try {
        if (-not (Test-RVToolsVersionPublished -Version $FloorVersion)) {
            throw "Floor version $FloorVersion is no longer published; the floor needs raising to a version the host still serves."
        }

        $current = [version]$FloorVersion

        for ($hop = 0; $hop -lt $MaxVersionHops; $hop++) {
            $candidates = @(
                ("{0}.{1}.{2}" -f $current.Major, $current.Minor, ($current.Build + 1)),
                ("{0}.{1}.1"   -f $current.Major, ($current.Minor + 1)),
                ("{0}.0.1"     -f ($current.Major + 1)),
                ("{0}.1.1"     -f ($current.Major + 1))
            )

            $next = $null
            foreach ($c in $candidates) {
                if (Test-RVToolsVersionPublished -Version $c) { $next = $c; break }
            }

            if (-not $next) { break }
            $current = [version]$next
        }

        $version = $current.ToString()

        Write-Log "Latest RVTools version       : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = "rvtools$version.msi"
            DownloadUrl = [string]::Format($DownloadUrlFormat, $version)
        }
    }
    catch {
        Write-Log "Failed to get RVTools version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageRVTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "RVTools - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestRVToolsRelease
    if (-not $releaseInfo) { throw "Could not resolve RVTools version." }

    $version     = $releaseInfo.Version
    $msiFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-MsiPayload -Path $localMsi

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

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

    # --- Derive ARP detection from MSI properties ---
    # A per-machine install registers the ARP entry under the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName -ExtraInstallArgs $PerMachineProperties
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Dell" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "RVTools"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart " + ($PerMachineProperties -join ' ')
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("RVTools")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
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

function Invoke-PackageRVTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "RVTools - PACKAGE phase"
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
        $info = Get-LatestRVToolsRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("RVTools GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "RVTools Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "FloorVersion                 : $FloorVersion"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageRVTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageRVTools
    }
    else {
        Invoke-StageRVTools
        Invoke-PackageRVTools
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-rvtools'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
