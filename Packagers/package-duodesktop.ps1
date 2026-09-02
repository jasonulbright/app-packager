<#
Vendor: Cisco
App: Duo Desktop
CMName: Duo Desktop
VendorUrl: https://duo.com/
CPE: cpe:2.3:a:duo:duo_desktop:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://duo.com/docs/checklist-release-notes
DownloadPageUrl: https://duo.com/docs/duo-desktop
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Duo Desktop (Device Health, x64) MSI for MECM.

.DESCRIPTION
    Downloads the vendor's evergreen Duo Desktop MSI, stages content to a
    versioned local folder with ARP detection metadata, and creates an MECM
    Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Cisco\Duo Desktop\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Duo Desktop version string and exits.

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
$LatestMsiUrl = "https://dl.duosecurity.com/DuoDesktop-latest.msi"

$VendorFolder = "Cisco"
$AppFolder    = "Duo Desktop"

$BaseDownloadRoot = Join-Path $DownloadRoot "Duo Desktop"
$MsiFileName      = "DuoDesktop.msi"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        The evergreen URL is served through a CDN that answers 200 with an HTML
        error body on a bad edge, which would otherwise stage as a valid-looking
        MSI and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestDuoDesktopVersion {
    <#
    .SYNOPSIS
        Returns the version the evergreen Duo Desktop URL currently resolves to.
    .DESCRIPTION
        The vendor publishes no version feed; the only cheap signal is the
        Content-Disposition filename on the evergreen MSI, which carries the
        build (DuoDesktop-<version>.msi). A response without that header means
        the endpoint changed shape and the version cannot be trusted.
    #>
    param([switch]$Quiet)

    Write-Log "Duo Desktop MSI URL          : $LatestMsiUrl" -Quiet:$Quiet

    try {
        $headers = (curl.exe -sIL --fail --show-error $LatestMsiUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the Duo Desktop download endpoint." }

        $m = [regex]::Match($headers, '(?im)^\s*content-disposition:.*filename="?DuoDesktop-(?<ver>\d+(?:\.\d+)+)\.msi"?')
        if (-not $m.Success) {
            throw "Content-Disposition did not carry a DuoDesktop-<version>.msi filename."
        }

        # The filename carries three fields while the MSI ProductVersion the
        # stage phase records carries four; padding keeps a version check from
        # reporting an update on every run against an unchanged build.
        $parts = @($m.Groups['ver'].Value -split '\.')
        while ($parts.Count -lt 4) { $parts += '0' }
        $version = ($parts[0..3] -join '.')

        Write-Log "Latest Duo Desktop version   : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Duo Desktop version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDuoDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Duo Desktop (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    # The evergreen URL has no version in it, so the cached copy is always
    # refreshed: a stale local file would silently re-stage an old build.
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""
    Write-Log "Downloading MSI..."
    Invoke-DownloadWithRetry -Url $LatestMsiUrl -OutFile $localMsi

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
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
    Write-Log "Copied MSI to staged folder  : $stagedMsi"

    # --- Derive ARP detection from MSI properties ---
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    # LAUNCH=False keeps the tray client from starting during a SYSTEM-context
    # install, where it would run in session 0 with no user to interact with.
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName -ExtraInstallArgs @("LAUNCH=False")
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Cisco Systems, Inc." }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Duo Desktop"
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart LAUNCH=False"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("DuoDeviceHealth")
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

function Invoke-PackageDuoDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Duo Desktop (x64) - PACKAGE phase"
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
        $version = Get-LatestDuoDesktopVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Duo Desktop GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Duo Desktop (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LatestMsiUrl                 : $LatestMsiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageDuoDesktop
    }
    elseif ($PackageOnly) {
        Invoke-PackageDuoDesktop
    }
    else {
        Invoke-StageDuoDesktop
        Invoke-PackageDuoDesktop
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-duodesktop'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
