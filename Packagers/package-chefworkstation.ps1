<#
Vendor: Chef Software
App: Chef Workstation
CMName: Chef Workstation
VendorUrl: https://www.chef.io/
CPE: cpe:2.3:a:chef:workstation:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/chef/chef-workstation/blob/main/CHANGELOG.md
DownloadPageUrl: https://www.chef.io/downloads/tools/workstation
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Chef Workstation (x64) MSI for MECM.

.DESCRIPTION
    Resolves the current stable build from the vendor package metadata service,
    downloads the x64 MSI, stages content to a versioned local folder with ARP
    detection metadata derived from the MSI property table, and creates an MECM
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
    Content is staged under: <FileServerPath>\Applications\Chef Software\Chef Workstation\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Chef Workstation).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Chef Workstation version string and exits.

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
    [int]$EstimatedRuntimeMins = 20,
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
$DownloadPageUrl = "https://www.chef.io/downloads/tools/workstation"
$MetadataUrl     = "https://omnitruck.chef.io/stable/chef-workstation/metadata?p=windows&pv=2019&m=x86_64"

$VendorFolder = "Chef Software"
$AppFolder    = "Chef Workstation"

$BaseDownloadRoot = Join-Path $DownloadRoot "Chef Workstation"
$MsiFileName      = "chef-workstation-x64.msi"

# --- Functions ---


function Assert-ChefPayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound file signature.
    .DESCRIPTION
        The package host answers 200 with an HTML error body for withdrawn
        builds, which would otherwise stage as a valid-looking MSI.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $expected = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount $expected.Length -ErrorAction Stop
    if ($bytes.Count -lt $expected.Length) {
        throw "Downloaded payload is too small to be an MSI: $Path"
    }
    for ($i = 0; $i -lt $expected.Length; $i++) {
        if ($bytes[$i] -ne $expected[$i]) {
            throw "Downloaded payload is not an MSI (no OLE compound file header): $Path"
        }
    }
}


function Resolve-ChefWorkstationBuild {
    <#
    .SYNOPSIS
        Returns the version and MSI URL of the current stable build.
    .DESCRIPTION
        The metadata endpoint answers with tab-separated key/value lines rather
        than JSON, so the response is parsed line-wise.
    #>
    param([switch]$Quiet)

    Write-Log "Metadata URL                 : $MetadataUrl" -Quiet:$Quiet

    try {
        $lines = curl.exe -L --fail --silent --show-error $MetadataUrl
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Chef Workstation metadata: $MetadataUrl" }

        $map = @{}
        foreach ($line in $lines) {
            $parts = $line -split '\s+', 2
            if ($parts.Count -eq 2) { $map[$parts[0].Trim()] = $parts[1].Trim() }
        }

        if ([string]::IsNullOrWhiteSpace($map['version'])) { throw "Metadata response contained no version." }
        if ([string]::IsNullOrWhiteSpace($map['url']))     { throw "Metadata response contained no url." }
        if ($map['url'] -notmatch '\.msi($|\?)')           { throw "Metadata url is not an MSI: $($map['url'])" }

        Write-Log "Latest Chef Workstation      : $($map['version'])" -Quiet:$Quiet
        Write-Log "Resolved MSI URL             : $($map['url'])" -Quiet:$Quiet

        return [pscustomobject]@{
            Version = $map['version']
            Url     = $map['url']
            Sha256  = $map['sha256']
        }
    }
    catch {
        Write-Log "Failed to resolve Chef Workstation build: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageChefWorkstation {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Chef Workstation (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get build ---
    $build = Resolve-ChefWorkstationBuild
    if (-not $build) { throw "Could not resolve Chef Workstation build." }

    Write-Log "Version (metadata)           : $($build.Version)"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $($build.Url)"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $build.Url -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-ChefPayloadIsMsi -Path $localMsi

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
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
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
    # The MSI Manufacturer property carries a quoted maintainer mailbox, which
    # does not belong in the CM Publisher field.
    $publisher = "Chef Software, Inc."

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Chef Workstation"
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @()
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            DisplayName         = $productName
            DisplayVersion      = $productVersionRaw
            Is64Bit             = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $productVersionRaw -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageChefWorkstation {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Chef Workstation (x64) - PACKAGE phase"
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
    Write-Log "Detection Value              : $($manifest.Detection.DisplayVersion)"
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
        $b = Resolve-ChefWorkstationBuild -Quiet
        if (-not $b) { exit 1 }
        Write-Output $b.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Chef Workstation GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Chef Workstation (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "MetadataUrl                  : $MetadataUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageChefWorkstation
    }
    elseif ($PackageOnly) {
        Invoke-PackageChefWorkstation
    }
    else {
        Invoke-StageChefWorkstation
        Invoke-PackageChefWorkstation
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-chefworkstation'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
