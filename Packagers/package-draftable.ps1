<#
Vendor: Draftable
App: Draftable Desktop
CMName: Draftable Desktop
VendorUrl: https://www.draftable.com/
CPE: cpe:2.3:a:draftable:draftable_desktop:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://help.draftable.com/hc/en-us
DownloadPageUrl: https://www.draftable.com/desktop
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Draftable Desktop (x64, system-wide MSI) for MECM.

.DESCRIPTION
    Downloads the vendor's evergreen system-wide MSI, stages content to a
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
    Content is staged under: <FileServerPath>\Applications\Draftable\Draftable Desktop\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER OfficeAddins
    Office add-in features to install alongside the desktop application.
    Any of WordAddin, ExcelAddin, PowerPointAddin, OutlookAddin. None by default.

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Draftable Desktop version string and exits.

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
    [ValidateSet('WordAddin','ExcelAddin','OutlookAddin')]
    [string[]]$OfficeAddins = @(),
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
$LatestMsiUrl  = "https://dl.draftable.com/desktop/DraftableDesktopSystem.msi"
$ReleasesUrl   = "https://dl.draftable.com/desktop/RELEASES"

$VendorFolder = "Draftable"
$AppFolder    = "Draftable Desktop"

$BaseDownloadRoot = Join-Path $DownloadRoot "Draftable Desktop"
$MsiFileName      = "DraftableDesktopSystem.msi"

# An explicit ADDLOCAL list installs exactly the named features and nothing
# else, so each feature's parent is listed alongside it: HttpListener sits
# under DmsIntegrations and StartMenuShortcut under Shortcuts. HttpListener is
# the local service the Office add-ins and the shell integration call into and
# is not installed by default, so it belongs in every deployment.
$BaseFeatures = @("DraftableDesktop","DmsIntegrations","HttpListener","Shortcuts","StartMenuShortcut")

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


function Get-LatestDraftableVersion {
    <#
    .SYNOPSIS
        Returns the version the evergreen Draftable Desktop URL currently serves.
    .DESCRIPTION
        The vendor publishes no version feed for the MSI, and the MSI itself is
        a ~600 MB download; the update manifest beside it names the current
        build in its package filename, which keeps a version check cheap. The
        stage phase still reads the MSI ProductVersion and treats that as
        authoritative for detection.
    #>
    param([switch]$Quiet)

    Write-Log "Draftable release manifest   : $ReleasesUrl" -Quiet:$Quiet

    try {
        $text = (curl.exe -L --fail --silent --show-error $ReleasesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Draftable release manifest." }

        $m = [regex]::Match($text, '(?i)DraftableDesktop-(?<ver>\d+(?:\.\d+)+)-')
        if (-not $m.Success) {
            throw "Release manifest carried no DraftableDesktop-<version> package name."
        }

        $version = $m.Groups['ver'].Value
        Write-Log "Latest Draftable version     : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Draftable version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-DraftableAddLocal {
    <#
    .SYNOPSIS
        Returns the ADDLOCAL feature list for the requested add-in selection.
    #>
    $features = @($BaseFeatures)
    if ($OfficeAddins.Count -gt 0) { $features += "OfficeAddins" }
    foreach ($a in $OfficeAddins) {
        if ($features -notcontains $a) { $features += $a }
    }
    return ("ADDLOCAL=" + ($features -join ','))
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDraftable {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Draftable Desktop (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $expectedVersion = Get-LatestDraftableVersion
    if (-not $expectedVersion) { throw "Could not resolve Draftable Desktop version." }

    # --- Download ---
    # The evergreen URL has no version in it, so a cached copy is kept only
    # while it still carries the version the release manifest advertises.
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    $reuseCached = $false
    if (Test-Path -LiteralPath $localMsi) {
        try {
            $cachedProps = Get-MsiPropertyMap -MsiPath $localMsi
            $cachedVersion = $cachedProps["ProductVersion"]
            if ($cachedVersion -and $cachedVersion.StartsWith($expectedVersion)) { $reuseCached = $true }
        }
        catch { $reuseCached = $false }
    }

    if ($reuseCached) {
        Write-Log "Cached MSI matches $expectedVersion. Skipping download."
    }
    else {
        Write-Log "Downloading MSI (~600 MB, this can take several minutes)..."
        Invoke-DownloadWithRetry -Url $LatestMsiUrl -OutFile $localMsi
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
    $addLocal = Get-DraftableAddLocal
    Write-Log "Feature selection            : $addLocal"

    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName -ExtraInstallArgs @($addLocal)
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Draftable" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Draftable Desktop"
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart $addLocal"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("DraftableDesktop")
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

function Invoke-PackageDraftable {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Draftable Desktop (x64) - PACKAGE phase"
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
        $version = Get-LatestDraftableVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Draftable GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Draftable Desktop (x64) Auto-Packager starting"
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
        Invoke-StageDraftable
    }
    elseif ($PackageOnly) {
        Invoke-PackageDraftable
    }
    else {
        Invoke-StageDraftable
        Invoke-PackageDraftable
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-draftable'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
