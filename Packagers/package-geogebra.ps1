<#
Vendor: International GeoGebra Institute
App: GeoGebra Classic 6
CMName: GeoGebra Classic
VendorUrl: https://www.geogebra.org/
CPE: cpe:2.3:a:geogebra:geogebra:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.geogebra.org/m/mMcexbAF
DownloadPageUrl: https://www.geogebra.org/download
UpdateCadenceDays: 60

.SYNOPSIS
    Packages GeoGebra Classic 6 MSI for MECM.

.DESCRIPTION
    Reads the current Classic 6 build from the vendor installer index, downloads
    the matching Windows MSI, stages content to a versioned local folder, and
    creates an MECM Application with registry-value detection.

    Two properties of this MSI drive the packaging:

      - It carries no ALLUSERS property, so it installs per-user unless the
        command line sets ALLUSERS=1. Both wrappers pass it.
      - Every build ships the same ProductCode and no UpgradeCode, so a newer
        build cannot major-upgrade an installed one. The generated install.ps1
        adds REINSTALL=ALL REINSTALLMODE=vomus when the ProductCode is already
        registered, and omits it on a clean machine where those properties are
        an error.

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
    Maximum allowed runtime in minutes. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available GeoGebra Classic 6 version string and exits.

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
    [int]$EstimatedRuntimeMins = 15,
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
$InstallerIndexUrl = "https://download.geogebra.org/installers/6.0/"
$VersionMarkerUrl  = "https://download.geogebra.org/installers/6.0/version.txt"

$VendorFolder = "International GeoGebra Institute"
$AppFolder    = "GeoGebra Classic"

$BaseDownloadRoot = Join-Path $DownloadRoot "GeoGebra"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        A CDN error page answers 200 with HTML, which would otherwise stage as a
        valid-looking MSI and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestGeoGebraRelease {
    <#
    .SYNOPSIS
        Returns the current GeoGebra Classic 6 build and its Windows MSI URL.
    .DESCRIPTION
        version.txt names the current build in the installer filename's dashed
        form (for example 6-0-929-3). The index is read as well so a stale
        marker cannot point at a file that is not published; the highest
        published build wins when the two disagree.
    #>
    param([switch]$Quiet)

    Write-Log "GeoGebra installer index     : $InstallerIndexUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $InstallerIndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the GeoGebra installer index." }

        $rx = [regex]'GeoGebra-Windows-Installer-(?<raw>\d+-\d+-\d+-\d+)\.msi'
        $builds = $rx.Matches($html) | ForEach-Object { $_.Groups['raw'].Value } | Sort-Object -Unique

        if (-not $builds -or @($builds).Count -lt 1) {
            throw "Could not locate any GeoGebra Windows MSI builds in the installer index."
        }

        $marker = (curl.exe -L --fail --silent --show-error $VersionMarkerUrl) -join ''
        $marker = $marker.Trim()
        if ($LASTEXITCODE -ne 0) { $marker = '' }

        $candidates = @($builds)
        if ($marker -and ($candidates -contains $marker)) {
            Write-Log "version.txt marker           : $marker" -Quiet:$Quiet
        }
        elseif ($marker) {
            Write-Log "version.txt marker $marker is not published in the index; falling back to the index." -Level WARN -Quiet:$Quiet
        }

        $raw = $candidates | Sort-Object { [version]($_ -replace '-', '.') } | Select-Object -Last 1
        $version = $raw -replace '-', '.'

        Write-Log "Latest GeoGebra build        : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = "GeoGebra-Windows-Installer-$raw.msi"
            DownloadUrl = "$InstallerIndexUrl" + "GeoGebra-Windows-Installer-$raw.msi"
        }
    }
    catch {
        Write-Log "Failed to get GeoGebra version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function New-GeoGebraInstallContent {
    <#
    .SYNOPSIS
        Returns install.ps1 content that repairs-over-installs an already
        registered ProductCode.
    .DESCRIPTION
        Every GeoGebra build reuses one ProductCode and publishes no UpgradeCode,
        so msiexec treats a newer build over an older one as a reinstall of the
        same product rather than an upgrade. REINSTALL=ALL REINSTALLMODE=vomus
        forces the file replacement in that case, and is omitted otherwise
        because msiexec rejects it when the product is not installed.
    #>
    param(
        [Parameter(Mandatory)][string]$MsiFileName,
        [Parameter(Mandatory)][string]$ProductCode
    )

    $MsiFileName = $MsiFileName -replace "'", "''"
    $ProductCode = $ProductCode -replace "'", "''"

    return (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        ('$productCode = ''{0}''' -f $ProductCode),
        '$installed = @(',
        '    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$productCode",',
        '    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$productCode"',
        ') | Where-Object { Test-Path -LiteralPath $_ }',
        '$msiArgs = @(''/i'', $msiPath, ''/qn'', ''/norestart'', ''ALLUSERS=1'')',
        'if ($installed) { $msiArgs += @(''REINSTALL=ALL'', ''REINSTALLMODE=vomus'') }',
        '$proc = Start-Process -FilePath ''msiexec.exe'' -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGeoGebra {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GeoGebra Classic 6 - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestGeoGebraRelease
    if (-not $releaseInfo) { throw "Could not resolve GeoGebra version." }

    $msiFileName = $releaseInfo.FileName

    Write-Log "Version                      : $($releaseInfo.Version)"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log "MSI filename                 : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
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

    $version = $productVersionRaw

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
    # The package template is Intel (32-bit), so its ARP key lands in the 32-bit
    # registry view.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey (32-bit view)"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $installContent = New-GeoGebraInstallContent -MsiFileName $msiFileName -ProductCode $productCode

    $msiWrappers = New-MsiWrapperContent -MsiFileName $msiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $msiWrappers.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "International GeoGebra Institute" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "GeoGebra Classic"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart ALLUSERS=1"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("GeoGebra")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $false
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

function Invoke-PackageGeoGebra {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GeoGebra Classic 6 - PACKAGE phase"
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
        $info = Get-LatestGeoGebraRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GeoGebra GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GeoGebra Classic 6 Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "InstallerIndexUrl            : $InstallerIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGeoGebra
    }
    elseif ($PackageOnly) {
        Invoke-PackageGeoGebra
    }
    else {
        Invoke-StageGeoGebra
        Invoke-PackageGeoGebra
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-geogebra'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
