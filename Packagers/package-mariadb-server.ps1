<#
Vendor: MariaDB
App: MariaDB Server (x64)
CMName: MariaDB Server
VendorUrl: https://mariadb.org/
CPE: cpe:2.3:a:mariadb:mariadb:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://mariadb.com/docs/release-notes/
DownloadPageUrl: https://mariadb.org/download/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages MariaDB Server (x64) MSI for MECM.

.DESCRIPTION
    Queries the MariaDB downloads REST API for the newest stable long-term
    support series, resolves the highest release in that series, downloads the
    winx64 MSI, stages content to a versioned local folder with ARP detection
    metadata, and creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER ServiceName
    Windows service name created by the install. The MSI creates no service at
    all when this is empty, which leaves the package as a client/tools-only
    install. Default: MariaDB

.PARAMETER RootPassword
    Password assigned to the MariaDB root account during install. Leaving this
    empty installs the server with a blank root password, which is only
    defensible on a host where the instance is not reachable off-box; supply a
    value for any deployment that is not a throwaway lab build. The value is
    written into install.ps1 in the staged content, so treat the content folder
    and its network copy as sensitive when it is set.

.PARAMETER Port
    TCP port for the server instance. Default: 3306

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder.

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs the latest available version string and exits.

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
    [string]$ServiceName = "MariaDB",
    [string]$RootPassword = "",
    [int]$Port = 3306,
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 45,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$RestApiRoot = "https://downloads.mariadb.org/rest-api/mariadb/"

$VendorFolder = "MariaDB"
$AppFolder    = "MariaDB Server"
$Publisher    = "MariaDB Corporation Ab"

$BaseDownloadRoot = Join-Path $DownloadRoot "MariaDBServer"

# --- Functions ---


function Assert-MsiPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file signature.
    .DESCRIPTION
        The download URLs point at a mirror network; a mirror that answers 200
        with an HTML error body would otherwise stage as a valid-looking MSI and
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


function Get-LatestMariaDBRelease {
    <#
    .SYNOPSIS
        Returns the newest stable LTS MariaDB Server version and its x64 MSI URL.
    .DESCRIPTION
        The major-release index carries preview and rolling series alongside the
        supported ones, so the series is chosen on release_status Stable plus
        release_support_type Long Term Support and then sorted as a version - the
        index lists 10.x after 12.x, and a string sort would prefer 10.11 over 11.8.

        Within a series the per-release endpoint is likewise not ordered, so the
        release ids are version-sorted before the winx64 MSI is picked out.
    #>
    param([switch]$Quiet)

    Write-Log "MariaDB REST API             : $RestApiRoot" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error $RestApiRoot) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the MariaDB REST API index." }

        $index = ConvertFrom-Json $json

        $ltsSeries = @($index.major_releases |
            Where-Object { $_.release_status -eq 'Stable' -and $_.release_support_type -eq 'Long Term Support' } |
            ForEach-Object { [string]$_.release_id })

        if ($ltsSeries.Count -lt 1) { throw "No stable long-term-support series in the MariaDB index." }

        $series = @($ltsSeries | Sort-Object { [version]$_ })[-1]
        Write-Log "Latest stable LTS series     : $series" -Quiet:$Quiet

        $seriesJson = (& curl.exe -L --fail --silent --show-error ("{0}{1}/" -f $RestApiRoot, $series)) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the MariaDB $series release list." }

        $seriesData = ConvertFrom-Json $seriesJson

        $releaseIds = @($seriesData.releases.PSObject.Properties | ForEach-Object { $_.Name })
        if ($releaseIds.Count -lt 1) { throw "MariaDB series $series lists no releases." }

        $version = @($releaseIds | Sort-Object { [version]$_ })[-1]
        $release = $seriesData.releases.$version

        $file = @($release.files | Where-Object {
            $_.package_type -eq 'MSI Package' -and $_.file_name -like '*winx64.msi'
        }) | Select-Object -First 1

        if (-not $file) { throw "MariaDB $version publishes no winx64 MSI." }

        # The API returns http:// URLs; the same host serves https.
        $downloadUrl = [string]$file.file_download_url -replace '^http://', 'https://'

        Write-Log "Latest MariaDB Server        : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = [string]$file.file_name
            DownloadUrl = $downloadUrl
            Sha256      = [string]$file.checksum.sha256sum
        }
    }
    catch {
        Write-Log "Failed to get MariaDB version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageMariaDB {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MariaDB Server (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-LatestMariaDBRelease
    if (-not $release) { throw "Could not resolve MariaDB Server version." }

    $version     = $release.Version
    $msiFileName = $release.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-MsiPayload -Path $localMsi

    # --- Verify vendor checksum ---
    if (-not [string]::IsNullOrWhiteSpace($release.Sha256)) {
        $actual = (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $release.Sha256.ToUpperInvariant()) {
            throw "SHA256 mismatch for $msiFileName (expected $($release.Sha256), got $actual)."
        }
        Write-Log "SHA256 matches vendor value  : $actual"
    }

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
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    # SERVICENAME is what turns a file-copy install into a running instance: the
    # MSI registers no service when the property is absent or empty.
    $extraArgs = @()
    if (-not [string]::IsNullOrWhiteSpace($ServiceName)) {
        $extraArgs += ("SERVICENAME={0}" -f $ServiceName)
        $extraArgs += ("PORT={0}" -f $Port)
    }
    if (-not [string]::IsNullOrWhiteSpace($RootPassword)) {
        $extraArgs += ("PASSWORD={0}" -f $RootPassword)
    }

    if ([string]::IsNullOrWhiteSpace($ServiceName)) {
        Write-Log "SERVICENAME empty - staging a tools-only install with no service." -Level WARN
    }
    elseif ([string]::IsNullOrWhiteSpace($RootPassword)) {
        Write-Log "No -RootPassword supplied - the instance installs with a blank root password." -Level WARN
    }

    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName -ExtraInstallArgs $extraArgs
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = $Publisher }

    # The install argument list is recorded without the PASSWORD property so the
    # manifest does not carry the secret that install.ps1 already holds.
    $manifestInstallArgs = @("/qn", "/norestart") + @($extraArgs | Where-Object { $_ -notlike 'PASSWORD=*' })

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "MariaDB Server (x64)"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = ($manifestInstallArgs -join ' ')
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("mysqld", "mysql", "mariadbd")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
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

function Invoke-PackageMariaDB {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MariaDB Server (x64) - PACKAGE phase"
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
        $info = Get-LatestMariaDBRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("MariaDB Server GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MariaDB Server (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ServiceName                  : $ServiceName"
    Write-Log "Port                         : $Port"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageMariaDB
    }
    elseif ($PackageOnly) {
        Invoke-PackageMariaDB
    }
    else {
        Invoke-StageMariaDB
        Invoke-PackageMariaDB
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-mariadb-server'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
