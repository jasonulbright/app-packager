<#
Vendor: 7-Zip
App: 7-Zip
CMName: 7-Zip
VendorUrl: https://www.7-zip.org/
CPE: cpe:2.3:a:7-zip:7-zip:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.7-zip.org/history.txt
DownloadPageUrl: https://www.7-zip.org/download.html
UpdateCadenceDays: 90

.SYNOPSIS
    Packages 7-Zip (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest 7-Zip x64 MSI from the official 7-zip.org download page,
    stages content to a versioned local folder with ARP detection metadata, and
    creates an MECM Application with registry-based detection.

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
    Content is staged under: <FileServerPath>\Applications\7-Zip\7-Zip\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\7-Zip).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available 7-Zip version string and exits.

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
$DownloadPageUrl = "https://www.7-zip.org/download.html"

$VendorFolder = "7-Zip"
$AppFolder    = "7-Zip"

$BaseDownloadRoot = Join-Path $DownloadRoot "7-Zip"
$MsiFileName      = "7zip-x64.msi"

# --- Functions ---


function Resolve-7ZipX64MsiUrl {
    param([switch]$Quiet)

    Write-Log "7-Zip download page          : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch 7-Zip download page: $DownloadPageUrl" }

        # Typical links: a/7z2501-x64.msi
        $rx = [regex]'href\s*=\s*"(?<href>[^"]*?7z(?<ver>\d{4})-x64\.msi)"'
        $rxMatches = $rx.Matches($html)

        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any x64 MSI links on the download page."
        }

        $candidates = foreach ($m in $rxMatches) {
            [pscustomobject]@{
                Href      = $m.Groups["href"].Value
                VerDigits = [int]$m.Groups["ver"].Value
            }
        }

        $best = $candidates | Sort-Object VerDigits -Descending | Select-Object -First 1
        $base = [uri]"https://www.7-zip.org/"
        $final = ([uri]::new($base, $best.Href)).AbsoluteUri

        if ($final -notmatch '\.msi($|\?)') {
            throw "Resolved URL does not appear to be an MSI: $final"
        }

        Write-Log "Resolved MSI URL             : $final" -Quiet:$Quiet

        return $final
    }
    catch {
        Write-Log "Failed to resolve 7-Zip MSI URL: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-7ZipDisplayVersion {
    param([Parameter(Mandatory)][string]$RawVersion)

    try {
        $v = [version]$RawVersion
        return ("{0:D2}.{1:D2}" -f $v.Major, $v.Minor)
    }
    catch {
        $parts = $RawVersion -split '\.'
        if ($parts.Count -ge 2) { return ("{0}.{1}" -f $parts[0], $parts[1]) }
        return $RawVersion
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-Stage7Zip {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "7-Zip (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $msiUrl = Resolve-7ZipX64MsiUrl
    if (-not $msiUrl) { throw "Could not resolve 7-Zip MSI download URL." }

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""
    Write-Log "Downloading MSI..."
    Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]   # e.g. 25.01.00.0
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))       { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    $displayVersion = Get-7ZipDisplayVersion -RawVersion $productVersionRaw  # e.g. 25.01

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion (raw)     : $productVersionRaw"
    Write-Log "Version (display)            : $displayVersion"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $displayVersion
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
    # This avoids a temp install/uninstall cycle that can crash Explorer when the
    # product registers shell extensions (e.g. 7-Zip context menu).
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    $arpEntry = [pscustomobject]@{
        RegistryKeyRelative = $arpRegistryKey
        DisplayName         = $productName
        DisplayVersion      = $productVersionRaw   # raw MSI ProductVersion = what Windows writes to registry
        Is64Bit             = $true
    }
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."

    Write-Log ""
    Write-Log "ARP DisplayName              : $($arpEntry.DisplayName)"
    Write-Log "ARP DisplayVersion           : $($arpEntry.DisplayVersion)"
    Write-Log "ARP RegistryKey              : $($arpEntry.RegistryKeyRelative)"
    Write-Log "ARP Is64Bit                  : $($arpEntry.Is64Bit)"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Igor Pavlov" }

    $appName = $productName

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $displayVersion
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("7zFM", "7zG")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpEntry.RegistryKeyRelative
            ValueName           = "DisplayVersion"
            DisplayName         = $arpEntry.DisplayName
            DisplayVersion      = $arpEntry.DisplayVersion
            Is64Bit             = $arpEntry.Is64Bit
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-Package7Zip {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "7-Zip (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    if (-not (Test-Path -LiteralPath $localMsi)) {
        throw "Local MSI not found - run Stage phase first: $localMsi"
    }

    $props = Get-MsiPropertyMap -MsiPath $localMsi
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $displayVersion   = Get-7ZipDisplayVersion -RawVersion $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $displayVersion
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
        Initialize-Folder -Path $BaseDownloadRoot

        $msiUrl = Resolve-7ZipX64MsiUrl -Quiet
        if (-not $msiUrl) { exit 1 }

        $localMsi = Join-Path $BaseDownloadRoot "7zip-x64.msi"
        Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi -Quiet

        $props = Get-MsiPropertyMap -MsiPath $localMsi
        if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) { exit 1 }

        $normalized = Get-7ZipDisplayVersion -RawVersion $props["ProductVersion"]
        Write-Output $normalized
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("7-Zip GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "7-Zip (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-Stage7Zip
    }
    elseif ($PackageOnly) {
        Invoke-Package7Zip
    }
    else {
        Invoke-Stage7Zip
        Invoke-Package7Zip
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
