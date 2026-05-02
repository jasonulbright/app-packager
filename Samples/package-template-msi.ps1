<#
Vendor: TODO
App: TODO
CMName: TODO
VendorUrl: https://TODO
CPE: cpe:2.3:a:TODO:TODO:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://TODO
DownloadPageUrl: https://TODO
UpdateCadenceDays: 30

.SYNOPSIS
    Packages TODO (x64) MSI for MECM.

.DESCRIPTION
    Starter template for an MSI-based packager. Duplicate this file into
    Packagers/package-<app>.ps1, remove the Sample/ path reference, and
    replace every TODO comment with app-specific logic.

    MSI packagers have two small wins over EXE packagers:
      - Get-MsiPropertyMap pulls ProductCode, ProductVersion, Manufacturer,
        ProductName straight out of the MSI. No temp install required.
      - ARP detection uses SOFTWARE\...\Uninstall\<ProductCode> which is
        deterministic for standard MSIs.

    This template targets x64 MSIs by convention. For x86 or
    "ExistsOn64" MSIs, adjust $Is64Bit where noted.

    Read Samples/AUTHORING.md for the full walkthrough.
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

# ---------------------------------------------------------------------------
# Configuration - edit these per app
# ---------------------------------------------------------------------------

# Public download page used for scraping. If the vendor publishes a single
# stable MSI URL (Chrome, Edge), skip the scraping and use it directly.
$DownloadPageUrl = "https://TODO"

# Must match the header Vendor / App values, and become the network share
# segments under \\<share>\Applications\<Vendor>\<App>\<Version>\.
$VendorFolder = "TODO"
$AppFolder    = "TODO"

# Local download root for this packager. One subfolder per app; the
# installer gets re-downloaded here on each Stage.
$BaseDownloadRoot = Join-Path $DownloadRoot $AppFolder
$MsiFileName      = "TODO-installer.msi"


# ---------------------------------------------------------------------------
# Resolve the current MSI URL
# ---------------------------------------------------------------------------
# Implement ONE of these three patterns:
#   1. Scrape the vendor download page for an href matching a version regex.
#   2. Hit a vendor API that returns the latest URL.
#   3. Use a stable URL that redirects to the current version (Chrome, Edge).
#
# See Samples/AUTHORING.md for worked examples.

function Resolve-MsiUrl {
    param([switch]$Quiet)

    Write-Log "Download page                : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch download page" }

        # TODO: replace this regex with one that targets the correct href.
        # Capture the version as digits so we can sort candidates and pick
        # the highest. Narrow the regex; do not anchor on non-stable text.
        $rx = [regex]'href\s*=\s*"(?<href>[^"]*?TODO-(?<ver>\d+(?:\.\d+)*)-x64\.msi)"'
        $matches = $rx.Matches($html)
        if ($matches.Count -lt 1) { throw "No MSI links on download page" }

        $best = $matches |
            ForEach-Object { [pscustomobject]@{ Href = $_.Groups['href'].Value; Ver = $_.Groups['ver'].Value } } |
            Sort-Object { [version]($_.Ver + '.0') } -Descending |
            Select-Object -First 1

        $base  = [uri]$DownloadPageUrl
        $final = ([uri]::new($base, $best.Href)).AbsoluteUri
        Write-Log "Resolved MSI URL             : $final" -Quiet:$Quiet
        return $final
    }
    catch {
        Write-Log ("Failed to resolve MSI URL: " + $_.Exception.Message) -Level ERROR
        return $null
    }
}


# Optional: if the vendor publishes a pretty display version different from
# the raw MSI ProductVersion (e.g. 7-Zip: 26.00 vs 26.00.00.0), normalize here.
# If the MSI ProductVersion IS the display version, just return $RawVersion.
function Get-DisplayVersion {
    param([Parameter(Mandatory)][string]$RawVersion)
    return $RawVersion  # TODO: customize if needed
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageApp {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $msiUrl = Resolve-MsiUrl
    if (-not $msiUrl) { throw "Could not resolve MSI download URL." }

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""
    Write-Log "Downloading MSI..."
    Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi

    # Pull the canonical ARP values straight from the MSI. No temp install.
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))       { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    $displayVersion = Get-DisplayVersion -RawVersion $productVersionRaw

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion (raw)     : $productVersionRaw"
    Write-Log "Version (display)            : $displayVersion"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # Versioned staged folder. Sibling of the raw downloaded MSI.
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

    # ARP uninstall key path is the ProductCode for any standard MSI.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log ""
    Write-Log "ARP DisplayName              : $productName"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log ""

    # install.bat / install.ps1 / uninstall.bat / uninstall.ps1 wrapper set.
    $wrapperContent = New-MsiWrapperContent -MsiFileName $MsiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # Stage manifest: everything Package needs to create the MECM app.
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $productName
        Publisher       = $manufacturer
        SoftwareVersion = $displayVersion
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @()   # TODO: list exe names so MECM can close them before upgrade
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            DisplayName         = $productName
            DisplayVersion      = $productVersionRaw  # raw = what Windows writes to registry
            Is64Bit             = $true               # TODO: set false for 32-bit MSIs
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageApp {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    if (-not (Test-Path -LiteralPath $localMsi)) {
        throw "Local MSI not found - run Stage phase first: $localMsi"
    }

    $props = Get-MsiPropertyMap -MsiPath $localMsi
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $displayVersion   = Get-DisplayVersion -RawVersion $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $displayVersion
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.DisplayVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot     = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    $localFiles = Get-ChildItem -Path $localContentPath -File -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }  # manifest stays local
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


# ---------------------------------------------------------------------------
# -GetLatestVersionOnly mode: write just the version string and exit
# ---------------------------------------------------------------------------
# Used by the main grid's Latest column and by One Click's cadence gate.
# Must be fast: no staging, no wrappers, no manifest.

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        Initialize-Folder -Path $BaseDownloadRoot

        $msiUrl = Resolve-MsiUrl -Quiet
        if (-not $msiUrl) { exit 1 }

        $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
        Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi -Quiet

        $props = Get-MsiPropertyMap -MsiPath $localMsi
        if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) { exit 1 }

        $normalized = Get-DisplayVersion -RawVersion $props["ProductVersion"]
        Write-Output $normalized
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if ($StageOnly)        { Invoke-StageApp }
    elseif ($PackageOnly)  { Invoke-PackageApp }
    else                   { Invoke-StageApp; Invoke-PackageApp }

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
