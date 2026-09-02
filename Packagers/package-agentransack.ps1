<#
Vendor: Mythicsoft
App: Agent Ransack
CMName: Agent Ransack
VendorUrl: https://www.mythicsoft.com/agentransack/
CPE: cpe:2.3:a:mythicsoft:agent_ransack:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.mythicsoft.com/agentransack/whatsnew/
DownloadPageUrl: https://www.mythicsoft.com/agentransack/download/
UpdateCadenceDays: 180

.SYNOPSIS
    Packages Agent Ransack (x64) MSI for MECM.

.DESCRIPTION
    Resolves the current x64 MSI archive from the vendor download page,
    expands it, stages content to a versioned local folder with ARP detection
    metadata, and creates an MECM Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The vendor ships the enterprise MSI inside a zip and publishes only a build
    number on the page, so the full product version comes from the MSI itself.
    GetLatestVersionOnly reuses an already-downloaded MSI whose version carries
    the advertised build rather than re-transferring the archive.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Mythicsoft\Agent Ransack\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Agent Ransack).
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
    Outputs only the latest available Agent Ransack version string and exits.

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
$DownloadPageUrl = "https://www.mythicsoft.com/agentransack/download/"

$VendorFolder = "Mythicsoft"
$AppFolder    = "Agent Ransack"

$BaseDownloadRoot = Join-Path $DownloadRoot "Agent Ransack"

# --- Functions ---


function Assert-PayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the file starts with the OLE compound-file signature that
        every Windows Installer package carries.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $expected = @(0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $buffer = New-Object byte[] 8
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 8) } finally { $stream.Dispose() }

    if ($read -lt 8) { throw "Payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($buffer[$i] -ne $expected[$i]) {
            throw "Payload is not a Windows Installer package (no OLE compound-file header): $Path"
        }
    }
}


function Resolve-AgentRansackMsiArchive {
    <#
    .SYNOPSIS
        Returns the download URL and build number of the current x64 MSI archive.
    .DESCRIPTION
        The archive link carries a per-release token path segment, so the href
        is read from the page rather than composed. Links are protocol-relative.
        The page lists more than one architecture, so the x64 archive is
        selected by name and the highest build wins.
    #>
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Agent Ransack download page: $DownloadPageUrl" }

    $rx = [regex]'href\s*=\s*"(?<href>[^"]*?agentransack_x64_msi_(?<build>\d+)\.zip)"'
    $rxMatches = $rx.Matches($html)
    if (-not $rxMatches -or $rxMatches.Count -lt 1) {
        throw "Could not locate an x64 MSI archive link on the download page."
    }

    $best = $rxMatches |
        ForEach-Object { [pscustomobject]@{ Href = $_.Groups['href'].Value; Build = [int]$_.Groups['build'].Value } } |
        Sort-Object Build -Descending |
        Select-Object -First 1

    $url = ([uri]::new([uri]$DownloadPageUrl, $best.Href)).AbsoluteUri

    Write-Log "Resolved MSI archive URL     : $url" -Quiet:$Quiet
    Write-Log "Advertised build             : $($best.Build)" -Quiet:$Quiet

    return [pscustomobject]@{ Url = $url; Build = $best.Build }
}


function Get-AgentRansackMsi {
    <#
    .SYNOPSIS
        Returns the path to a verified local MSI for the advertised build,
        downloading and expanding the archive when it is not already cached.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Archive,
        [switch]$Quiet
    )

    Initialize-Folder -Path $BaseDownloadRoot

    $msiPath = Join-Path $BaseDownloadRoot ("agentransack_x64_{0}.msi" -f $Archive.Build)
    if (Test-Path -LiteralPath $msiPath) {
        Write-Log "Local MSI exists. Skipping download." -Quiet:$Quiet
        Assert-PayloadIsMsi -Path $msiPath
        return $msiPath
    }

    $zipPath = Join-Path $BaseDownloadRoot ("agentransack_x64_msi_{0}.zip" -f $Archive.Build)
    if (-not (Test-Path -LiteralPath $zipPath)) {
        Write-Log "Downloading MSI archive..." -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $Archive.Url -OutFile $zipPath -Quiet:$Quiet
    }

    $expandPath = Join-Path $BaseDownloadRoot ("expand-{0}" -f $Archive.Build)
    if (Test-Path -LiteralPath $expandPath) { Remove-Item -LiteralPath $expandPath -Recurse -Force -ErrorAction Stop }
    Initialize-Folder -Path $expandPath
    Expand-Archive -LiteralPath $zipPath -DestinationPath $expandPath -Force -ErrorAction Stop

    $extracted = Get-ChildItem -LiteralPath $expandPath -Filter '*.msi' -File -Recurse -ErrorAction Stop |
        Select-Object -First 1
    if (-not $extracted) { throw "The downloaded archive contained no MSI: $zipPath" }

    Assert-PayloadIsMsi -Path $extracted.FullName
    Move-Item -LiteralPath $extracted.FullName -Destination $msiPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $expandPath -Recurse -Force -ErrorAction SilentlyContinue

    Write-Log "Extracted MSI                : $msiPath" -Quiet:$Quiet
    return $msiPath
}


function Get-LatestAgentRansackVersion {
    param([switch]$Quiet)

    try {
        $archive = Resolve-AgentRansackMsiArchive -Quiet:$Quiet
        $msiPath = Get-AgentRansackMsi -Archive $archive -Quiet:$Quiet

        $props = Get-MsiPropertyMap -MsiPath $msiPath
        $version = $props["ProductVersion"]
        if ([string]::IsNullOrWhiteSpace($version)) { throw "MSI ProductVersion missing." }

        Write-Log "Latest Agent Ransack version : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Agent Ransack version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAgentRansack {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Agent Ransack (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $archive = Resolve-AgentRansackMsiArchive
    $localMsi = Get-AgentRansackMsi -Archive $archive
    $msiFileName = Split-Path -Leaf $localMsi

    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName    = $props["ProductName"]
    $productVersion = $props["ProductVersion"]
    $manufacturer   = $props["Manufacturer"]
    $productCode    = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))    { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "MSI ProductCode missing." }

    # The page advertises only a build number; it appears as the third field of
    # the MSI version. A mismatch means the archive and the page disagree.
    $versionFields = $productVersion -split '\.'
    if ($versionFields.Count -lt 3 -or $versionFields[2] -ne [string]$archive.Build) {
        throw "MSI ProductVersion ($productVersion) does not carry the advertised build ($($archive.Build))."
    }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersion"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
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
    # For a standard MSI install the ARP uninstall key name is the ProductCode
    # GUID, so detection needs no temp install/uninstall cycle.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersion"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrappers = New-MsiWrapperContent -MsiFileName $msiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Mythicsoft Ltd" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Agent Ransack"
        Publisher       = $publisher
        SoftwareVersion = $productVersion
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("AgentRansack")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersion
            Is64Bit             = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $productVersion -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageAgentRansack {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Agent Ransack (x64) - PACKAGE phase"
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

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
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
        $v = Get-LatestAgentRansackVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Agent Ransack GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Agent Ransack (x64) Auto-Packager starting"
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
        Invoke-StageAgentRansack
    }
    elseif ($PackageOnly) {
        Invoke-PackageAgentRansack
    }
    else {
        Invoke-StageAgentRansack
        Invoke-PackageAgentRansack
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-agentransack'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
