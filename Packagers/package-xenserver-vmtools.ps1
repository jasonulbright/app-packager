<#
Vendor: Cloud Software Group
App: XenServer VM Tools for Windows
CMName: XenServer VM Tools
VendorUrl: https://www.xenserver.com/
CPE: cpe:2.3:a:citrix:xenserver_vm_tools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.xenserver.com/en-us/xenserver/8/vms/windows/vm-tools.html
DownloadPageUrl: https://www.xenserver.com/downloads
UpdateCadenceDays: 90

.SYNOPSIS
    Packages XenServer VM Tools for Windows (x64) MSI for MECM.

.DESCRIPTION
    Scrapes the XenServer downloads page for the current management agent MSI,
    stages content to a versioned local folder with ARP detection metadata
    derived from MSI properties, and creates an MECM Application with
    registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The install disables the agent's own update path (ALLOWAUTOUPDATE=NO,
    IDENTIFYAUTOUPDATE=NO) so the deployed version stays the version MECM
    detects.

    Installing the tools replaces the guest's storage and network drivers, so
    the deployment type is expected to require a restart on the endpoint.

    GetLatestVersionOnly reads the version out of the downloads page and exits
    without downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Cloud Software Group\XenServer VM Tools\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\XenServerVMTools).
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
    Outputs only the latest available VM Tools version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$DownloadPageUrl = "https://www.xenserver.com/downloads"
$MsiUrlPattern   = 'https://downloads\.xenserver\.com/vm-tools-windows/[^"'']+\.msi'
$VersionPattern  = 'vm-tools-windows/(\d+\.\d+\.\d+)/'

# The agent updates itself out of band unless both properties are set; a
# self-updated agent no longer matches the packaged version MECM detects.
$ExtraInstallProperties = @('ALLOWAUTOUPDATE=NO', 'IDENTIFYAUTOUPDATE=NO')

$VendorFolder = "Cloud Software Group"
$AppFolder    = "XenServer VM Tools"

$BaseDownloadRoot = Join-Path $DownloadRoot "XenServerVMTools"

# --- Functions ---


function Resolve-XenServerVMToolsRelease {
    <#
    .SYNOPSIS
        Returns the newest x64 VM Tools MSI URL and its version from the vendor
        downloads page.
    #>
    param([switch]$Quiet)

    Write-Log "XenServer downloads page     : $DownloadPageUrl" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the XenServer downloads page: $DownloadPageUrl" }

    $urlMatches = [regex]::Matches($html, $MsiUrlPattern)
    if ($urlMatches.Count -lt 1) {
        throw "Could not locate a VM Tools MSI link on $DownloadPageUrl."
    }

    $candidates = foreach ($m in $urlMatches) {
        $url = $m.Value
        # The vendor also publishes an x86 agent under the same folder.
        if ($url -notmatch 'x64') { continue }
        $verMatch = [regex]::Match($url, $VersionPattern)
        if (-not $verMatch.Success) { continue }
        [pscustomobject]@{
            Url     = $url
            Version = $verMatch.Groups[1].Value
        }
    }
    if (-not $candidates) {
        throw "VM Tools MSI links were found but none matched an x64 build with a parseable version segment."
    }

    $best = $candidates | Sort-Object -Property @{ Expression = { [version]$_.Version } } -Descending | Select-Object -First 1

    Write-Log "Latest VM Tools version      : $($best.Version)" -Quiet:$Quiet
    Write-Log "Resolved MSI URL             : $($best.Url)" -Quiet:$Quiet

    return $best
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageXenServerVMTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "XenServer VM Tools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Resolve and download ---
    $release = Resolve-XenServerVMToolsRelease
    $version = $release.Version
    $msiFileName = "managementagentx64-$version.msi"

    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $release.Url -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log ""
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

    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName -ExtraInstallArgs $ExtraInstallProperties
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Cloud Software Group" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "XenServer VM Tools $version"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = (($ExtraInstallProperties -join ' ') + " /qn /norestart")
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @()
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            Operator            = "IsEquals"
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

function Invoke-PackageXenServerVMTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "XenServer VM Tools (x64) - PACKAGE phase"
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
        $release = Resolve-XenServerVMToolsRelease -Quiet
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("XenServer VM Tools GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "XenServer VM Tools (x64) Auto-Packager starting"
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
        Invoke-StageXenServerVMTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageXenServerVMTools
    }
    else {
        Invoke-StageXenServerVMTools
        Invoke-PackageXenServerVMTools
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-xenserver-vmtools'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
