<#
Vendor: TortoiseSVN
App: TortoiseSVN
CMName: TortoiseSVN
VendorUrl: https://tortoisesvn.net/
CPE: cpe:2.3:a:tortoisesvn:tortoisesvn:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://tortoisesvn.net/Changelog.txt
DownloadPageUrl: https://tortoisesvn.net/downloads.html
IconSource: Installer
LocalSource: Required
WsusSupport: Yes
ArpName: TortoiseSVN*

.SYNOPSIS
    Packages TortoiseSVN (x64) MSI for ConfigMgr.

.DESCRIPTION
    Stages the TortoiseSVN x64 MSI from the installer source folder of this
    packager (Options > Applications > Local Installer Sources), stages content
    to a versioned local folder with ARP detection metadata, and creates a
    ConfigMgr Application with registry-based detection.

    The TortoiseSVN site protects its download page against scripts, so the
    packager has no download and does not try to get around the protection.
    Download the x64 MSI by hand and put it in the folder. The file name rule is
    the vendor name: TortoiseSVN-<version>-x64-svn-<svn version>.msi (for example
    TortoiseSVN-1.14.9.29743-x64-svn-1.14.5.msi). When the folder holds more
    than one match, the file with the highest ProductVersion is staged. The
    version comes from the ProductVersion of that file, not from a web page.

    Supports two-phase operation:
      -StageOnly    Read the MSI from the folder, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER SourceFolder
    The folder that holds the TortoiseSVN x64 MSI. It replaces the folder saved in
    Options for this run.

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    There is no web source to read. Outputs the ProductVersion of the newest
    matching MSI in the installer source folder and exits, or fails with a
    message that names the folder when it holds no matching file. No ConfigMgr
    changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - The TortoiseSVN x64 MSI in the installer source folder
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
    [string]$SourceFolder = "",
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
# The installer comes from the installer source folder (no download). The
# name rule is the vendor name of the x64 MSI.
$PackagerName    = "package-tortoisesvn"
$InstallerFilter = "TortoiseSVN-*-x64-*.msi"
$FolderHint = "Download the 64-bit MSI from tortoisesvn.net by hand (TortoiseSVN-<version>-x64-svn-<svn version>.msi) and put it in that folder. The packager does not download it."

$VendorFolder = "TortoiseSVN"
$AppFolder    = "TortoiseSVN"

$BaseDownloadRoot = Join-Path $DownloadRoot "TortoiseSVN"

# --- Functions ---


function Resolve-TortoiseSVNMsi {
    <#
    .SYNOPSIS
        The MSI in the installer source folder with the highest ProductVersion.
        Throws a message that names the folder and the file to put there.
    #>
    return (Resolve-LocalSourceInstaller -PackagerName $PackagerName -Filter $InstallerFilter -Override $SourceFolder -Hint $FolderHint)
}


function Get-LatestTortoiseSVNRelease {
    param([switch]$Quiet)

    try {
        $localMsi = Resolve-TortoiseSVNMsi
        Write-Log "Installer source             : $localMsi" -Quiet:$Quiet
        $version = (Get-MsiPropertyMap -MsiPath $localMsi)['ProductVersion']
        if ([string]::IsNullOrWhiteSpace($version)) { throw "MSI ProductVersion missing: $localMsi" }
        Write-Log "Latest TortoiseSVN version   : $version" -Quiet:$Quiet
        return @{ Version = $version; Path = $localMsi }
    }
    catch {
        # The vendor site blocks scripts, so there is no web version to read; say where the version comes from.
        Write-Log "TortoiseSVN has no web version source: the version is read from the MSI in the installer source folder. $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTortoiseSVN {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TortoiseSVN (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Installer from the installer source folder ---
    $localMsi    = Resolve-TortoiseSVNMsi
    $MsiFileName = Split-Path -Leaf $localMsi
    Write-Log "Installer source             : $localMsi"

    # The file comes from the operator's folder. A valid signature is logged; an unsigned file is logged too, not refused.
    $signature = Get-AuthenticodeSignature -LiteralPath $localMsi
    if ($signature.Status -eq 'Valid') {
        Write-Log "Installer signer             : $($signature.SignerCertificate.Subject)"
    }
    else {
        Write-Log "Installer signature is $($signature.Status): $localMsi" -Level WARN
    }

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $($props['ProductName'])"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $($props['Manufacturer'])"
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
    $publisher = "TortoiseSVN"
    $appName = $productName

    # TortoiseSVN's MSI ProductVersion has 3 parts ("1.14.29743"), because the
    # MSI version format cannot hold 4 parts with a large build number. The file
    # name carries 4 parts ("1.14.9.29743"), but the version comes from the file
    # content, so SoftwareVersion, Check Latest and the ARP detection all use the
    # ProductVersion, which is the value that the install writes to DisplayVersion.
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("TortoiseProc")
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

function Invoke-PackageTortoiseSVN {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TortoiseSVN (x64) - PACKAGE phase"
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
        $info = Get-LatestTortoiseSVNRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TortoiseSVN (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTortoiseSVN
    }
    elseif ($PackageOnly) {
        Invoke-PackageTortoiseSVN
    }
    else {
        Invoke-StageTortoiseSVN
        Invoke-PackageTortoiseSVN
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-tortoisesvn'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
