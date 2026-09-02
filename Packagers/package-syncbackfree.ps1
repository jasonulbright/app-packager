<#
Vendor: 2BrightSparks
App: SyncBackFree
CMName: SyncBackFree
VendorUrl: https://www.2brightsparks.com/
CPE: cpe:2.3:a:2brightsparks:syncback_free:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.2brightsparks.com/syncback/changes.html
DownloadPageUrl: https://www.2brightsparks.com/download-syncbackfree.html
UpdateCadenceDays: 90

.SYNOPSIS
    Packages SyncBackFree for MECM.

.DESCRIPTION
    Reads the current version from the vendor download page, downloads the
    administrator (all-users) Inno Setup installer, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The free edition ships as a single 32-bit build; there is no x64 payload to
    stage. The administrator installer always installs for all users, so no
    scope switch is passed.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\2BrightSparks\SyncBackFree\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\SyncBackFree).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available SyncBackFree version string and exits.

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
$DownloadPageUrl = "https://www.2brightsparks.com/download-syncbackfree.html"
$InstallerUrl    = "https://www.2brightsparks.com/assets/software/SyncBack_Setup.exe"
$InstallerFile   = "SyncBack_Setup.exe"

$VendorFolder = "2BrightSparks"
$AppFolder    = "SyncBackFree"

$BaseDownloadRoot = Join-Path $DownloadRoot "SyncBackFree"

# The 32-bit administrator installer places the product under the x86 program
# files tree; the path is fixed in the vendor installer script, so both the
# uninstall command and detection can be pinned to it.
$InstallPath = "{0}\2BrightSparks\SyncBackFree" -f ${env:ProgramFiles(x86)}

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        A mirror that answers 200 with an HTML error body would otherwise stage
        as a valid-looking installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestSyncBackFreeVersion {
    <#
    .SYNOPSIS
        Returns the SyncBackFree version advertised on the vendor download page.
    .DESCRIPTION
        The installer URL is unversioned and always serves the current build, so
        the page's structured-data softwareVersion field is the only version
        source available before download.
    #>
    param([switch]$Quiet)

    Write-Log "Download page                : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the SyncBackFree download page." }

        $m = [regex]::Match($html, '"softwareVersion"\s*:\s*"(?<ver>\d+(?:\.\d+)+)"')
        if (-not $m.Success) {
            throw "Could not locate a softwareVersion value on the download page."
        }

        $version = $m.Groups['ver'].Value
        Write-Log "Latest SyncBackFree version  : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get SyncBackFree version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSyncBackFree {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SyncBackFree - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $version = Get-LatestSyncBackFreeVersion
    if (-not $version) { throw "Could not resolve SyncBackFree version." }

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFile
    Write-Log "Local installer path         : $localExe"
    Write-Log "Download URL                 : $InstallerUrl"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe

    Assert-ExePayload -Path $localExe

    # The download URL carries no version, so a page that has been updated ahead
    # of the file (or behind it) would otherwise stage a payload whose detection
    # rule can never be true.
    $fileVersion = (Get-Item -LiteralPath $localExe).VersionInfo.ProductVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    Write-Log "Installer ProductVersion     : $fileVersion"
    if ($fileVersion -ne $version) {
        throw "Installer ProductVersion ($fileVersion) does not match the page version ($version)."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFile
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $InstallerFile `
        -InstallArgs "'/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART'" `
        -UninstallCommand ("{0}\unins000.exe" -f $InstallPath) `
        -UninstallArgs "'/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # Detection reads the installed executable rather than an ARP key: the
    # vendor publishes no uninstall key name, and the product exe carries the
    # same version the download page advertises.
    Write-Log ""
    Write-Log "Detection file               : $InstallPath\SyncBackFree.exe"
    Write-Log "Detection version            : $version"
    Write-Log "Uninstall command            : $InstallPath\unins000.exe /VERYSILENT"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "SyncBackFree"
        Publisher       = "2BrightSparks Pte Ltd"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFile
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("SyncBackFree")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallPath
            FileName      = "SyncBackFree.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
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

function Invoke-PackageSyncBackFree {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SyncBackFree - PACKAGE phase"
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
    Write-Log "Detection File               : $($manifest.Detection.FilePath)\$($manifest.Detection.FileName)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $v = Get-LatestSyncBackFreeVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("SyncBackFree GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SyncBackFree Auto-Packager starting"
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
        Invoke-StageSyncBackFree
    }
    elseif ($PackageOnly) {
        Invoke-PackageSyncBackFree
    }
    else {
        Invoke-StageSyncBackFree
        Invoke-PackageSyncBackFree
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-syncbackfree'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
