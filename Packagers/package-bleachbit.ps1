<#
Vendor: BleachBit
App: BleachBit
CMName: BleachBit
VendorUrl: https://www.bleachbit.org/
CPE: cpe:2.3:a:bleachbit:bleachbit:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.bleachbit.org/news
DownloadPageUrl: https://www.bleachbit.org/download/windows
UpdateCadenceDays: 90

.SYNOPSIS
    Packages BleachBit for MECM.

.DESCRIPTION
    Reads the current stable Windows release from the vendor download page,
    downloads the matching NSIS installer, stages content to a versioned local
    folder, and creates an MECM Application with file-version-based detection
    on the installed bleachbit.exe.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\BleachBit\BleachBit\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\BleachBit).
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
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available BleachBit version string and exits.

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
$DownloadPageUrl = "https://www.bleachbit.org/download/windows"
$DownloadHost    = "https://download.bleachbit.org"

$VendorFolder = "BleachBit"
$AppFolder    = "BleachBit"

$BaseDownloadRoot = Join-Path $DownloadRoot "BleachBit"

# The installer is a 32-bit NSIS package built with
# MULTIUSER_INSTALLMODE_64_BIT 0, so a per-machine install lands in the 32-bit
# Program Files tree on x64.
$InstallDir = "{0}\BleachBit" -f ${env:ProgramFiles(x86)}

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor download host answers 200 with an HTML interstitial for the
        /get/ path, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestBleachBitRelease {
    <#
    .SYNOPSIS
        Returns the current stable BleachBit version and its direct installer URL.
    .DESCRIPTION
        The vendor download page is the version source rather than the GitHub
        releases API: Windows installers accompany only some tags there, and the
        release marked latest by the API is frequently a beta.

        The page links the /get/ interstitial, not the payload; the direct URL is
        read from that page so a host change is followed, with the known host as
        the fallback.
    #>
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch BleachBit download page: $DownloadPageUrl" }

        # The English-only build carries a different filename; the localized
        # build is the default download and is the one packaged here.
        $rx = [regex]'BleachBit-(?<ver>\d+(?:\.\d+)+)-setup\.exe'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate a Windows installer link on the download page."
        }

        $version = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        $fileName = "BleachBit-$version-setup.exe"

        Write-Log "Latest BleachBit version     : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = (Resolve-BleachBitInstallerUrl -FileName $fileName -Quiet:$Quiet)
        }
    }
    catch {
        Write-Log "Failed to get BleachBit version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Resolve-BleachBitInstallerUrl {
    param(
        [Parameter(Mandatory)][string]$FileName,
        [switch]$Quiet
    )

    $fallback     = "$DownloadHost/$FileName"
    $interstitial = "$DownloadHost/get/$FileName"

    try {
        $html = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $interstitial) -join "`n"
        if ($LASTEXITCODE -eq 0) {
            $m = [regex]::Match($html, 'https?://[^"'' ]*?' + [regex]::Escape($FileName))
            if ($m.Success -and $m.Value -ne $interstitial) {
                Write-Log "Resolved installer URL       : $($m.Value)" -Quiet:$Quiet
                return $m.Value
            }
        }
    }
    catch { }

    Write-Log "Falling back to known host   : $fallback" -Level WARN -Quiet:$Quiet
    return $fallback
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBleachBit {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BleachBit - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestBleachBitRelease
    if (-not $releaseInfo) { throw "Could not resolve BleachBit version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', 'Mozilla/5.0')
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # /allusers pins the NsisMultiUser install mode: a silent run does not
    # elevate on its own, and without the flag an already-present per-user
    # install would be preferred over the per-machine one being deployed.
    $uninstallExe = Join-Path $InstallDir "uninstall.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/allusers', '/S'" `
        -UninstallCommand $uninstallExe -UninstallArgs "'/allusers', '/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    # The installed binary carries a four-part build version (6.0.2.3702) while
    # the release is named with three parts, so detection compares versions
    # rather than requiring equality.
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : bleachbit.exe"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "BleachBit"
        Publisher        = "BleachBit.org"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/allusers /S"
        UninstallArgs    = "/allusers /S"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("bleachbit", "bleachbit_console")
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "bleachbit.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
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

function Invoke-PackageBleachBit {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BleachBit - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $info = Get-LatestBleachBitRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("BleachBit GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BleachBit Auto-Packager starting"
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
        Invoke-StageBleachBit
    }
    elseif ($PackageOnly) {
        Invoke-PackageBleachBit
    }
    else {
        Invoke-StageBleachBit
        Invoke-PackageBleachBit
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-bleachbit'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
