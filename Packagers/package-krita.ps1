<#
Vendor: KDE
App: Krita
CMName: Krita
VendorUrl: https://krita.org/
CPE: cpe:2.3:a:kde:krita:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://krita.org/en/news/
DownloadPageUrl: https://krita.org/en/download/
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Krita (x64) for MECM.

.DESCRIPTION
    Resolves the newest stable release from the KDE download archive, downloads
    the x64 NSIS installer, stages content to a versioned local folder with ARP
    detection metadata, and creates an MECM Application with registry-based
    detection.

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
    Content is staged under: <FileServerPath>\Applications\KDE\Krita\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Krita).
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
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Krita version string and exits.

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
$ReleaseRootUrl = "https://download.kde.org/stable/krita/"

$VendorFolder = "KDE"
$AppFolder    = "Krita"

$BaseDownloadRoot = Join-Path $DownloadRoot "Krita"

# The x64 installer's InstallDir and its ARP key name are fixed in the vendor's
# NSIS script, so both the uninstall command and detection can be pinned
# without passing /D (which NSIS rejects when the path is quoted).
$InstallPath   = "{0}\Krita (x64)" -f $env:ProgramFiles
$ArpKeyName    = "Krita_x64"
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ArpKeyName"

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


function Get-LatestKritaRelease {
    <#
    .SYNOPSIS
        Returns the newest Krita release version and its x64 installer URL.
    .DESCRIPTION
        The archive keeps one directory per release plus an 'updates' directory,
        and versions carry three or four fields (6.0.2, 6.0.2.1), so the listing
        is sorted as versions rather than trusted in document order. A release
        directory can appear before its Windows installer is published, so the
        highest version that actually carries one wins.
    #>
    param([switch]$Quiet)

    Write-Log "Krita release archive        : $ReleaseRootUrl" -Quiet:$Quiet

    try {
        $rootHtml = (curl.exe -L --fail --silent --show-error $ReleaseRootUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Krita release index: $ReleaseRootUrl" }

        $dirRx = [regex]'href\s*=\s*"(?<ver>\d+(?:\.\d+)+)/"'
        $versions = $dirRx.Matches($rootHtml) |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique

        if (-not $versions -or @($versions).Count -lt 1) {
            throw "Could not locate any Krita release directories in the archive index."
        }

        $ordered = @($versions) | Sort-Object { [version]$_ } -Descending
        foreach ($v in $ordered) {
            $releaseUrl = "$ReleaseRootUrl$v/"
            $html = (curl.exe -L --fail --silent --show-error $releaseUrl) -join "`n"
            if ($LASTEXITCODE -ne 0) { continue }

            $fileName = "krita-x64-$v-setup.exe"
            if ($html -notmatch [regex]::Escape('href="' + $fileName + '"')) { continue }

            Write-Log "Latest Krita version         : $v" -Quiet:$Quiet
            Write-Log "Installer asset              : $fileName" -Quiet:$Quiet

            return [pscustomobject]@{
                Version     = $v
                FileName    = $fileName
                DownloadUrl = "$releaseUrl$fileName"
            }
        }

        throw "No Krita release directory contained an x64 setup executable."
    }
    catch {
        Write-Log "Failed to get Krita version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageKrita {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Krita (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestKritaRelease
    if (-not $releaseInfo) { throw "Could not resolve Krita version." }

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
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand ("{0}\uninstall.exe" -f $InstallPath) `
        -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # The installer writes DisplayVersion under a fixed ARP key name rather than
    # a ProductCode GUID; a version comparison keeps detection true after an
    # in-place upgrade that lands a higher build.
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log "Uninstall command            : $InstallPath\uninstall.exe /S"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Krita"
        Publisher       = "KDE"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("krita")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
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

function Invoke-PackageKrita {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Krita (x64) - PACKAGE phase"
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
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $info = Get-LatestKritaRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Krita GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Krita (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseRootUrl               : $ReleaseRootUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageKrita
    }
    elseif ($PackageOnly) {
        Invoke-PackageKrita
    }
    else {
        Invoke-StageKrita
        Invoke-PackageKrita
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-krita'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
