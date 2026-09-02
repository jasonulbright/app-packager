<#
Vendor: Irfan Skiljan
App: IrfanView
CMName: IrfanView
VendorUrl: https://www.irfanview.com/
CPE: cpe:2.3:a:irfanview:irfanview:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.irfanview.info/main_history.htm
DownloadPageUrl: https://www.irfanview.info/64bit.htm
UpdateCadenceDays: 180

.SYNOPSIS
    Packages IrfanView (x64) for MECM.

.DESCRIPTION
    Resolves the current version from the vendor 64-bit download page,
    downloads the English x64 self-extracting installer, stages content to a
    versioned local folder, and creates an MECM Application with file-based
    detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers and stage manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\IrfanView\IrfanView\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\IrfanView).
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
    Outputs only the latest available IrfanView version string and exits.

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
$Download64Url = "https://www.irfanview.info/64bit.htm"
$FilesBaseUrl  = "https://www.irfanview.info/files/"

$VendorFolder = "IrfanView"
$AppFolder    = "IrfanView"

$BaseDownloadRoot = Join-Path $DownloadRoot "IrfanView"
$InstallPath      = "{0}\IrfanView" -f $env:ProgramFiles

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor host answers a first request for a file path with a 200 and
        an HTML "click again to start download" page, so a payload that skipped
        the referer handshake would otherwise stage as a valid-looking
        installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestIrfanViewRelease {
    <#
    .SYNOPSIS
        Returns the current IrfanView version and its English x64 installer URL.
    .DESCRIPTION
        Filenames encode the version without a separator (iview475 = 4.75), so
        the digits are read from the x64 archive link on the 64-bit page and
        re-expanded. The x64 page is used rather than the general download page
        because the two architectures can carry different current versions.
    #>
    param([switch]$Quiet)

    Write-Log "IrfanView 64-bit page        : $Download64Url" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $Download64Url) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch IrfanView 64-bit download page." }

        $rx = [regex]'iview(?<digits>\d{3,4})_x64\.zip'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate an x64 download link on the 64-bit page."
        }

        $digits = ($rxMatches | ForEach-Object { $_.Groups['digits'].Value } |
            Sort-Object -Unique { [int]$_ } | Select-Object -Last 1)

        $version = "{0}.{1}" -f $digits.Substring(0, 1), $digits.Substring(1)
        $fileName = "iview${digits}_x64_setup.exe"

        Write-Log "Latest IrfanView version     : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = "$FilesBaseUrl$fileName"
        }
    }
    catch {
        Write-Log "Failed to get IrfanView version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageIrfanView {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "IrfanView (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestIrfanViewRelease
    if (-not $releaseInfo) { throw "Could not resolve IrfanView version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName
    $downloadUrl       = $releaseInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading IrfanView..."
        # The host serves an interstitial page unless the request carries the
        # file's own URL as referer; the MZ check below catches a miss.
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe -ExtraCurlArgs @('-e', $downloadUrl)
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
    # /allusers=1 forces the per-machine install a device-targeted deployment
    # requires; without it the installer writes to the invoking profile.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/silent', '/allusers=1', '/desktop=1', '/group=1', '/assoc=1'" `
        -UninstallCommand "$InstallPath\iv_uninstall.exe" `
        -UninstallArgs "'/silent'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : i_view64.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "IrfanView"
        Publisher       = "Irfan Skiljan"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/silent /allusers=1 /desktop=1 /group=1 /assoc=1"
        UninstallArgs   = "/silent"
        RunningProcess  = @("i_view64")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallPath
            FileName      = "i_view64.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageIrfanView {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "IrfanView (x64) - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $info = Get-LatestIrfanViewRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("IrfanView GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "IrfanView (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "Download64Url                : $Download64Url"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageIrfanView
    }
    elseif ($PackageOnly) {
        Invoke-PackageIrfanView
    }
    else {
        Invoke-StageIrfanView
        Invoke-PackageIrfanView
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-irfanview'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
