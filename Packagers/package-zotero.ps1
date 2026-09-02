<#
Vendor: Corporation for Digital Scholarship
App: Zotero
CMName: Zotero
VendorUrl: https://www.zotero.org/
CPE: cpe:2.3:a:zotero:zotero:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.zotero.org/support/changelog
DownloadPageUrl: https://www.zotero.org/download/
UpdateCadenceDays: 45

.SYNOPSIS
    Packages Zotero (x64) EXE for MECM.

.DESCRIPTION
    Resolves the current release-channel build from the vendor download
    redirector, downloads the win-x64 setup EXE, stages content to a versioned
    local folder, and creates an MECM Application with file-version detection.

    Zotero ships the Mozilla installer: /S installs silently per-machine and the
    product is removed by uninstall\helper.exe rather than msiexec.

    Supports two-phase operation:
      -StageOnly    Download installer, generate wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Corporation for Digital Scholarship\Zotero\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Zotero).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 10

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers and
    stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Zotero version string and exits.

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
    [int]$EstimatedRuntimeMins = 10,
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
$RedirectorUrl = "https://www.zotero.org/download/client/dl?channel=release&platform=win-x64"

$VendorFolder = "Corporation for Digital Scholarship"
$AppFolder    = "Zotero"

$BaseDownloadRoot = Join-Path $DownloadRoot "Zotero"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        The redirector answers 200 with an HTML page when the requested channel
        has no build, which would otherwise stage as a valid-looking installer.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestZoteroRelease {
    <#
    .SYNOPSIS
        Returns the current Zotero release version and its win-x64 setup EXE URL.
    .DESCRIPTION
        The download page carries no version string; the redirector encodes it
        in the 302 Location. Headers are fetched without following the redirect
        so version resolution does not pull the full installer body.
    #>
    param([switch]$Quiet)

    Write-Log "Download redirector          : $RedirectorUrl" -Quiet:$Quiet

    try {
        $headers = (curl.exe -s -S -I $RedirectorUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Zotero download redirector: $RedirectorUrl" }

        $m = [regex]::Match($headers, '(?im)^location:\s*(?<url>\S+)\s*$')
        if (-not $m.Success) { throw "Zotero redirector did not answer with a Location header." }

        $url = $m.Groups['url'].Value
        $vm = [regex]::Match($url, '/client/release/(?<ver>\d+(?:\.\d+)+)/(?<file>Zotero-[^/]+\.exe)$')
        if (-not $vm.Success) { throw "Could not parse a version and filename from the redirect target: $url" }

        $version = $vm.Groups['ver'].Value

        Write-Log "Latest version               : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $vm.Groups['file'].Value
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get Zotero version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageZotero {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zotero (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestZoteroRelease
    if (-not $releaseInfo) { throw "Could not resolve Zotero version." }

    $version          = $releaseInfo.Version
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
    # The Mozilla installer registers no MSI product, so uninstall goes through
    # the native helper. A missing helper means the product is already gone.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand (Join-Path $env:ProgramFiles 'Zotero\uninstall\helper.exe') `
        -UninstallArgs "'/S'"

    $customUninstall = (
        '$helperPath = Join-Path $env:ProgramFiles ''Zotero\uninstall\helper.exe''',
        'if (-not (Test-Path -LiteralPath $helperPath)) { exit 0 }',
        '$proc = Start-Process -FilePath $helperPath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    $detectionPath = Join-Path $env:ProgramFiles 'Zotero'

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : zotero.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Zotero"
        Publisher       = "Corporation for Digital Scholarship"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("zotero")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "zotero.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageZotero {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zotero (x64) - PACKAGE phase"
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
        $info = Get-LatestZoteroRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Zotero GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zotero (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "RedirectorUrl                : $RedirectorUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageZotero
    }
    elseif ($PackageOnly) {
        Invoke-PackageZotero
    }
    else {
        Invoke-StageZotero
        Invoke-PackageZotero
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-zotero'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
