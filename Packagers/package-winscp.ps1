<#
Vendor: Martin Prikryl
App: WinSCP (x64)
CMName: WinSCP
VendorUrl: https://winscp.net/
CPE: cpe:2.3:a:winscp:winscp:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://winscp.net/eng/docs/history
DownloadPageUrl: https://winscp.net/eng/download.php
IconSource: External
WsusSupport: Yes
ArpName: WinSCP*

.SYNOPSIS
    Packages WinSCP (x64) for ConfigMgr.

.DESCRIPTION
    Downloads the latest WinSCP installer from winscp.net, stages content to a
    versioned local folder, and writes the stage manifest with file-version
    detection of WinSCP.exe. The installer is never run on the packaging machine.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers and stage manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\WinSCP\WinSCP\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\WinSCP).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download, generate content wrappers and stage
    manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file-version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available WinSCP version string and exits.

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
$VendorFolder = "WinSCP"
$AppFolder    = "WinSCP"

$BaseDownloadRoot = Join-Path $DownloadRoot "WinSCP"

# --- Functions ---


function Get-LatestWinSCPVersion {
    param([switch]$Quiet)

    $url = "https://winscp.net/eng/downloads.php"
    Write-Log "WinSCP downloads page        : $url" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error (ConvertTo-SafeCurlUrl $url)) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP downloads page: $url" }

        $version = $null
        if ($html -match 'Download\s+WinSCP\s+([0-9]+\.[0-9]+(?:\.[0-9]+)?)') {
            $version = $matches[1]
        }
        elseif ($html -match 'WinSCP-([0-9]+\.[0-9]+(?:\.[0-9]+)?)-Setup\.exe') {
            $version = $matches[1]
        }

        if (-not $version) {
            throw "Could not parse latest WinSCP version from downloads page."
        }

        Write-Log "Latest WinSCP version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get WinSCP version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-WinSCPDownloadUrl {
    param([Parameter(Mandatory)][string]$Version)

    # SourceForge hosts WinSCP releases; URL redirects to a mirror
    $url = "https://sourceforge.net/projects/winscp/files/WinSCP/$Version/WinSCP-$Version-Setup.exe/download"
    Write-Log "SourceForge download URL     : $url"
    return $url
}

function Test-DownloadedInstaller {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -lt 1MB) { return $false }

    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 64
        [void]$fs.Read($buf, 0, $buf.Length)
        $head = [System.Text.Encoding]::ASCII.GetString($buf)
        if ($head -match '<!DOCTYPE|<html|<HTML') { return $false }
    }
    finally { $fs.Dispose() }

    return $true
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWinSCP {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinSCP (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestWinSCPVersion
    if (-not $version) { throw "Could not resolve WinSCP version." }

    if ($version -notmatch '^\d+\.\d+(\.\d+)?$') {
        throw "Parsed version '$version' does not look like x.y or x.y.z"
    }

    $installerFileName = "WinSCP-$version-Setup.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download from SourceForge ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        $downloadUrl = Get-WinSCPDownloadUrl -Version $version

        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe -ExtraCurlArgs @('-L', '--max-redirs', '10')

        if (-not (Test-DownloadedInstaller -Path $localExe)) {
            try { Remove-Item -LiteralPath $localExe -Force -ErrorAction SilentlyContinue } catch {}
            throw "Downloaded file did not validate as installer (too small or HTML content)."
        }

        Write-Log ("Download validated             ({0} bytes)" -f (Get-Item -LiteralPath $localExe).Length)
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

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
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/VERYSILENT'', ''/NORESTART'', ''/ALLUSERS'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        '$uninstaller = Join-Path $env:ProgramFiles ''WinSCP\unins000.exe''',
        'if (-not (Test-Path -LiteralPath $uninstaller)) {',
        '    $uninstaller = Join-Path ${env:ProgramFiles(x86)} ''WinSCP\unins000.exe''',
        '}',
        'if (Test-Path -LiteralPath $uninstaller) {',
        '    $proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/VERYSILENT'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}',
        'Write-Warning ''WinSCP uninstall executable not found.''',
        'exit 0'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # Nothing is installed on this machine: the product name follows the
    # installer's own "WinSCP <version>" display name and detection reads the
    # installed WinSCP.exe version.
    $publisher = "Martin Prikryl"
    $appName   = "WinSCP $version"

    Write-Log "Detection file               : C:\Program Files (x86)\WinSCP\WinSCP.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /ALLUSERS"
        UninstallArgs   = "/VERYSILENT /NORESTART"
        RunningProcess  = @("WinSCP")
        Detection       = @{
            Type          = "File"
            FilePath      = "C:\Program Files (x86)\WinSCP"
            FileName      = "WinSCP.exe"
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

function Invoke-PackageWinSCP {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinSCP (x64) - PACKAGE phase"
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
    Write-Log "Detection RegKey             : $($manifest.Detection.RegistryKeyRelative)"
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

    # --- ConfigMgr application ---
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
        $v = Get-LatestWinSCPVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
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
    Write-Log "WinSCP (x64) Auto-Packager starting"
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
        Invoke-StageWinSCP
    }
    elseif ($PackageOnly) {
        Invoke-PackageWinSCP
    }
    else {
        Invoke-StageWinSCP
        Invoke-PackageWinSCP
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-winscp'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
