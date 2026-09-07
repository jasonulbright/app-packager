<#
Vendor: TeamViewer
App: TeamViewer Host (x64)
CMName: TeamViewer Host
VendorUrl: https://www.teamviewer.com/
CPE: cpe:2.3:a:teamviewer:teamviewer:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.teamviewer.com/en-us/whats-new/
DownloadPageUrl: https://www.teamviewer.com/en-us/download/windows/
IconSource: Installer

.SYNOPSIS
    Packages TeamViewer Host (x64) EXE for MECM.

.DESCRIPTION
    Downloads the latest TeamViewer Host x64 setup EXE from TeamViewer's static
    download URL, stages content to a versioned local folder with file-version
    detection metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download EXE, derive version from file properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    TeamViewer Host is the unattended-access variant of TeamViewer. It is
    deployed to endpoints that need to be remotely managed by IT without a
    user initiating the session. The Full client (package-teamviewer.ps1) is
    deployed to IT staff who initiate support sessions.

    NOTE: The vendor's MSI-for-Mass-Deployment URL
    (https://download.teamviewer.com/download/version_15x/TeamViewer_Host.msi)
    started returning HTTP 404 in April 2026. The MSI variant is only offered
    to authenticated Management Console users now. TeamViewer also churns
    these URLs roughly yearly -- re-verify $ExeDownloadUrl whenever the
    packager starts failing on -GetLatestVersionOnly.

    This packager consequently uses the NSIS EXE installer (/S silent flag).
    The EXE path accepts fewer unattended-enrollment parameters than the MSI
    did: APITOKEN and ASSIGNMENTOPTIONS are MSI-only properties and cannot be
    passed. CUSTOMCONFIGID, RemoveDesktopShortcut, and similar tenant-specific
    configuration values are likewise not honored by the EXE. If
    teamviewer-host-config.json is present, the values are logged as WARNINGS
    so operators know they were ignored, and the schema stays compatible so a
    future MSI-aware packager can consume the same file.

    GetLatestVersionOnly downloads the EXE, reads the file version, and exits.
    No lighter-weight version API is available.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\TeamViewer\TeamViewer Host\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TeamViewerHost).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download EXE, derive version from file
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Downloads the TeamViewer Host EXE, reads the file version, outputs the
    version string, and exits. No MECM changes are made.

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
# TeamViewer Host NSIS installer. /S flag is silent. APITOKEN /
# CUSTOMCONFIGID / ASSIGNMENTOPTIONS properties do NOT apply to the EXE -
# those are MSI-only (Management Console login required for the MSI).
$ExeDownloadUrl    = "https://dl.teamviewer.com/download/TeamViewer_Host_Setup_x64.exe"
$InstallerFileName = "TeamViewer_Host_Setup_x64.exe"

$VendorFolder = "TeamViewer"
$AppFolder    = "TeamViewer Host"

$BaseDownloadRoot = Join-Path $DownloadRoot "TeamViewerHost"
$ConfigFile       = Join-Path $PSScriptRoot "teamviewer-host-config.json"

# --- Functions ---


function Get-TeamViewerHostExeVersion {
    <#
    .SYNOPSIS
        Reads the product version from a TeamViewer Host EXE and returns it as
        a dotted version string.

    .DESCRIPTION
        The NSIS installer leaves the version resource string table empty, so
        FileVersionInfo.ProductVersion and FileVersion come back blank. The
        fixed block of the same resource is populated, and its four Product
        parts match the vendor MSI ProductVersion.

        Strategy:
          1. Try the string table first, in case a vendor build populates it.
          2. Fall back to the fixed block.
          3. Throw when both are empty.
    #>
    param([Parameter(Mandatory)][string]$ExePath)

    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
    $v  = $vi.ProductVersion
    if ([string]::IsNullOrWhiteSpace($v)) { $v = $vi.FileVersion }
    if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }

    $fixed = "{0}.{1}.{2}.{3}" -f $vi.ProductMajorPart, $vi.ProductMinorPart, $vi.ProductBuildPart, $vi.ProductPrivatePart
    if ($fixed -ne '0.0.0.0') { return $fixed }

    throw "Could not read a product version from $ExePath - the version resource string table and its fixed block are both empty."
}


function Write-IgnoredConfigWarnings {
    <#
    .SYNOPSIS
        Logs WARN lines for teamviewer-host-config.json keys that applied
        only to the MSI path and are ignored by the EXE installer.
    #>
    if (-not (Test-Path -LiteralPath $ConfigFile)) { return }

    try {
        $cfg = Get-Content -LiteralPath $ConfigFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Log "teamviewer-host-config.json unreadable: $($_.Exception.Message)" -Level WARN
        return
    }
    if (-not $cfg) { return }

    $ignored = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.ApiToken))          { $ignored += 'ApiToken (APITOKEN)' }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.CustomConfigId))    { $ignored += 'CustomConfigId (CUSTOMCONFIGID)' }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.AssignmentOptions)) { $ignored += 'AssignmentOptions (ASSIGNMENTOPTIONS)' }
    if ($cfg.RemoveDesktopShortcut -eq $true)                              { $ignored += 'RemoveDesktopShortcut (REMOVE=f.DesktopShortcut)' }

    if ($ignored.Count -gt 0) {
        Write-Log "teamviewer-host-config.json contains MSI-only settings that the EXE installer cannot accept:" -Level WARN
        foreach ($k in $ignored) {
            Write-Log "  - $k" -Level WARN
        }
        Write-Log "These values are being ignored. Enrol affected endpoints from the TeamViewer Management Console after install." -Level WARN
    }
}


function Get-LatestTeamViewerHostVersion {
    param([switch]$Quiet)

    Write-Log "TeamViewer Host EXE URL      : $ExeDownloadUrl" -Quiet:$Quiet

    try {
        Initialize-Folder -Path $BaseDownloadRoot

        $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
        Invoke-CachedDownload -Url $ExeDownloadUrl -OutFile $localExe -Quiet:$Quiet

        $version = Get-TeamViewerHostExeVersion -ExePath $localExe
        Write-Log "Latest TeamViewer Host ver   : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get TeamViewer Host version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeamViewerHost {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download EXE ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local EXE path               : $localExe"
    Invoke-CachedDownload -Url $ExeDownloadUrl -OutFile $localExe

    # --- Read version from EXE file properties ---
    $version = Get-TeamViewerHostExeVersion -ExePath $localExe
    Write-Log "EXE ProductVersion           : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Log MSI-only config that is being ignored on the EXE path ---
    Write-IgnoredConfigWarnings

    # --- Generate content wrappers ---
    # Install: stop any running TeamViewer process before the NSIS silent
    # installer runs (it can't replace an open exe), then run /S, then
    # post-install kill TeamViewer.exe so a SYSTEM-session GUI doesn't sit
    # on the user's desktop (TV post-install occasionally spawns one under
    # the MECM install context).
    $installContent = (
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall: NSIS uninstaller at standard Host install path
    $uninstallContent = (
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        '$uninstall = Join-Path $env:ProgramFiles ''TeamViewer\uninstall.exe''',
        'if (-not (Test-Path -LiteralPath $uninstall)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstall -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "TeamViewer Host $version"
    $publisher = "TeamViewer"

    $detectionPath = "{0}\TeamViewer" -f $env:ProgramFiles
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : TeamViewer.exe"
    Write-Log "Detection version (>=)       : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("TeamViewer_Service", "TeamViewer")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "TeamViewer.exe"
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

function Invoke-PackageTeamViewerHost {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - PACKAGE phase"
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
        $v = Get-LatestTeamViewerHostVersion -Quiet
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
    Write-Log "TeamViewer Host (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ExeDownloadUrl               : $ExeDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTeamViewerHost
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeamViewerHost
    }
    else {
        Invoke-StageTeamViewerHost
        Invoke-PackageTeamViewerHost
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-teamviewerhost'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
