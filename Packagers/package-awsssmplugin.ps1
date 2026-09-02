<#
Vendor: Amazon Web Services
App: Session Manager Plugin
CMName: AWS Session Manager Plugin
VendorUrl: https://docs.aws.amazon.com/systems-manager/latest/userguide/session-manager-working-with-install-plugin.html
CPE: cpe:2.3:a:amazon:session_manager_plugin:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.aws.amazon.com/systems-manager/latest/userguide/plugin-version-history.html
DownloadPageUrl: https://docs.aws.amazon.com/systems-manager/latest/userguide/install-plugin-windows.html
IconSource: None
UpdateCadenceDays: 60

.SYNOPSIS
    Packages the AWS Session Manager Plugin for MECM.

.DESCRIPTION
    Reads the current version from the vendor distribution bucket, downloads the
    matching WiX bundle installer, stages content to a versioned local folder,
    and creates an MECM Application with file-existence detection on the
    installed session-manager-plugin.exe.

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
    Content is staged under: <FileServerPath>\Applications\Amazon Web Services\Session Manager Plugin\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Session Manager Plugin).
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
    Outputs only the latest available Session Manager Plugin version string and exits.

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
$DistributionRoot = "https://s3.amazonaws.com/session-manager-downloads/plugin"
$VersionUrl       = "$DistributionRoot/latest/VERSION"

$VendorFolder = "Amazon Web Services"
$AppFolder    = "Session Manager Plugin"

$BaseDownloadRoot  = Join-Path $DownloadRoot "Session Manager Plugin"
$InstallerFileName = "SessionManagerPluginSetup.exe"

$InstallDir = "{0}\Amazon\SessionManagerPlugin\bin" -f $env:ProgramFiles

# --- Functions ---


function Assert-SsmPluginPayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The distribution bucket answers 200 with an XML error body for keys that
        no longer resolve, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) }
    finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestSsmPluginVersion {
    <#
    .SYNOPSIS
        Returns the version string published alongside the current plugin build.
    #>
    param([switch]$Quiet)

    Write-Log "Version URL                  : $VersionUrl" -Quiet:$Quiet

    try {
        $raw = (curl.exe -L --fail --silent --show-error $VersionUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch version marker: $VersionUrl" }

        $version = $raw.Trim()
        if ($version -notmatch '^\d+(\.\d+){1,3}$') {
            throw "Version marker did not contain a version string: '$version'"
        }

        Write-Log "Latest plugin version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Session Manager Plugin version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSsmPlugin {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Session Manager Plugin - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestSsmPluginVersion
    if (-not $version) { throw "Could not resolve Session Manager Plugin version." }

    # The versioned key is pinned rather than /latest/ so the staged payload and
    # the recorded version cannot drift apart mid-run.
    $downloadUrl = "$DistributionRoot/$version/windows/$InstallerFileName"

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    # --- Download ---
    $localExe = Join-Path $localContentPath $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-SsmPluginPayloadIsExecutable -Path $localExe

    # --- Generate content wrappers ---
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs "'/install', '/quiet', '/norestart'" `
        -UninstallCommand 'unused'

    # The bundle caches itself under a per-build GUID folder, so the ARP
    # QuietUninstallString is the only value that names the right copy.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
$hit = $null
foreach ($root in $keys) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $hit = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.DisplayName -like 'Session Manager Plugin*' } |
        Select-Object -First 1
    if ($hit) { break }
}
if (-not $hit) { exit 0 }
$cmd = if ($hit.QuietUninstallString) { $hit.QuietUninstallString } else { $hit.UninstallString }
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = ($cmd -split '\s+/')[0].Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/uninstall', '/quiet', '/norestart') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : session-manager-plugin.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "AWS Session Manager Plugin $version"
        Publisher       = "Amazon Web Services"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/install /quiet /norestart"
        UninstallArgs   = "/uninstall /quiet /norestart"
        RunningProcess  = @("session-manager-plugin")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallDir
            FileName     = "session-manager-plugin.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
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

function Invoke-PackageSsmPlugin {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Session Manager Plugin - PACKAGE phase"
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

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
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
        $v = Get-LatestSsmPluginVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Session Manager Plugin GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AWS Session Manager Plugin Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionUrl                   : $VersionUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSsmPlugin
    }
    elseif ($PackageOnly) {
        Invoke-PackageSsmPlugin
    }
    else {
        Invoke-StageSsmPlugin
        Invoke-PackageSsmPlugin
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-awsssmplugin'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
