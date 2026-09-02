<#
Vendor: JAM Software
App: TreeSize Free
CMName: TreeSize Free
VendorUrl: https://www.jam-software.com/treesize_free
CPE: cpe:2.3:a:jam-software:treesize_free:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.jam-software.com/treesize_free/changes.shtml
DownloadPageUrl: https://www.jam-software.com/treesize_free
UpdateCadenceDays: 90

.SYNOPSIS
    Packages TreeSize Free (x64) for MECM.

.DESCRIPTION
    Reads the current version from the vendor change log, downloads the
    TreeSize Free setup from the vendor download endpoint, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection on the installed TreeSizeFree.exe.

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
    Content is staged under: <FileServerPath>\Applications\JAM Software\TreeSize Free\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TreeSizeFree).
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
    Outputs only the latest available TreeSize Free version string and exits.

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
$ChangeLogUrl = "https://www.jam-software.com/treesize_free/changes.shtml"
$DownloadUrl  = "https://www.jam-software.com/treesize_free/TreeSizeFreeSetup.exe"

$VendorFolder = "JAM Software"
$AppFolder    = "TreeSize Free"

$BaseDownloadRoot  = Join-Path $DownloadRoot "TreeSizeFree"
$InstallerFileName = "TreeSizeFreeSetup.exe"

$InstallDir = "{0}\JAM Software\TreeSize Free" -f $env:ProgramFiles

# --- Functions ---


function Get-LatestTreeSizeFreeVersion {
    <#
    .SYNOPSIS
        Returns the highest version heading in the vendor change log.
    .DESCRIPTION
        The change log lists every historical release, so the highest
        [version] wins rather than the first heading in document order.
    #>
    param([switch]$Quiet)

    Write-Log "Change log URL               : $ChangeLogUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $ChangeLogUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch TreeSize Free change log: $ChangeLogUrl" }

        $rx = [regex]'Version\s+(?<ver>[\d.]+)</h3>'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any version headings in the change log."
        }

        $best = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value.TrimEnd('.') } |
            Where-Object { $_ -match '^\d+(\.\d+)+$' } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        if ([string]::IsNullOrWhiteSpace($best)) { throw "Version selection produced an empty result." }

        Write-Log "Latest TreeSize Free version : $best" -Quiet:$Quiet
        return $best
    }
    catch {
        Write-Log "Failed to get TreeSize Free version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTreeSizeFree {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TreeSize Free (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestTreeSizeFreeVersion
    if (-not $version) { throw "Could not resolve TreeSize Free version." }

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Download ---
    # The vendor download endpoint is version-less and always serves current,
    # so the payload is fetched per staged version rather than cached by name.
    $localExe = Join-Path $BaseDownloadRoot ("{0}-{1}" -f $version, $InstallerFileName)
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $DownloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # The vendor download endpoint answers 200 with an HTML page rather than
    # a redirect when the setup is unavailable.
    $header = Get-Content -LiteralPath $localExe -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($header.Count -lt 2 -or $header[0] -ne 0x4D -or $header[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $localExe"
    }

    $binaryVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if (-not [string]::IsNullOrWhiteSpace($binaryVersion)) {
        Write-Log "Installer FileVersion        : $binaryVersion"
    }

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

    # --- Generate content wrappers ---
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs "'/VERYSILENT', '/NORESTART', '/MERGETASKS=!desktopicon'" `
        -UninstallCommand 'unused'

    # The setup names its uninstaller by install order, so the ARP
    # UninstallString is the only value that names the right one.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
$cmd = $null
foreach ($root in $keys) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $hit = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.DisplayName -like 'TreeSize Free*' } |
        Select-Object -First 1
    if ($hit -and $hit.UninstallString) { $cmd = $hit.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/NORESTART') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "TreeSize Free (x64)"
    $publisher = "JAM Software"

    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : TreeSizeFree.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /MERGETASKS=!desktopicon"
        UninstallArgs   = "/VERYSILENT /NORESTART"
        RunningProcess  = @("TreeSizeFree")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "TreeSizeFree.exe"
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

function Invoke-PackageTreeSizeFree {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TreeSize Free (x64) - PACKAGE phase"
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
        $v = Get-LatestTreeSizeFreeVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("TreeSize Free GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TreeSize Free (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ChangeLogUrl                 : $ChangeLogUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTreeSizeFree
    }
    elseif ($PackageOnly) {
        Invoke-PackageTreeSizeFree
    }
    else {
        Invoke-StageTreeSizeFree
        Invoke-PackageTreeSizeFree
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-treesizefree'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
