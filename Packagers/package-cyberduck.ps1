<#
Vendor: iterate GmbH
App: Cyberduck
CMName: Cyberduck
VendorUrl: https://cyberduck.io/
CPE: cpe:2.3:a:cyberduck:cyberduck:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://cyberduck.io/changelog/
DownloadPageUrl: https://cyberduck.io/download/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Cyberduck (x64) for MECM.

.DESCRIPTION
    Resolves the latest Cyberduck release from the vendor Sparkle changelog feed,
    downloads the Windows installer, stages content to a versioned local folder,
    and creates an MECM Application with file-existence detection.

    The installer is a WiX burn bundle installed silently with /quiet. Bonjour is
    suppressed because it installs a separate machine-wide service that most
    managed environments do not want.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Cyberduck version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$ChangelogFeedUrl = "https://version.cyberduck.io/windows/changelog.rss"

$VendorFolder = "iterate GmbH"
$AppFolder    = "Cyberduck"

$BaseDownloadRoot = Join-Path $DownloadRoot "Cyberduck"
$InstallPath      = "C:\Program Files\Cyberduck"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The download host answers a missing build with a 302 to an HTML error
        page, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestCyberduckRelease {
    <#
    .SYNOPSIS
        Returns the newest Cyberduck version and its installer URL.
    .DESCRIPTION
        The feed lists every release newest-first, and the installer filename
        combines the marketing version with the build number, so both parts are
        taken from the same enclosure element rather than assembled separately.
        The published URL carries a doubled path separator, which the download
        host accepts but which is normalised here.
    #>
    param([switch]$Quiet)

    Write-Log "Changelog feed URL           : $ChangelogFeedUrl" -Quiet:$Quiet

    try {
        $xml = (curl.exe -L --fail --silent --show-error $ChangelogFeedUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Cyberduck changelog feed." }

        $m = [regex]::Match($xml, '<enclosure[^>]*\burl="(?<url>[^"]*?Cyberduck-Installer-(?<ver>\d+(?:\.\d+)+)\.exe)"')
        if (-not $m.Success) { throw "Could not locate an installer enclosure in the changelog feed." }

        $version = $m.Groups['ver'].Value
        $url     = $m.Groups['url'].Value -replace '(?<!:)//', '/'

        Write-Log "Latest Cyberduck version     : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = "Cyberduck-Installer-$version.exe"
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get Cyberduck version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCyberduck {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Cyberduck (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestCyberduckRelease
    if (-not $releaseInfo) { throw "Could not resolve Cyberduck version." }

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
        Write-Log "Downloading Cyberduck..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/quiet', '/norestart', 'InstallBonjour=0'" `
        -UninstallCommand 'unused'

    # A burn bundle registers under a generated ARP key and removes itself
    # through the cached copy of the bundle, so the uninstall command is read
    # back from the registry rather than assumed. Absent entry means the product
    # is not present, so removal exits clean.
    $customUninstall = (
        'Get-Process Cyberduck -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        'Start-Sleep -Seconds 2',
        '$roots = @(''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'', ''HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'')',
        '$entry = $null',
        'foreach ($root in $roots) {',
        '    if (-not (Test-Path -LiteralPath $root)) { continue }',
        '    $entry = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |',
        '        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |',
        '        Where-Object { $_.DisplayName -like ''Cyberduck*'' -and $_.BundleCachePath } |',
        '        Select-Object -First 1',
        '    if ($entry) { break }',
        '}',
        'if (-not $entry) { exit 0 }',
        'if (-not (Test-Path -LiteralPath $entry.BundleCachePath)) { exit 0 }',
        '$proc = Start-Process -FilePath $entry.BundleCachePath -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : Cyberduck.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Cyberduck"
        Publisher       = "iterate GmbH"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/quiet /norestart InstallBonjour=0"
        UninstallArgs   = "/uninstall /quiet /norestart"
        RunningProcess  = @("Cyberduck")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallPath
            FileName     = "Cyberduck.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
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

function Invoke-PackageCyberduck {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Cyberduck (x64) - PACKAGE phase"
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
        $info = Get-LatestCyberduckRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Cyberduck GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Cyberduck (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ChangelogFeedUrl             : $ChangelogFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCyberduck
    }
    elseif ($PackageOnly) {
        Invoke-PackageCyberduck
    }
    else {
        Invoke-StageCyberduck
        Invoke-PackageCyberduck
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-cyberduck'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
