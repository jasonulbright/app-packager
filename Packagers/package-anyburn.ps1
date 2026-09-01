<#
Vendor: PowerSoft
App: AnyBurn
CMName: AnyBurn
VendorUrl: https://anyburn.com/
CPE: cpe:2.3:a:anyburn:anyburn:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://anyburn.com/tutorials/history.htm
DownloadPageUrl: https://anyburn.com/download.htm
UpdateCadenceDays: 120

.SYNOPSIS
    Packages AnyBurn (free edition, x64 host) for MECM.

.DESCRIPTION
    Reads the current version from the vendor home page, downloads the free
    edition installer, stages content to a versioned local folder, and creates
    an MECM Application with script-based detection against the product's
    Add/Remove Programs entry.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The installer filename carries no version, so the staged copy is named for
    the resolved version to keep versioned content folders self-describing.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\PowerSoft\AnyBurn\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AnyBurn).
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
    create MECM application with script-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available AnyBurn version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$HomePageUrl     = "https://anyburn.com/index.htm"
$InstallerUrl    = "https://anyburn.com/anyburn_setup.exe"
$SourceFileName  = "anyburn_setup.exe"

$VendorFolder = "PowerSoft"
$AppFolder    = "AnyBurn"

$BaseDownloadRoot = Join-Path $DownloadRoot "AnyBurn"

# --- Functions ---


function Assert-ExecutablePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor host answers 200 with an HTML error page for paths it no
        longer serves, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($bytes, 0, 2) } finally { $stream.Dispose() }
    if ($read -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestAnyBurnVersion {
    <#
    .SYNOPSIS
        Returns the version advertised on the vendor home page.
    .DESCRIPTION
        The download page links a version-less installer filename, so the home
        page banner ("AnyBurn v6.9") is the only published version marker.
    #>
    param([switch]$Quiet)

    Write-Log "Home page URL                : $HomePageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $HomePageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the AnyBurn home page: $HomePageUrl" }

        $rx = [regex]'(?i)AnyBurn\s*v(?<ver>\d+\.\d+(?:\.\d+)?)'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate a version banner on the home page."
        }

        $best = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        Write-Log "Latest AnyBurn version       : $best" -Quiet:$Quiet
        return $best
    }
    catch {
        Write-Log "Failed to get AnyBurn version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAnyBurn {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyBurn - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestAnyBurnVersion
    if (-not $version) { throw "Could not resolve the AnyBurn version." }

    $installerFileName = "anyburn_setup_$version.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $InstallerUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExecutablePayload -Path $localExe

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
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand 'unused'

    # The NSIS uninstaller lives under the install directory the installer
    # chose, so the ARP UninstallString is the only value that names it.
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
        Where-Object { $_.DisplayName -like 'AnyBurn*' } |
        Select-Object -First 1
    if ($hit -and $hit.UninstallString) { $cmd = $hit.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # The vendor publishes no install path and the installer picks between a
    # 32-bit and a 64-bit build at run time, so detection reads the ARP entry
    # instead of hard-coding a Program Files location.
    $detectionScript = @"
`$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
`$wanted = [version]'$version'
foreach (`$root in `$keys) {
    if (-not (Test-Path -LiteralPath `$root)) { continue }
    `$hit = Get-ChildItem -LiteralPath `$root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath `$_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { `$_.DisplayName -like 'AnyBurn*' } |
        Select-Object -First 1
    if (-not `$hit) { continue }
    `$found = `$null
    if (-not [version]::TryParse([string]`$hit.DisplayVersion, [ref]`$found)) { continue }
    if (`$found -ge `$wanted) { Write-Output 'Installed'; exit 0 }
}
exit 0
"@

    Write-Log ""
    Write-Log "Detection                    : ARP entry 'AnyBurn*' with DisplayVersion >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "AnyBurn"
        Publisher       = "PowerSoft"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("AnyBurn")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
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

function Invoke-PackageAnyBurn {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyBurn - PACKAGE phase"
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
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
        }
    }

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
        $v = Get-LatestAnyBurnVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("AnyBurn GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AnyBurn Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "HomePageUrl                  : $HomePageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAnyBurn
    }
    elseif ($PackageOnly) {
        Invoke-PackageAnyBurn
    }
    else {
        Invoke-StageAnyBurn
        Invoke-PackageAnyBurn
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-anyburn'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
