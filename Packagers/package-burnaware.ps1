<#
Vendor: Burnaware
App: BurnAware Free
CMName: BurnAware Free
VendorUrl: https://www.burnaware.com/
CPE: cpe:2.3:a:burnaware:burnaware:*:*:*:*:free:*:*:*
ReleaseNotesUrl: https://www.burnaware.com/history.html
DownloadPageUrl: https://www.burnaware.com/download.html
IconSource: Installer
UpdateCadenceDays: 180

.SYNOPSIS
    Packages BurnAware Free for MECM.

.DESCRIPTION
    Reads the current version from the vendor download page, downloads the
    matching Inno Setup installer, stages content to a versioned local folder,
    and creates an MECM Application with registry-based detection on the
    Inno Setup ARP key.

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
    Content is staged under: <FileServerPath>\Applications\Burnaware\BurnAware Free\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\BurnAware Free).
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
    Outputs only the latest available BurnAware Free version string and exits.

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
$DownloadPageUrl = "https://www.burnaware.com/download.html"
$DownloadBase    = "https://www.burnaware.com/downloads"

$VendorFolder = "Burnaware"
$AppFolder    = "BurnAware Free"

$BaseDownloadRoot = Join-Path $DownloadRoot "BurnAware Free"

# The Free edition ships a 32-bit payload only, so Setup runs in 32-bit mode
# and its ARP key lands in the WOW6432Node view on 64-bit Windows.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BurnAware Free_is1"

# --- Functions ---


function Assert-BurnAwarePayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor host answers 200 with an HTML page for download paths that
        do not exist (for example an x64 Free build), which would otherwise
        stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestBurnAwareVersion {
    <#
    .SYNOPSIS
        Returns the highest BurnAware Free version linked on the download page.
    .DESCRIPTION
        The page links the Free, Premium and Pro editions together; only the
        Free installer filename is matched so a Premium-only bump cannot
        advance the staged version.
    #>
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch BurnAware download page: $DownloadPageUrl" }

        $rx = [regex]'burnaware_free_(?<ver>\d+(?:\.\d+)+)\.exe'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any BurnAware Free installer links on the download page."
        }

        $best = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        if ([string]::IsNullOrWhiteSpace($best)) { throw "Version selection produced an empty result." }

        Write-Log "Latest BurnAware Free version: $best" -Quiet:$Quiet
        return $best
    }
    catch {
        Write-Log "Failed to get BurnAware version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBurnAware {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BurnAware Free - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestBurnAwareVersion
    if (-not $version) { throw "Could not resolve BurnAware Free version." }

    $installerFileName = "burnaware_free_$version.exe"
    $downloadUrl       = "$DownloadBase/$installerFileName"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-BurnAwarePayloadIsExecutable -Path $localExe

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
        -InstallArgs "'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES'" `
        -UninstallCommand 'unused'

    # Inno Setup names the uninstaller unins###.exe by install order, so the
    # ARP UninstallString is the only value that names the right one.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BurnAware Free_is1',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\BurnAware Free_is1'
)
$cmd = $null
foreach ($key in $keys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $hit = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($hit -and $hit.UninstallString) { $cmd = $hit.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "BurnAware Free"
        Publisher       = "Burnaware"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES"
        UninstallArgs   = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES"
        RunningProcess  = @("BurnAware")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $false
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

function Invoke-PackageBurnAware {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BurnAware Free - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
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
        $v = Get-LatestBurnAwareVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("BurnAware GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "BurnAware Free Auto-Packager starting"
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
        Invoke-StageBurnAware
    }
    elseif ($PackageOnly) {
        Invoke-PackageBurnAware
    }
    else {
        Invoke-StageBurnAware
        Invoke-PackageBurnAware
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-burnaware'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
