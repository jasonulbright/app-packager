<#
Vendor: g10 Code GmbH
App: Gpg4win
CMName: Gpg4win
VendorUrl: https://www.gpg4win.org/
CPE: cpe:2.3:a:gpg4win:gpg4win:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.gpg4win.org/version-history.html
DownloadPageUrl: https://www.gpg4win.org/download.html
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Gpg4win for MECM.

.DESCRIPTION
    Resolves the latest Gpg4win release from the vendor file index, downloads the
    NSIS installer, stages content to a versioned local folder, and creates an
    MECM Application with registry-value detection.

    Gpg4win is the publicly downloadable GnuPG distribution for Windows. It is
    the substitute for GnuPG VS-Desktop, which g10 Code distributes only under a
    contractual agreement and therefore has no unauthenticated download URL to
    automate against.

    The installer is an NSIS package installed silently with /S. Its script
    selects the 64-bit registry view before it writes the ARP uninstall key,
    so detection reads the 64-bit view.

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
    Estimated runtime in minutes. Default: 10

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Gpg4win version string and exits.

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
$FileIndexUrl = "https://files.gpg4win.org/"

$VendorFolder = "g10 Code"
$AppFolder    = "Gpg4win"

$BaseDownloadRoot = Join-Path $DownloadRoot "Gpg4win"

# The ARP key the NSIS installer writes; the script runs SetRegView 64 first,
# so the key sits in the 64-bit view, not under WOW6432Node.
$ArpKeyRelative = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Gpg4win"

# --- Functions ---


function Assert-ExecutablePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A mirror error page answers 200 with HTML, which would otherwise stage as
        a valid-looking installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($bytes, 0, 2) } finally { $stream.Dispose() }
    if ($read -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestGpg4winRelease {
    <#
    .SYNOPSIS
        Returns the newest stable Gpg4win version and its installer URL.
    .DESCRIPTION
        The file index holds every release ever published plus beta builds, the
        separate gpg4win-light product, and a floating gpg4win-latest.exe. The
        filename pattern is anchored to three numeric components so only stable
        full releases match, and candidates are ordered as versions rather than
        trusted in index order.
    #>
    param([switch]$Quiet)

    Write-Log "Gpg4win file index           : $FileIndexUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $FileIndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Gpg4win file index." }

        $rx = [regex]'(?<![\w-])gpg4win-(?<ver>\d+\.\d+\.\d+)\.exe(?![\w-])'
        $versions = $rx.Matches($html) | ForEach-Object { $_.Groups['ver'].Value } | Sort-Object -Unique

        if (-not $versions -or @($versions).Count -lt 1) {
            throw "Could not locate any stable Gpg4win installers in the file index."
        }

        $best = @($versions) | Sort-Object { [version]$_ } | Select-Object -Last 1
        $fileName = "gpg4win-$best.exe"

        Write-Log "Latest Gpg4win version       : $best" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best
            FileName    = $fileName
            DownloadUrl = "$FileIndexUrl$fileName"
        }
    }
    catch {
        Write-Log "Failed to get Gpg4win version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGpg4win {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Gpg4win - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestGpg4winRelease
    if (-not $releaseInfo) { throw "Could not resolve Gpg4win version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Gpg4win..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand 'unused'

    # The install directory moved between major versions, so the uninstaller
    # path is read from the ARP UninstallString rather than assumed. The 64-bit
    # view is checked first because that is where the current installer writes;
    # the WOW6432Node key covers older releases.
    $customUninstall = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Gpg4win',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Gpg4win'
)
$cmd = $null
foreach ($k in $keys) {
    if (-not (Test-Path -LiteralPath $k)) { continue }
    $cmd = (Get-ItemProperty -LiteralPath $k -Name UninstallString -ErrorAction SilentlyContinue).UninstallString
    if (-not [string]::IsNullOrWhiteSpace($cmd)) { break }
}
if ([string]::IsNullOrWhiteSpace($cmd)) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection key                : $ArpKeyRelative (64-bit view)"
    Write-Log "Detection DisplayVersion     : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Gpg4win"
        Publisher       = "g10 Code GmbH"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("kleopatra", "gpa", "gpg-agent")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpKeyRelative
            ValueName           = "DisplayVersion"
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

function Invoke-PackageGpg4win {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Gpg4win - PACKAGE phase"
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
        $info = Get-LatestGpg4winRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Gpg4win GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Gpg4win Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "FileIndexUrl                 : $FileIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGpg4win
    }
    elseif ($PackageOnly) {
        Invoke-PackageGpg4win
    }
    else {
        Invoke-StageGpg4win
        Invoke-PackageGpg4win
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-gpg4win'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
