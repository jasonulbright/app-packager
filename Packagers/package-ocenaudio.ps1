<#
Vendor: Ocenaudio Team
App: ocenaudio
CMName: ocenaudio
VendorUrl: https://www.ocenaudio.com/
CPE: cpe:2.3:a:ocenaudio:ocenaudio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.ocenaudio.com/whatsnew
DownloadPageUrl: https://www.ocenaudio.com/download
UpdateCadenceDays: 90

.SYNOPSIS
    Packages ocenaudio (x64, all-users) for MECM.

.DESCRIPTION
    Reads the published version from the vendor file-information page, downloads
    the Windows x64 NSIS installer, stages content to a versioned local folder,
    and creates an MECM Application with script-based ARP detection.

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
    Outputs only the latest available ocenaudio version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
# /start_download serves a meta-refresh interstitial; /downloads is the file it
# refreshes to, and /fileinfo is the only page that states the version.
$FileInfoUrl = "https://www.ocenaudio.com/fileinfo/ocenaudio_windows64.exe"
$DownloadUrl = "https://www.ocenaudio.com/downloads/ocenaudio_windows64.exe"

$VendorFolder = "Ocenaudio Team"
$AppFolder    = "ocenaudio"

$BaseDownloadRoot  = Join-Path $DownloadRoot "ocenaudio"
$InstallerFileName = "ocenaudio_windows64.exe"
$UninstallKeyName  = "ocenaudio"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor answers 200 with an HTML interstitial for some download
        paths, which would otherwise stage as a valid-looking installer and
        fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestOcenaudioVersion {
    param([switch]$Quiet)

    Write-Log "ocenaudio file info page     : $FileInfoUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $FileInfoUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch ocenaudio file info page: $FileInfoUrl" }

        $rx = [regex]'Version:</strong>\s*Version\s*(?<ver>\d+(?:\.\d+){1,3})'
        $m = $rx.Match($html)
        if (-not $m.Success) {
            throw "Could not locate a version on the file info page."
        }

        $version = $m.Groups['ver'].Value
        Write-Log "Latest ocenaudio version     : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get ocenaudio version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageOcenaudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "ocenaudio (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $version = Get-LatestOcenaudioVersion
    if (-not $version) { throw "Could not resolve ocenaudio version." }

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $DownloadUrl"
    Write-Log ""

    # --- Download ---
    # The vendor publishes one unversioned filename per platform, so the cached
    # copy is re-fetched every run rather than trusted to match $version.
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Downloading ocenaudio..."
    Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe

    Assert-ExePayload -Path $localExe

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -notlike "$version*") {
        throw "Downloaded installer reports version '$fileVersion' but the file info page announced '$version'; refusing to stage a mismatched payload."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # The installer places the uninstaller next to the payload under a target
    # the operator can change, so uninstall resolves UninstallString from the
    # ARP entry instead of assuming a Program Files path.
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/AllUsers', '/S') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallScript = @"
`$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$UninstallKeyName',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$UninstallKeyName'
)
foreach (`$key in `$keys) {
    if (-not (Test-Path -LiteralPath `$key)) { continue }
    `$entry = Get-ItemProperty -LiteralPath `$key -ErrorAction SilentlyContinue
    if (-not `$entry -or -not `$entry.UninstallString) { continue }
    # UninstallString is '"<dir>\uninst.exe" /AllUsers'; only the quoted
    # executable is wanted so the context switch can be re-supplied here.
    `$exe = (`$entry.UninstallString -replace '^"([^"]+)".*`$', '`$1').Trim('"')
    if (-not (Test-Path -LiteralPath `$exe)) { continue }
    `$proc = Start-Process -FilePath `$exe -ArgumentList @('/AllUsers', '/S') -Wait -PassThru -NoNewWindow
    exit `$proc.ExitCode
}
Write-Error 'ocenaudio uninstall entry not found.'
exit 1
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # The installer offers an operator-selectable target directory and the
    # 32-bit NSIS stub's registry view is not fixed, so detection reads the
    # fixed ARP key name from both views rather than a file path.
    $detectionScript = @"
`$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$UninstallKeyName',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$UninstallKeyName'
)
`$wanted = [version]'$version'
foreach (`$key in `$keys) {
    if (-not (Test-Path -LiteralPath `$key)) { continue }
    `$entry = Get-ItemProperty -LiteralPath `$key -ErrorAction SilentlyContinue
    if (-not `$entry) { continue }
    `$found = `$null
    if (-not [version]::TryParse([string]`$entry.DisplayVersion, [ref]`$found)) { continue }
    if (`$found -ge `$wanted) { Write-Output 'Installed'; exit 0 }
}
exit 0
"@

    Write-Log ""
    Write-Log "Detection                    : ARP key '$UninstallKeyName' with DisplayVersion >= $version"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "ocenaudio $version"
        Publisher       = "Ocenaudio Team"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/AllUsers /S"
        UninstallArgs   = "/AllUsers /S"
        RunningProcess  = @("ocenaudio")
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

function Invoke-PackageOcenaudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "ocenaudio (x64) - PACKAGE phase"
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
    Write-Log "Detection type               : $($manifest.Detection.Type)"
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
        $version = Get-LatestOcenaudioVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("ocenaudio GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "ocenaudio (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadUrl                  : $DownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageOcenaudio
    }
    elseif ($PackageOnly) {
        Invoke-PackageOcenaudio
    }
    else {
        Invoke-StageOcenaudio
        Invoke-PackageOcenaudio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-ocenaudio'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
