<#
Vendor: NGWIN
App: PicPick
CMName: PicPick
VendorUrl: https://picpick.app/
CPE: cpe:2.3:a:ngwin:picpick:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://picpick.app/en/changelog/
DownloadPageUrl: https://picpick.app/en/download/free/
UpdateCadenceDays: 60

.SYNOPSIS
    Packages PicPick (x86) for MECM.

.DESCRIPTION
    Resolves the newest PicPick build from the vendor's free download page,
    downloads the machine-scope installer, stages content to a versioned local
    folder, and creates an MECM Application with script-based detection on the
    ARP entry.

    The free edition is licensed for personal use only; commercial and
    organizational use requires a purchased PicPick Pro license. The packaged
    payload is the vendor's free installer, so a license entitlement must be
    in place before deploying it in a business environment.

    The vendor ships a 32-bit installer only; the deployment is per-machine
    and its ARP entry therefore lands in the 32-bit registry view on x64.

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
    Content is staged under: <FileServerPath>\Applications\NGWIN\PicPick\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\PicPick).
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
    Outputs only the latest available PicPick version string and exits.

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
$DownloadPageUrl = "https://picpick.app/en/download/free/"

$VendorFolder = "NGWIN"
$AppFolder    = "PicPick"

$BaseDownloadRoot = Join-Path $DownloadRoot "PicPick"

$InstallerFileName = "picpick_inst.exe"

$UninstallKeyName = "PicPick"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The download host answers a missing build path with a 200 HTML page,
        which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestPicPickRelease {
    <#
    .SYNOPSIS
        Returns the newest PicPick version and its installer URL.
    .DESCRIPTION
        The free download page links the installer under a version-numbered
        path on the vendor's download host. The page also links a portable zip
        and the Pro build, so links are matched on the installer filename and
        sorted as versions rather than trusted in document order.
    #>
    param([switch]$Quiet)

    Write-Log "PicPick download page        : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch PicPick download page: $DownloadPageUrl" }

        $rx = [regex]'https://download\.picpick\.net/(?<ver>\d+(\.\d+)+)/picpick_inst\.exe'
        $found = $rx.Matches($html) |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique

        if (-not $found -or @($found).Count -lt 1) {
            throw "Could not locate a PicPick installer link on the download page."
        }

        $best = @($found) | Sort-Object { [version]$_ } | Select-Object -Last 1

        Write-Log "Latest PicPick version       : $best" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best
            FileName    = $InstallerFileName
            DownloadUrl = "https://download.picpick.net/$best/$InstallerFileName"
        }
    }
    catch {
        Write-Log "Failed to get PicPick version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePicPick {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PicPick (x86) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestPicPickRelease
    if (-not $releaseInfo) { throw "Could not resolve PicPick version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    # The vendor publishes one unversioned filename per build path, so a cached
    # copy from an earlier release would shadow the new one; it is re-fetched.
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -notlike "$version*") {
        throw "Downloaded installer reports version '$fileVersion' but the download page announced '$version'; refusing to stage a mismatched payload."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # The NSIS payload writes its uninstaller next to an operator-selectable
    # target, so uninstall resolves UninstallString from the ARP entry instead
    # of assuming a Program Files path.
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
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
    # UninstallString is the uninstaller path, quoted or bare; only the
    # executable is wanted so the silent switch can be re-supplied here.
    `$exe = (`$entry.UninstallString -replace '^"([^"]+)".*`$', '`$1').Trim('"')
    if (-not (Test-Path -LiteralPath `$exe)) { continue }
    `$proc = Start-Process -FilePath `$exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
    exit `$proc.ExitCode
}
Write-Error 'PicPick uninstall entry not found.'
exit 1
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # The 32-bit installer's ARP entry lands in the 32-bit view on x64 clients
    # and the native view on x86, so both are read.
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
        AppName         = "PicPick"
        Publisher       = "NGWIN"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("picpick")
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

function Invoke-PackagePicPick {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PicPick (x86) - PACKAGE phase"
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
        $info = Get-LatestPicPickRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("PicPick GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PicPick (x86) Auto-Packager starting"
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
        Invoke-StagePicPick
    }
    elseif ($PackageOnly) {
        Invoke-PackagePicPick
    }
    else {
        Invoke-StagePicPick
        Invoke-PackagePicPick
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-picpick'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
