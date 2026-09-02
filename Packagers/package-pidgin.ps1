<#
Vendor: Pidgin
App: Pidgin
CMName: Pidgin
VendorUrl: https://pidgin.im/
CPE: cpe:2.3:a:pidgin:pidgin:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://keep.imfreedom.org/pidgin/pidgin/file/tip/ChangeLog
DownloadPageUrl: https://sourceforge.net/projects/pidgin/files/Pidgin/
UpdateCadenceDays: 120

.SYNOPSIS
    Packages Pidgin (x86) for MECM.

.DESCRIPTION
    Resolves the newest Pidgin release from the vendor's SourceForge file
    feed, downloads the offline installer, stages content to a versioned
    local folder, and creates an MECM Application with script-based detection
    on the ARP entry.

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
    Content is staged under: <FileServerPath>\Applications\Pidgin\Pidgin\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Pidgin).
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
    Outputs only the latest available Pidgin version string and exits.

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
$FileFeedUrl = "https://sourceforge.net/projects/pidgin/rss?path=/Pidgin"

$VendorFolder = "Pidgin"
$AppFolder    = "Pidgin"

$BaseDownloadRoot = Join-Path $DownloadRoot "Pidgin"

$UninstallKeyName = "Pidgin"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        SourceForge answers a refused mirror request with a 200 and a two-byte
        'no' body, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestPidginRelease {
    <#
    .SYNOPSIS
        Returns the newest Pidgin release version and its offline installer URL.
    .DESCRIPTION
        The feed lists every file of every release, so entries are filtered to
        the offline installer shape and sorted as versions rather than trusted
        in feed order. The online installer downloads its dependencies at run
        time and is unusable on a locked-down client, so only the offline
        variant is considered.
    #>
    param([switch]$Quiet)

    Write-Log "SourceForge file feed        : $FileFeedUrl" -Quiet:$Quiet

    try {
        # No -A: SourceForge refuses browser user agents on scripted requests.
        $xml = (curl.exe -L --fail --silent --show-error $FileFeedUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch SourceForge file feed: $FileFeedUrl" }

        $rx = [regex]'pidgin-(?<ver>\d+(\.\d+)+)-offline\.exe'
        $found = $rx.Matches($xml) |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique

        if (-not $found -or @($found).Count -lt 1) {
            throw "Could not locate any Pidgin offline installer entries in the file feed."
        }

        $best = @($found) | Sort-Object { [version]$_ } | Select-Object -Last 1
        $fileName = "pidgin-$best-offline.exe"

        Write-Log "Latest Pidgin version        : $best" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best
            FileName    = $fileName
            DownloadUrl = "https://sourceforge.net/projects/pidgin/files/Pidgin/$best/$fileName/download"
        }
    }
    catch {
        Write-Log "Failed to get Pidgin version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePidgin {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Pidgin (x86) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestPidginRelease
    if (-not $releaseInfo) { throw "Could not resolve Pidgin version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -notlike "$version*") {
        throw "Downloaded installer reports version '$fileVersion' but the file feed announced '$version'; refusing to stage a mismatched payload."
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
    # The NSIS payload writes its uninstaller next to an operator-selectable
    # target, so uninstall resolves UninstallString from the ARP entry instead
    # of assuming a Program Files path.
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$installerFileName'
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
Write-Error 'Pidgin uninstall entry not found.'
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
        AppName         = "Pidgin"
        Publisher       = "Pidgin"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("pidgin")
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

function Invoke-PackagePidgin {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Pidgin (x86) - PACKAGE phase"
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
        $info = Get-LatestPidginRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Pidgin GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Pidgin (x86) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "FileFeedUrl                  : $FileFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePidgin
    }
    elseif ($PackageOnly) {
        Invoke-PackagePidgin
    }
    else {
        Invoke-StagePidgin
        Invoke-PackagePidgin
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-pidgin'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
