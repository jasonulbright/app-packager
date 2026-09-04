<#
Vendor: Piriform Software Ltd.
App: Defraggler
CMName: Defraggler
VendorUrl: https://www.ccleaner.com/defraggler
CPE: cpe:2.3:a:piriform:defraggler:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.ccleaner.com/defraggler/builds
DownloadPageUrl: https://www.ccleaner.com/defraggler
IconSource: Installer
UpdateCadenceDays: 365

.SYNOPSIS
    Packages Defraggler (x64-capable installer) for MECM.

.DESCRIPTION
    Reads the current build from the vendor builds page and the version from the
    installer file version resource, stages content to a versioned local folder,
    and creates an MECM Application with file-version detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The installer is an NSIS package: /S installs silently for all users and
    lays down both the 32-bit and 64-bit binaries under %ProgramFiles%\Defraggler.

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
    Outputs only the latest available Defraggler version string and exits.

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
$BuildsPageUrl     = "https://www.ccleaner.com/defraggler/builds"
$DownloadUrlFormat = "https://download.ccleaner.com/dfsetup{0}.exe"

$VendorFolder = "Piriform"
$AppFolder    = "Defraggler"

$BaseDownloadRoot = Join-Path $DownloadRoot "Defraggler"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        The CDN answers 200 with an HTML body for retired build numbers, which
        would otherwise stage as a valid-looking installer and fail at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestDefragglerRelease {
    <#
    .SYNOPSIS
        Returns the current Defraggler version and its download URL.
    .DESCRIPTION
        The vendor builds page names the current build (dfsetup222.exe); the
        filename encodes the major and two-digit minor only, and the vendor
        publishes no version list any more. The full version comes from the
        installer's file version resource (2.22.33.995 -> 2.22.995, the form
        the vendor used for releases), so the installer is downloaded into the
        download root when it is not already there.
    #>
    param([switch]$Quiet)

    Write-Log "Builds page URL              : $BuildsPageUrl" -Quiet:$Quiet

    try {
        # The vendor site intermittently answers 404 to back-to-back requests
        # and to requests without a browser user agent; both are transient.
        $ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        $html = $null
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $html = (curl.exe -L --fail --silent --show-error -A $ua $BuildsPageUrl) -join "`n"
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($html)) { break }
            $html = $null
            Start-Sleep -Seconds (2 * $attempt)
        }
        if (-not $html) { throw "Failed to fetch the Defraggler builds page: $BuildsPageUrl" }

        $rx = [regex]'dfsetup(?<slug>\d{3,4})\.exe'
        $slugs = @($rx.Matches($html) | ForEach-Object { $_.Groups['slug'].Value } | Sort-Object -Unique)
        if ($slugs.Count -lt 1) {
            throw "Could not locate a dfsetup<build>.exe link on the builds page."
        }
        $slug = $slugs | Sort-Object { [int]$_ } | Select-Object -Last 1
        $fileName = "dfsetup$slug.exe"
        $downloadUrl = ($DownloadUrlFormat -f $slug)

        Initialize-Folder -Path $BaseDownloadRoot
        $localExe = Join-Path $BaseDownloadRoot $fileName
        if (-not (Test-Path -LiteralPath $localExe)) {
            Write-Log "Downloading installer to read its version: $downloadUrl" -Quiet:$Quiet
            Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', $ua)
        }
        Assert-ExePayload -Path $localExe

        $fileVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
        if ([string]::IsNullOrWhiteSpace($fileVersion) -or $fileVersion.Trim() -notmatch '^\d+\.\d+\.\d+\.\d+$') {
            throw "Installer carries no four-part file version resource: '$fileVersion'"
        }
        $v = [version]$fileVersion.Trim()
        $best = "{0}.{1}.{2}" -f $v.Major, $v.Minor, $v.Revision

        Write-Log "Latest Defraggler version    : $best" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best
            FileName    = $fileName
            DownloadUrl = $downloadUrl
        }
    }
    catch {
        Write-Log "Failed to get Defraggler version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDefraggler {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Defraggler - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestDefragglerRelease
    if (-not $releaseInfo) { throw "Could not resolve Defraggler version." }

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
        Write-Log "Downloading Defraggler..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

    # The installer carries the same four-part file version as the installed
    # binaries, so detection compares against a value read from the payload
    # rather than the three-part marketing version on the history page.
    $fileVersion = (Get-Item -LiteralPath $localExe).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($fileVersion)) {
        throw "Installer carries no file version resource; cannot derive detection value."
    }
    $fileVersion = $fileVersion.Trim()

    Write-Log "Installer file version       : $fileVersion"
    Write-Log ""

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

    # The uninstaller path is not fixed across builds; the ARP UninstallString
    # is the only value guaranteed to point at the copy that installed this client.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Defraggler',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Defraggler'
)
$cmd = $null
foreach ($k in $keys) {
    $p = Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue
    if ($p -and $p.UninstallString) { $cmd = $p.UninstallString; break }
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
    $detectionPath = "{0}\Defraggler" -f $env:ProgramFiles

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : Defraggler64.exe >= $fileVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Defraggler"
        Publisher        = "Piriform Software Ltd."
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/S"
        UninstallCommand = "$detectionPath\uninst.exe"
        UninstallArgs    = "/S"
        RunningProcess   = @("Defraggler", "Defraggler64")
        Detection        = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "Defraggler64.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $fileVersion
            Is64Bit       = $true
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

function Invoke-PackageDefraggler {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Defraggler - PACKAGE phase"
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
        $info = Get-LatestDefragglerRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Defraggler GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Defraggler Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "BuildsPageUrl                : $BuildsPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageDefraggler
    }
    elseif ($PackageOnly) {
        Invoke-PackageDefraggler
    }
    else {
        Invoke-StageDefraggler
        Invoke-PackageDefraggler
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-defraggler'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
