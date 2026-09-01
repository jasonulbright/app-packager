<#
Vendor: Anaconda, Inc.
App: Anaconda Distribution
CMName: Anaconda
VendorUrl: https://www.anaconda.com/
CPE: cpe:2.3:a:anaconda:anaconda:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.anaconda.com/anaconda/release-notes/
DownloadPageUrl: https://repo.anaconda.com/archive/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Anaconda Distribution (x64) for MECM.

.DESCRIPTION
    Resolves the latest Anaconda3 Windows x86_64 installer from the vendor
    archive index, stages content to a versioned local folder, and creates an
    MECM Application with file-existence detection.

    The installer is an NSIS package. It is installed for all users into
    C:\ProgramData\Anaconda3 with Python registration disabled so it does not
    take over the machine-wide Python file associations.

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
    Estimated runtime in minutes. Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 90

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Anaconda version string and exits.

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
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 90,
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
$ArchiveIndexUrl = "https://repo.anaconda.com/archive/"
$DownloadBase    = "https://repo.anaconda.com/archive"

$VendorFolder = "Anaconda"
$AppFolder    = "Anaconda"

$BaseDownloadRoot = Join-Path $DownloadRoot "Anaconda"
$InstallPath      = "C:\ProgramData\Anaconda3"

# --- Functions ---


function Get-LatestAnacondaVersion {
    param([switch]$Quiet)

    Write-Log "Archive index URL            : $ArchiveIndexUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $ArchiveIndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Anaconda archive index: $ArchiveIndexUrl" }

        $rx = [regex]'Anaconda3-(?<ver>\d{4}\.\d{2}-\d+)-Windows-x86_64\.exe'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any Windows x86_64 installer entries on the archive index."
        }

        # Archive names sort chronologically as text: YYYY.MM-N.
        $version = ($rxMatches | ForEach-Object { $_.Groups['ver'].Value } |
                    Sort-Object -Unique | Select-Object -Last 1)
        if ([string]::IsNullOrWhiteSpace($version)) { throw "Version match was empty." }

        Write-Log "Latest Anaconda version      : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Anaconda version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAnaconda {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anaconda Distribution (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $version = Get-LatestAnacondaVersion
    if (-not $version) { throw "Could not resolve Anaconda version." }

    $installerFileName = "Anaconda3-$version-Windows-x86_64.exe"
    $downloadUrl       = "$DownloadBase/$installerFileName"

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    # The payload is roughly 1 GB; the cached copy under the download root is
    # reused so a repeated Stage does not re-transfer it.
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Anaconda (large payload)..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
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
    # NSIS requires /D last and unquoted; it is the final element of the
    # argument array for that reason.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/InstallationType=AllUsers', '/RegisterPython=0', '/S', '/D=$InstallPath'" `
        -UninstallCommand "$InstallPath\Uninstall-Anaconda3.exe" `
        -UninstallArgs "'/S'"

    # The uninstaller lives inside the install directory and is absent on a
    # machine that never received the product; exit 0 keeps removal
    # idempotent instead of failing the deployment type.
    $customUninstall = (
        ('$uninstaller = ''{0}\Uninstall-Anaconda3.exe''' -f $InstallPath),
        'if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    $appName   = "Anaconda"
    $publisher = "Anaconda, Inc."

    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : python.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/InstallationType=AllUsers /RegisterPython=0 /S /D=$InstallPath"
        UninstallArgs   = "/S"
        RunningProcess  = @("anaconda-navigator", "python", "jupyter-notebook")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallPath
            FileName     = "python.exe"
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

function Invoke-PackageAnaconda {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anaconda Distribution (x64) - PACKAGE phase"
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
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $v = Get-LatestAnacondaVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Anaconda GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anaconda Distribution (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ArchiveIndexUrl              : $ArchiveIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAnaconda
    }
    elseif ($PackageOnly) {
        Invoke-PackageAnaconda
    }
    else {
        Invoke-StageAnaconda
        Invoke-PackageAnaconda
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-anaconda'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
