<#
Vendor: Signal Ridge Labs
App: Spectra PDF
CMName: Spectra PDF
VendorUrl: https://github.com/jasonulbright/spectra-pdf
CPE: cpe:2.3:a:signalridgelabs:spectra_pdf:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/jasonulbright/spectra-pdf/releases
DownloadPageUrl: https://github.com/jasonulbright/spectra-pdf/releases
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Spectra PDF (first-party, x64) for MECM.

.DESCRIPTION
    First-party Tauri/NSIS application released through GitHub releases on
    a private repository, so acquisition goes through the authenticated
    GitHub CLI (gh) instead of anonymous HTTP - the one packager in this
    set with that requirement. Install is perMachine NSIS /S; uninstall is
    the fixed "C:\Program Files\Spectra PDF\uninstall.exe" /S (user data
    kept for redeployment; pass /removeuserdata manually for full removal).

    Supports the standard two-phase contract:
      -StageOnly    Download release asset via gh, write wrappers + manifest
      -PackageOnly  Copy to network, create MECM application

.REQUIREMENTS
    - GitHub CLI (gh) authenticated with access to jasonulbright/spectra-pdf
    - PowerShell 5.1, ConfigMgr console for the Package phase
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
$Repo             = 'jasonulbright/spectra-pdf'
$VendorFolder     = 'Signal Ridge Labs'
$AppFolder        = 'Spectra PDF'
$BaseDownloadRoot = Join-Path $DownloadRoot 'SpectraPDF'

function Assert-GhAvailable {
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if (-not $gh) { throw "GitHub CLI (gh) is required for the private Spectra PDF repository and was not found on PATH." }
}

function Get-SpectraPdfLatestRelease {
    param([switch]$Quiet)
    Assert-GhAvailable
    $json = gh release view --repo $Repo --json tagName,assets 2>$null | ConvertFrom-Json
    if (-not $json -or -not $json.tagName) { throw "Could not read the latest release of $Repo (gh authenticated?)." }
    $version = ([string]$json.tagName).TrimStart('v')
    $asset = @($json.assets | Where-Object { $_.name -match '_x64-setup\.exe$' }) | Select-Object -First 1
    if (-not $asset) { throw "Release $($json.tagName) carries no x64 NSIS setup asset." }
    Write-Log "Latest Spectra PDF release    : $version ($($asset.name))" -Quiet:$Quiet
    return [pscustomobject]@{ Version = $version; AssetName = [string]$asset.name }
}

# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSpectraPdf {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Spectra PDF (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot
    $release = Get-SpectraPdfLatestRelease
    $version = $release.Version

    $localExe = Join-Path $BaseDownloadRoot $release.AssetName
    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading release asset    : $($release.AssetName)"
        gh release download --repo $Repo --pattern $release.AssetName --dir $BaseDownloadRoot --clobber
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $localExe)) { throw "gh release download failed." }
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # A repo/API error page must never stage as an installer.
    $head = New-Object byte[] 2
    $fs = [System.IO.File]::OpenRead($localExe)
    try { [void]$fs.Read($head, 0, 2) } finally { $fs.Dispose() }
    if ($head[0] -ne 0x4D -or $head[1] -ne 0x5A) { throw "Downloaded file is not a valid EXE (no MZ header): $($release.AssetName)" }

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath
    Copy-Item -LiteralPath $localExe -Destination (Join-Path $localContentPath $release.AssetName) -Force -ErrorAction Stop
    Write-Log "Copied installer to staged   : $localContentPath"

    # Tauri NSIS, perMachine (tauri.conf.json): install /S; the uninstaller
    # sits at a fixed path and keeps user data unless /removeuserdata.
    $wrappers = New-ExeWrapperContent -InstallerFileName $release.AssetName -InstallArgs "'/S'" `
        -UninstallCommand 'C:\Program Files\Spectra PDF\uninstall.exe' -UninstallArgs "'/S'"
    $customUninstall = (
        '$u = ''C:\Program Files\Spectra PDF\uninstall.exe''',
        'if (-not (Test-Path -LiteralPath $u)) { exit 0 }',
        '$proc = Start-Process -FilePath $u -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $customUninstall

    $manifestPath = Join-Path $localContentPath 'stage-manifest.json'
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = 'Spectra PDF'
        Publisher       = 'Signal Ridge Labs'
        SoftwareVersion = $version
        InstallerFile   = $release.AssetName
        InstallerType   = 'EXE'
        InstallArgs     = '/S'
        UninstallArgs   = '/S'
        UninstallCommand = '"C:\Program Files\Spectra PDF\uninstall.exe" /S'
        RunningProcess  = @('Spectra PDF')
        Detection       = @{
            # Tauri NSIS writes the ARP key named after productName.
            Type                = 'RegistryKeyValue'
            RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Spectra PDF'
            ValueName           = 'DisplayVersion'
            Operator            = 'GreaterEquals'
            ExpectedValue       = $version
            Is64Bit             = $true
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}

# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageSpectraPdf {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Spectra PDF (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot
    $release = Get-SpectraPdfLatestRelease -Quiet
    $localContentPath = Join-Path $BaseDownloadRoot $release.Version
    $manifestPath = Join-Path $localContentPath 'stage-manifest.json'
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }
    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout
    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    foreach ($f in (Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop)) {
        if ($f.Name -eq 'stage-manifest.json') { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else { Write-Log "Already on network           : $relative" }
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
        $release = Get-SpectraPdfLatestRelease -Quiet
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Spectra PDF GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Spectra PDF Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) { Invoke-StageSpectraPdf }
    elseif ($PackageOnly) { Invoke-PackageSpectraPdf }
    else {
        Invoke-StageSpectraPdf
        Invoke-PackageSpectraPdf
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-spectrapdf'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
