<#
Vendor: win.rar GmbH
App: WinRAR
CMName: WinRAR
VendorUrl: https://www.win-rar.com/
CPE: cpe:2.3:a:rarlab:winrar:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.win-rar.com/whatsnew.html
DownloadPageUrl: https://www.win-rar.com/download.html

.SYNOPSIS
    Packages WinRAR (x64) for MECM.

.DESCRIPTION
    Downloads the latest WinRAR x64 EXE from rarlab.com, stages content to a
    versioned local folder with ARP-based detection metadata, and creates an
    MECM Application.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest version is resolved by scraping the rarlab.com download page for
    the x64 installer filename (which encodes the version number).

    NOTE: WinRAR is trialware. For licensed enterprise deployments, place a
    rarreg.key file in the staged content folder before packaging. The install
    wrapper copies it to the WinRAR install directory after installation.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\win.rar GmbH\WinRAR (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Scrapes the rarlab.com download page for the latest version, outputs the
    version string, and exits. No MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$DownloadPageUrl = "https://www.rarlab.com/download.htm"
$CdnBaseUrl      = "https://www.rarlab.com/rar"

$VendorFolder = "win.rar GmbH"
$AppFolder    = "WinRAR (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "WinRAR"

# --- Functions ---


function Get-LatestWinRARVersion {
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch rarlab.com download page." }

        # Parse filename like "winrar-x64-720.exe" from the page
        if ($html -match 'winrar-x64-(\d+)\.exe') {
            $verDigits = $Matches[1]
        }
        else {
            throw "Could not find winrar-x64 installer link on download page."
        }

        # Convert digit string to version: "720" -> "7.20", "701" -> "7.01"
        if ($verDigits.Length -eq 3) {
            $version = "{0}.{1}" -f $verDigits.Substring(0,1), $verDigits.Substring(1,2)
        }
        elseif ($verDigits.Length -eq 4) {
            $version = "{0}.{1}" -f $verDigits.Substring(0,2), $verDigits.Substring(2,2)
        }
        else {
            throw "Unexpected version digit format: $verDigits"
        }

        Write-Log "Latest WinRAR version        : $version" -Quiet:$Quiet
        return [PSCustomObject]@{
            Version       = $version
            VerDigits     = $verDigits
        }
    }
    catch {
        Write-Log "Failed to get WinRAR version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWinRAR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinRAR (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $versionInfo = Get-LatestWinRARVersion
    if (-not $versionInfo) { throw "Could not determine latest WinRAR version." }

    $version   = $versionInfo.Version
    $verDigits = $versionInfo.VerDigits

    $installerFileName = "winrar-x64-$verDigits.exe"
    $downloadUrl = "$CdnBaseUrl/$installerFileName"

    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading WinRAR..."
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
    # WinRAR install wrapper also copies rarreg.key if present (for licensed deployments)
    $installPs1 = @(
        "`$exePath = Join-Path `$PSScriptRoot '$installerFileName'"
        "`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/S') -Wait -PassThru -NoNewWindow"
        "if (`$proc.ExitCode -ne 0) { exit `$proc.ExitCode }"
        ""
        "# Copy license key if present in content folder"
        "`$keyFile = Join-Path `$PSScriptRoot 'rarreg.key'"
        "if (Test-Path -LiteralPath `$keyFile) {"
        "    Copy-Item -LiteralPath `$keyFile -Destination `"`$env:ProgramFiles\WinRAR\rarreg.key`" -Force"
        "}"
        "exit 0"
    ) -join "`r`n"

    $uninstallPs1 = @(
        "`$uninstaller = `"`$env:ProgramFiles\WinRAR\uninstall.exe`""
        "if (-not (Test-Path -LiteralPath `$uninstaller)) { Write-Error 'WinRAR uninstaller not found.'; exit 1 }"
        "`$proc = Start-Process -FilePath `$uninstaller -ArgumentList @('/S') -Wait -PassThru -NoNewWindow"
        "exit `$proc.ExitCode"
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WinRAR archiver"

    $appName   = "WinRAR $version (64-bit)"
    $publisher = "win.rar GmbH"

    Write-Log "ARP Registry Key             : $arpRegistryKey"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("WinRAR")
        Detection       = @{
            # WinRAR writes DisplayVersion as "7.20.0" but the vendor page
            # reports just "7.20". BeginsWith tolerates the trailing ".0"
            # and any future 4th-octet additions.
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            Operator            = "BeginsWith"
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

function Invoke-PackageWinRAR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinRAR (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

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

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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
        $info = Get-LatestWinRARVersion -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinRAR (x64) Auto-Packager starting"
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
        Invoke-StageWinRAR
    }
    elseif ($PackageOnly) {
        Invoke-PackageWinRAR
    }
    else {
        Invoke-StageWinRAR
        Invoke-PackageWinRAR
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
