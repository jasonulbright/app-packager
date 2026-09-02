<#
Vendor: PDFgear Software
App: PDFgear
CMName: PDFgear
VendorUrl: https://www.pdfgear.com/
ReleaseNotesUrl: https://www.pdfgear.com/download/
DownloadPageUrl: https://www.pdfgear.com/download/
IconSource: Installer
UpdateCadenceDays: 60

.SYNOPSIS
    Packages PDFgear (x64) for MECM.

.DESCRIPTION
    Resolves the current build from the vendor download page, downloads the
    setup executable, stages content to a versioned local folder, and creates
    an MECM Application with registry-based detection.

    The installer is an InnoSetup package that elevates itself and installs
    per-machine under Program Files.

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
    Outputs only the latest available PDFgear version string and exits.

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
$DownloadPageUrl = "https://www.pdfgear.com/download/"

$VendorFolder = "PDFgear"
$AppFolder    = "PDFgear"

$BaseDownloadRoot = Join-Path $DownloadRoot "PDFgear"
$InstallDir       = "C:\Program Files\PDFgear"

# The InnoSetup AppId is fixed across releases, so the ARP key stays stable
# while DisplayVersion tracks the installed build.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{7DACF63A-4EE4-4837-9AF9-C65D4509FFB4}_is1"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        The download host answers 200 with an HTML body for withdrawn builds,
        which would otherwise stage as a valid-looking installer and fail only
        at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestPdfgearRelease {
    <#
    .SYNOPSIS
        Returns the newest PDFgear version and its Windows setup URL.
    .DESCRIPTION
        The download page carries the versioned installer link directly. It can
        reference more than one build, so the highest [version] wins rather than
        document order.
    #>
    param([switch]$Quiet)

    Write-Log "PDFgear download page        : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch PDFgear download page: $DownloadPageUrl" }

        $rx = [regex]'(?<url>https://downloadfiles\.pdfgear\.com/releases/windows/pdfgear_setup_v(?<ver>\d+(?:\.\d+)+)\.exe)'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate a Windows installer link on the download page."
        }

        $candidates = foreach ($m in $rxMatches) {
            [pscustomobject]@{
                Version = $m.Groups['ver'].Value
                Url     = $m.Groups['url'].Value
            }
        }

        $best = $candidates | Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1

        Write-Log "Latest PDFgear version       : $($best.Version)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            FileName    = Split-Path -Leaf $best.Url
            DownloadUrl = $best.Url
        }
    }
    catch {
        Write-Log "Failed to get PDFgear version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePdfgear {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFgear (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestPdfgearRelease
    if (-not $releaseInfo) { throw "Could not resolve PDFgear version." }

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
        Write-Log "Downloading PDFgear..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

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
        -UninstallCommand "$InstallDir\unins000.exe" `
        -UninstallArgs "'/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"

    # The InnoSetup uninstaller aborts while the application holds its files.
    # An absent uninstaller means the product is not present, so removal exits
    # clean.
    $uninstallScript = (
        'Get-Process PDFgear -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        'Start-Sleep -Seconds 2',
        ('$uninstaller = ''{0}\unins000.exe''' -f $InstallDir),
        'if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/SP-'', ''/VERYSILENT'', ''/SUPPRESSMSGBOXES'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallScript

    # --- Write stage manifest ---
    Write-Log "Detection key                : $ArpRegistryKey"
    Write-Log "Detection value              : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "PDFgear"
        Publisher        = "PDFgear Software"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES"
        UninstallCommand = "$InstallDir\unins000.exe"
        UninstallArgs    = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess   = @("PDFgear")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
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

function Invoke-PackagePdfgear {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFgear (x64) - PACKAGE phase"
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
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $info = Get-LatestPdfgearRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("PDFgear GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFgear (x64) Auto-Packager starting"
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
        Invoke-StagePdfgear
    }
    elseif ($PackageOnly) {
        Invoke-PackagePdfgear
    }
    else {
        Invoke-StagePdfgear
        Invoke-PackagePdfgear
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-pdfgear'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
