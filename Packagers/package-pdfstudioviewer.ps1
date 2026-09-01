<#
Vendor: Qoppa Software
App: PDF Studio Viewer
CMName: PDF Studio Viewer
VendorUrl: https://www.qoppa.com/
ReleaseNotesUrl: https://www.qoppa.com/pdfstudioviewer/
DownloadPageUrl: https://www.qoppa.com/pdfstudioviewer/download/
UpdateCadenceDays: 180

.SYNOPSIS
    Packages PDF Studio Viewer (x64) for MECM.

.DESCRIPTION
    Downloads the vendor's unversioned win64 installer, reads the product
    version from the installer's own version resource, stages content to a
    versioned local folder, and creates an MECM Application with file-based
    detection.

    The installer is an install4j package. -q runs it unattended and -dir pins
    the target so the uninstaller path stays predictable across upgrades.

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
    Outputs only the latest available PDF Studio Viewer version string and exits.

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
$InstallerUrl      = "https://download.qoppa.com/pdfstudioviewer/PDFStudioViewer_win64.exe"
$InstallerFileName = "PDFStudioViewer_win64.exe"

$VendorFolder = "Qoppa Software"
$AppFolder    = "PDF Studio Viewer"

$BaseDownloadRoot = Join-Path $DownloadRoot "PDF Studio Viewer"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        The download host answers 200 with an HTML body for blocked requests,
        which would otherwise stage as a valid-looking installer and fail only
        at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-PdfStudioViewerLayout {
    <#
    .SYNOPSIS
        Returns the install directory and launcher name for a given version.
    .DESCRIPTION
        install4j names both after the release year, so a 2024.0.1 build lands
        in PDFStudioViewer2024 and launches PDFStudioViewer2024.exe. Deriving
        both from the version keeps detection correct across yearly releases.
    #>
    param([Parameter(Mandatory)][string]$Version)

    $year = ($Version -split '\.')[0]
    return [pscustomobject]@{
        InstallDir   = "C:\Program Files\PDFStudioViewer$year"
        LauncherName = "PDFStudioViewer$year.exe"
        ProcessName  = "PDFStudioViewer$year"
    }
}


function Get-PdfStudioViewerVersion {
    <#
    .SYNOPSIS
        Returns the product version of the published installer.
    .DESCRIPTION
        The vendor publishes a single unversioned URL and states no version on
        the download page. The installer's own version resource is the only
        published version fact, and it sits in the PE header ahead of the
        install4j payload, so a range request for the first megabytes is enough
        to read it without pulling the full ~170 MB file.
    #>
    param([switch]$Quiet)

    Write-Log "Installer URL                : $InstallerUrl" -Quiet:$Quiet

    $probePath = Join-Path $BaseDownloadRoot "version-probe.bin"

    try {
        Initialize-Folder -Path $BaseDownloadRoot

        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $probePath -ExtraCurlArgs @('-r', '0-4194303') -Quiet:$Quiet
        Assert-ExePayload -Path $probePath

        $version = (Get-Item -LiteralPath $probePath).VersionInfo.ProductVersion
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Installer header carries no ProductVersion."
        }
        $version = $version.Trim()

        Write-Log "Latest PDF Studio Viewer     : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get PDF Studio Viewer version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
    finally {
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePdfStudioViewer {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDF Studio Viewer (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Download URL                 : $InstallerUrl"
    Write-Log "Downloading PDF Studio Viewer..."

    # The URL is unversioned, so a cached copy cannot be trusted to be current;
    # the download always runs and the version comes from what arrived.
    Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe

    Assert-ExePayload -Path $localExe

    $version = (Get-Item -LiteralPath $localExe).VersionInfo.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) { throw "Installer carries no ProductVersion." }
    $version = $version.Trim()

    $layout = Get-PdfStudioViewerLayout -Version $version

    Write-Log ""
    Write-Log "Version                      : $version"
    Write-Log "Install directory            : $($layout.InstallDir)"
    Write-Log "Launcher                     : $($layout.LauncherName)"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # install4j: -q is unattended with installer defaults, -dir pins the target
    # so the uninstaller path stays predictable across upgrades.
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs ("'-q', '-dir', '{0}'" -f $layout.InstallDir) `
        -UninstallCommand ("{0}\uninstall.exe" -f $layout.InstallDir) `
        -UninstallArgs "'-q'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    Write-Log "Detection path               : $($layout.InstallDir)"
    Write-Log "Detection file               : $($layout.LauncherName)"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "PDF Studio Viewer"
        Publisher        = "Qoppa Software"
        SoftwareVersion  = $version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = ("-q -dir `"{0}`"" -f $layout.InstallDir)
        UninstallCommand = ("{0}\uninstall.exe" -f $layout.InstallDir)
        UninstallArgs    = "-q"
        RunningProcess   = @($layout.ProcessName)
        Detection        = @{
            Type         = "File"
            FilePath     = $layout.InstallDir
            FileName     = $layout.LauncherName
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

function Invoke-PackagePdfStudioViewer {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDF Studio Viewer (x64) - PACKAGE phase"
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
        $version = Get-PdfStudioViewerVersion -Quiet
        if (-not $version) { exit 1 }
        Write-Output $version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("PDF Studio Viewer GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDF Studio Viewer (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "InstallerUrl                 : $InstallerUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePdfStudioViewer
    }
    elseif ($PackageOnly) {
        Invoke-PackagePdfStudioViewer
    }
    else {
        Invoke-StagePdfStudioViewer
        Invoke-PackagePdfStudioViewer
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-pdfstudioviewer'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
