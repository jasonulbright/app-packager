<#
Vendor: Python Software Foundation
App: Python
CMName: Python
VendorUrl: https://www.python.org/downloads/
CPE: cpe:2.3:a:python:python:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.python.org/3/whatsnew/
DownloadPageUrl: https://www.python.org/downloads/

.SYNOPSIS
    Packages the latest Python (x64) for MECM.

.DESCRIPTION
    Queries the endoflife.date API for the latest stable Python release,
    downloads the 64-bit EXE installer from python.org, stages content to a
    versioned local folder with file-existence detection metadata, and creates
    an MECM Application.

    Python installs to a version-specific directory
    (C:\Program Files\Python3XX\), so detection targets python.exe in that
    versioned path.

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
    Content is staged under: <FileServerPath>\Applications\Python Software Foundation\Python\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Python).
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
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Python version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
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
$VersionApiUrl = "https://endoflife.date/api/python.json"
$DownloadBaseUrl = "https://www.python.org/ftp/python"

$VendorFolder = "Python Software Foundation"
$AppFolder    = "Python"

$BaseDownloadRoot = Join-Path $DownloadRoot "Python"

# --- Functions ---


function Get-LatestPythonVersion {
    <#
    .SYNOPSIS
        Queries the endoflife.date API for the latest stable Python release.
        Returns a PSCustomObject with Version, FileName, and DownloadUrl.
    #>
    param([switch]$Quiet)

    Write-Log "Python versions API          : $VersionApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $VersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Python version info: $VersionApiUrl" }

        $cycles = ConvertFrom-Json $json
        if (-not $cycles -or $cycles.Count -eq 0) {
            throw "No cycles found in API response."
        }

        # First element is the latest cycle (e.g., 3.14)
        $version = $cycles[0].latest
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "No latest version found in first cycle."
        }

        $fileName    = "python-$version-amd64.exe"
        $downloadUrl = "$DownloadBaseUrl/$version/$fileName"

        Write-Log "Latest Python version        : $version" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = $downloadUrl
        }
    }
    catch {
        Write-Log "Failed to get Python version info: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePython {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Python (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-LatestPythonVersion
    if (-not $release) { throw "Could not resolve Python version." }

    $version           = $release.Version
    $installerFileName = $release.FileName
    $downloadUrl       = $release.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Python installer..."
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
    # Python uses property-style args and the same EXE for uninstall,
    # so we build wrappers manually instead of New-ExeWrapperContent.
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/quiet'', ''InstallAllUsers=1'', ''PrependPath=1'', ''Include_test=0'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Python Launcher (py.exe) is a separate MSI component that persists
    # after the main Python uninstall. Remove it via ARP ProductCode lookup.
    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/quiet'') -Wait -PassThru -NoNewWindow',
        '$exitCode = $proc.ExitCode',
        '$launcher = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |',
        '    Where-Object { $_.DisplayName -match ''Python Launcher'' } | Select-Object -First 1',
        'if ($launcher) {',
        '    $code = $launcher.PSChildName',
        '    Start-Process msiexec.exe -ArgumentList @(''/x'', $code, ''/qn'', ''/norestart'') -Wait -NoNewWindow',
        '}',
        'exit $exitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # Python installs to C:\Program Files\Python3XX\ where XX = major + minor without dots
    $versionParts = $version -split '\.'
    $majorMinor   = '{0}{1}' -f $versionParts[0], $versionParts[1]
    $detectionPath = '{0}\Python{1}' -f $env:ProgramFiles, $majorMinor

    $appName   = "Python $version (64-bit)"
    $publisher = "Python Software Foundation"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : python.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0"
        UninstallArgs   = "/uninstall /quiet"
        RunningProcess  = @("python", "pythonw")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "python.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackagePython {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Python (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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
        $rel = Get-LatestPythonVersion -Quiet
        if (-not $rel) { exit 1 }
        Write-Output $rel.Version
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
    Write-Log "Python (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionApiUrl                : $VersionApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePython
    }
    elseif ($PackageOnly) {
        Invoke-PackagePython
    }
    else {
        Invoke-StagePython
        Invoke-PackagePython
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
