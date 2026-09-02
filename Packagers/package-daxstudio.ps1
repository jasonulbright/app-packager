<#
Vendor: DAX Studio
App: DAX Studio
CMName: DAX Studio
VendorUrl: https://daxstudio.org/
CPE: cpe:2.3:a:daxstudio:dax_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/DaxStudio/DaxStudio/releases
DownloadPageUrl: https://daxstudio.org/downloads/
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages DAX Studio (x64) for MECM.

.DESCRIPTION
    Resolves the latest DAX Studio installer from the GitHub releases API, stages
    content to a versioned local folder, and creates an MECM Application with
    file-existence detection.

    The installer is an InnoSetup package installed silently for all users.

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
    Outputs only the latest available DAX Studio version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/DaxStudio/DaxStudio/releases/latest"

$VendorFolder = "DAX Studio"
$AppFolder    = "DAX Studio"

$BaseDownloadRoot = Join-Path $DownloadRoot "DaxStudio"
$InstallPath      = "C:\Program Files\DAX Studio"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A release-asset redirect that lands on an error page answers 200 with
        HTML, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestDaxStudioRelease {
    <#
    .SYNOPSIS
        Returns the newest DAX Studio version and its setup EXE URL.
    .DESCRIPTION
        The asset filename separates version components with underscores while
        the tag uses dots, so the dotted version is reconstructed from the
        asset name rather than parsed out of the tag.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        # The API rejects requests without a user agent.
        $json = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" $GitHubApiUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $m = [regex]::Match($json, '"name":\s*"DaxStudio_(?<ver>\d+(?:_\d+)+)_setup\.exe"')
        if (-not $m.Success) { throw "Could not locate a setup asset in the latest release." }

        $underscoreVersion = $m.Groups['ver'].Value
        $version           = $underscoreVersion -replace '_', '.'
        $fileName          = "DaxStudio_${underscoreVersion}_setup.exe"

        Write-Log "Latest DAX Studio version    : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = "https://github.com/DaxStudio/DaxStudio/releases/download/v$version/$fileName"
        }
    }
    catch {
        Write-Log "Failed to get DAX Studio version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDaxStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DAX Studio (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestDaxStudioRelease
    if (-not $releaseInfo) { throw "Could not resolve DAX Studio version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName
    $downloadUrl       = $releaseInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading DAX Studio..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/ALLUSERS'" `
        -UninstallCommand "$InstallPath\unins000.exe" `
        -UninstallArgs "'/SP-', '/VERYSILENT', '/NORESTART'"

    # The InnoSetup uninstaller aborts while the application holds its files.
    # Absent uninstaller means the product is not present, so removal exits clean.
    $customUninstall = (
        'Get-Process DaxStudio -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        'Start-Sleep -Seconds 2',
        ('$uninstaller = ''{0}\unins000.exe''' -f $InstallPath),
        'if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/SP-'', ''/VERYSILENT'', ''/SUPPRESSMSGBOXES'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : DaxStudio.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "DAX Studio"
        Publisher       = "DAX Studio"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /ALLUSERS"
        UninstallArgs   = "/SP- /VERYSILENT /NORESTART"
        RunningProcess  = @("DaxStudio")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallPath
            FileName     = "DaxStudio.exe"
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

function Invoke-PackageDaxStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DAX Studio (x64) - PACKAGE phase"
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
        $info = Get-LatestDaxStudioRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("DAX Studio GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DAX Studio (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "GitHubApiUrl                 : $GitHubApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageDaxStudio
    }
    elseif ($PackageOnly) {
        Invoke-PackageDaxStudio
    }
    else {
        Invoke-StageDaxStudio
        Invoke-PackageDaxStudio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-daxstudio'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
