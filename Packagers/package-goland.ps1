<#
Vendor: JetBrains
App: GoLand
CMName: GoLand
VendorUrl: https://www.jetbrains.com/go/
CPE: cpe:2.3:a:jetbrains:goland:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.jetbrains.com/go/whatsnew/
DownloadPageUrl: https://www.jetbrains.com/go/download/?section=windows
UpdateCadenceDays: 60

.SYNOPSIS
    Packages GoLand (x64) for MECM.

.DESCRIPTION
    Resolves the latest GoLand release from the JetBrains product releases API,
    downloads the Windows installer, stages content to a versioned local folder,
    and creates an MECM Application with registry-key detection.

    The installer is an NSIS package installed silently with /S. Detection is the
    presence of the versioned ARP uninstall key ("GoLand <version>").

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
    Maximum allowed runtime in minutes. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available GoLand version string and exits.

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
    [int]$MaximumRuntimeMins = 45,
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
$ReleasesApiUrl = "https://data.services.jetbrains.com/products/releases?code=GO&latest=true&type=release"

$VendorFolder = "JetBrains"
$AppFolder    = "GoLand"

$BaseDownloadRoot = Join-Path $DownloadRoot "GoLand"

# --- Functions ---


function Assert-ExecutablePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A CDN error page or consent interstitial answers 200 with HTML, which
        would otherwise stage as a valid-looking installer.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($bytes, 0, 2) } finally { $stream.Dispose() }
    if ($read -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestGoLandRelease {
    <#
    .SYNOPSIS
        Returns the newest GoLand release version and its Windows installer URL.
    .DESCRIPTION
        The releases API reports the download link per platform; the windows
        entry is read from the response rather than assembled from the version
        so a JetBrains filename change does not produce a 404 at download time.
    #>
    param([switch]$Quiet)

    Write-Log "Releases API URL             : $ReleasesApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleasesApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the JetBrains releases API." }

        $data = ConvertFrom-Json $json
        $release = $data.GO | Select-Object -First 1
        if (-not $release) { throw "The releases API returned no GO entries." }

        $version = $release.version
        if ([string]::IsNullOrWhiteSpace($version)) { throw "Release entry carries no version." }

        $link = $release.downloads.windows.link
        if ([string]::IsNullOrWhiteSpace($link)) { throw "Release $version carries no windows download link." }

        $fileName = Split-Path -Path ([uri]$link).AbsolutePath -Leaf

        Write-Log "Latest GoLand version        : $version" -Quiet:$Quiet
        Write-Log "Installer filename           : $fileName" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = $link
        }
    }
    catch {
        Write-Log "Failed to get GoLand version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGoLand {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GoLand (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestGoLandRelease
    if (-not $releaseInfo) { throw "Could not resolve GoLand version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading GoLand..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExecutablePayload -Path $localExe

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

    # --- Detection key ---
    $arpKeyRelative = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GoLand $version"

    # --- Generate content wrappers ---
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand 'unused'

    # The install directory carries the version, so the uninstaller path is
    # read from the ARP UninstallString rather than assumed.
    $customUninstall = (
        ('$key = ''HKLM:\{0}''' -f $arpKeyRelative),
        'if (-not (Test-Path -LiteralPath $key)) { exit 0 }',
        '$uninstaller = (Get-ItemProperty -LiteralPath $key -Name UninstallString -ErrorAction SilentlyContinue).UninstallString',
        'if ([string]::IsNullOrWhiteSpace($uninstaller)) { exit 0 }',
        '$uninstaller = $uninstaller.Trim(''"'')',
        'if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection key                : $arpKeyRelative"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "GoLand"
        Publisher       = "JetBrains"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("goland64")
        Detection       = @{
            Type                = "RegistryKey"
            RegistryKeyRelative = $arpKeyRelative
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

function Invoke-PackageGoLand {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GoLand (x64) - PACKAGE phase"
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
        $info = Get-LatestGoLandRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GoLand GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GoLand (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleasesApiUrl               : $ReleasesApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGoLand
    }
    elseif ($PackageOnly) {
        Invoke-PackageGoLand
    }
    else {
        Invoke-StageGoLand
        Invoke-PackageGoLand
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-goland'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
