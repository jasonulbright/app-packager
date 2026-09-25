<#
Vendor: Jason Ulbright
App: AppPackager Suite (User)
CMName: AppPackager Suite
VendorUrl: https://github.com/jasonulbright/app-packager-suite
CPE: cpe:2.3:a:jasonulbright:app_packager_suite:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/jasonulbright/app-packager-suite/releases
DownloadPageUrl: https://github.com/jasonulbright/app-packager-suite/releases/latest
IconSource: Installer
UpdateCadenceDays: 30

.SYNOPSIS
    Packages AppPackager Suite for ConfigMgr as a per-user install.

.DESCRIPTION
    Downloads the latest SuiteSetup-<version>.exe from the GitHub releases
    API, stages content to a versioned local folder, and creates a ConfigMgr
    Application that installs for the logged-on user.

    SuiteSetup is an NSIS installer that requests no elevation: /S installs
    silently into %LOCALAPPDATA%\AppPackagerSuite and writes an HKCU
    uninstall entry. Detection reads that entry's DisplayVersion.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Jason Ulbright\AppPackager Suite\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AppPackagerSuite).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download the installer, generate content
    wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available AppPackager Suite version string and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/jasonulbright/app-packager-suite/releases/latest"

$VendorFolder = "Jason Ulbright"
$AppFolder    = "AppPackager Suite"

$BaseDownloadRoot = Join-Path $DownloadRoot "AppPackagerSuite"

# The installer writes its ARP entry under HKCU with a fixed key name.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AppPackagerSuite"

$InstallDir = "%LOCALAPPDATA%\AppPackagerSuite"

# --- Functions ---


function ConvertFrom-AppPackagerSuiteAssetName {
    <#
    .SYNOPSIS
        Returns the version carried by a SuiteSetup-<version>.exe asset name,
        or nothing for any other asset.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)

    $m = [regex]::Match($Name, '^SuiteSetup-(?<ver>\d+(?:\.\d+){3})\.exe$')
    if (-not $m.Success) { return $null }
    return $m.Groups['ver'].Value
}


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


function Get-LatestAppPackagerSuiteRelease {
    param([switch]$Quiet)

    Write-Log "GitHub API URL               : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error -A "PowerShell" @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch GitHub release info: $GitHubApiUrl" }

        $release = ConvertFrom-Json $json

        # The release also carries a portable zip and checksums.txt; only the
        # setup asset name carries the four-part version.
        $asset = $release.assets |
            Where-Object { ConvertFrom-AppPackagerSuiteAssetName -Name ([string]$_.name) } |
            Select-Object -First 1

        if (-not $asset) {
            throw "Could not find a SuiteSetup asset in the latest GitHub release."
        }

        $version = ConvertFrom-AppPackagerSuiteAssetName -Name $asset.name

        Write-Log "Latest AppPackager Suite     : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $asset.browser_download_url
            FileName    = $asset.name
        }
    }
    catch {
        Write-Log "Failed to get AppPackager Suite version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAppPackagerSuite {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AppPackager Suite (User) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestAppPackagerSuiteRelease
    if (-not $releaseInfo) { throw "Could not resolve AppPackager Suite version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Install context              : User"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', 'PowerShell')
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe
    Assert-ArpDetectionKey -InstallerPath $localExe -ExpectedKey $ArpRegistryKey -Is64BitView $false

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

    # The NSIS uninstaller copies itself to a temp Au_.exe and returns before
    # that copy finishes, so the wrapper waits for the copy before reporting.
    $uninstallContent = @'
$uninstaller = Join-Path $env:LOCALAPPDATA 'AppPackagerSuite\Uninstall.exe'
if (-not (Test-Path -LiteralPath $uninstaller)) { exit 0 }
$proc = Start-Process -FilePath $uninstaller -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
while (Get-Process -Name 'Au_' -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 1 }
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : HKCU\$ArpRegistryKey"
    Write-Log "Install directory            : $InstallDir"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = "AppPackager Suite (User)"
        Publisher                = "Jason Ulbright"
        SoftwareVersion          = $version
        DisplayName              = "AppPackager Suite"
        InstallerFile            = $installerFileName
        InstallerType            = "EXE"
        InstallArgs              = "/S"
        UninstallCommand         = "%LOCALAPPDATA%\AppPackagerSuite\Uninstall.exe"
        UninstallArgs            = "/S"
        RunningProcess           = @()
        InstallationBehaviorType = "InstallForUser"
        LogonRequirementType     = "OnlyWhenUserLoggedOn"
        Detection                = @{
            Type                = "RegistryKeyValue"
            Hive                = "CurrentUser"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
            Is64Bit             = $false
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

function Invoke-PackageAppPackagerSuite {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AppPackager Suite (User) - PACKAGE phase"
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
    Write-Log "Install behavior             : $($manifest.InstallationBehaviorType)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

    # --- ConfigMgr application ---
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
        $info = Get-LatestAppPackagerSuiteRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("AppPackager Suite GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "AppPackager Suite Auto-Packager starting"
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
        Invoke-StageAppPackagerSuite
    }
    elseif ($PackageOnly) {
        Invoke-PackageAppPackagerSuite
    }
    else {
        Invoke-StageAppPackagerSuite
        Invoke-PackageAppPackagerSuite
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-apppackagersuite'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
