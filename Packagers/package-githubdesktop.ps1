<#
Vendor: GitHub
App: GitHub Desktop (User)
CMName: GitHub Desktop
VendorUrl: https://desktop.github.com/
CPE: cpe:2.3:a:github:github_desktop:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://desktop.github.com/release-notes/
DownloadPageUrl: https://desktop.github.com/download/
IconSource: Installer
UpdateCadenceDays: 30

.SYNOPSIS
    Packages GitHub Desktop (x64) for ConfigMgr as a per-user install.

.DESCRIPTION
    Resolves the newest GitHub Desktop build from the vendor's download
    endpoint, which redirects to a versioned installer URL, downloads
    GitHubDesktopSetup-x64.exe, stages content to a versioned local folder, and
    creates a ConfigMgr Application that installs for the logged-on user.

    The GitHub repository does not publish every shipped build as a GitHub
    release, so the releases API can lag the version the app updates to.

    GitHubDesktopSetup-x64.exe is a Squirrel installer: -s installs silently
    into %LOCALAPPDATA%\GitHubDesktop and writes an HKCU uninstall entry.
    Detection reads that entry's DisplayVersion.

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
    Content is staged under: <FileServerPath>\Applications\GitHub\GitHub Desktop\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\GitHubDesktop).
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
    Outputs only the latest available GitHub Desktop version string and exits.

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
$DownloadEndpoint = "https://central.github.com/deployments/desktop/desktop/latest/win32"

$VendorFolder = "GitHub"
$AppFolder    = "GitHub Desktop"

$BaseDownloadRoot = Join-Path $DownloadRoot "GitHubDesktop"

$InstallerFileName = "GitHubDesktopSetup-x64.exe"

# Squirrel names the per-user ARP entry after the package id.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GitHubDesktop"

# --- Functions ---

function ConvertFrom-GitHubDesktopDownloadUrl {
    <#
    .SYNOPSIS
        Returns the version from a resolved installer URL such as
        https://desktop.githubusercontent.com/releases/3.6.6-8b85519e/GitHubDesktopSetup-x64.exe.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Url)

    $m = [regex]::Match($Url, '/releases/(?<ver>\d+(?:\.\d+){1,3})-[0-9A-Za-z]+/GitHubDesktopSetup-x64\.exe$')
    if (-not $m.Success) { return $null }
    return $m.Groups['ver'].Value
}


function Get-LatestGitHubDesktopRelease {
    <#
    .SYNOPSIS
        Returns the newest GitHub Desktop version and its installer URL.
    .DESCRIPTION
        The version is taken from the redirect target of the vendor's
        unversioned download endpoint, which is the build the app updates to.
    #>
    param([switch]$Quiet)

    Write-Log "Download endpoint            : $DownloadEndpoint" -Quiet:$Quiet

    try {
        $resolved = (curl.exe -sIL --fail --show-error -A "Mozilla/5.0" -o NUL -w "%{url_effective}" $DownloadEndpoint) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve the GitHub Desktop download redirect: $DownloadEndpoint" }

        $version = ConvertFrom-GitHubDesktopDownloadUrl -Url $resolved
        if (-not $version) { throw "Redirect target does not carry a versioned installer path: $resolved" }

        Write-Log "Latest GitHub Desktop version: $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $resolved
        }
    }
    catch {
        Write-Log "Failed to get GitHub Desktop version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function New-GitHubDesktopInstallWrapper {
    # The Squirrel setup parent exits within seconds while child processes
    # finish the install; without polling for GitHubDesktop.exe the deployment
    # reports success against an empty install folder.
    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        'if (-not (Test-Path -LiteralPath $exePath)) { Write-Error "Missing GitHub Desktop installer"; exit 2 }',
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''-s'') -Wait -PassThru -NoNewWindow',
        '$installedExe = Join-Path $env:LOCALAPPDATA ''GitHubDesktop\GitHubDesktop.exe''',
        '$deadline = (Get-Date).AddMinutes(5)',
        'while (-not (Test-Path -LiteralPath $installedExe)) {',
        '    if ((Get-Date) -gt $deadline) {',
        '        Write-Error "GitHub Desktop install did not produce $installedExe within 5 minutes"',
        '        exit 4',
        '    }',
        '    Start-Sleep -Seconds 5',
        '}',
        'Start-Sleep -Seconds 3',
        'try { Stop-Process -Name GitHubDesktop -Force -ErrorAction SilentlyContinue } catch { }',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}


function New-GitHubDesktopUninstallWrapper {
    # Squirrel removes the ARP entry and reports success while the app is
    # running, leaving the app-<version> folder and the process in place.
    return (
        'Get-Process GitHubDesktop -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        'Start-Sleep -Seconds 2',
        '$updateExe = Join-Path $env:LOCALAPPDATA ''GitHubDesktop\Update.exe''',
        'if (-not (Test-Path -LiteralPath $updateExe)) { exit 0 }',
        '$proc = Start-Process -FilePath $updateExe -ArgumentList @(''--uninstall'', ''-s'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGitHubDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GitHub Desktop (User) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestGitHubDesktopRelease
    if (-not $releaseInfo) { throw "Could not resolve GitHub Desktop version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log "Install context              : User"
    Write-Log ""

    # --- Download ---
    # The installer name carries no version, so a cached copy from an earlier
    # build would shadow the new one; the payload is re-fetched every stage.
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    Assert-PayloadIsExecutable -Path $localExe

    Assert-ArpDetectionKey -InstallerPath $localExe -ExpectedKey $ArpRegistryKey -Is64BitView $false

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localExe -Destination (Join-Path $localContentPath $InstallerFileName) -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-GitHubDesktopInstallWrapper) `
        -UninstallPs1Content (New-GitHubDesktopUninstallWrapper)

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = "GitHub Desktop (User)"
        Publisher                = "GitHub"
        SoftwareVersion          = $version
        DisplayName              = "GitHub Desktop"
        InstallerFile            = $InstallerFileName
        InstallerType            = "EXE"
        InstallArgs              = "-s"
        UninstallCommand         = "%LOCALAPPDATA%\GitHubDesktop\Update.exe"
        UninstallArgs            = "--uninstall -s"
        RunningProcess           = @("GitHubDesktop")
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

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageGitHubDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GitHub Desktop (User) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }

    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()
    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifest = Read-StageManifest -Path (Join-Path $localContentPath "stage-manifest.json")

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Install behavior             : $($manifest.InstallationBehaviorType)"
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
        $info = Get-LatestGitHubDesktopRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GitHub Desktop GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "GitHub Desktop (User) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadEndpoint             : $DownloadEndpoint"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGitHubDesktop
    }
    elseif ($PackageOnly) {
        Invoke-PackageGitHubDesktop
    }
    else {
        Invoke-StageGitHubDesktop
        Invoke-PackageGitHubDesktop
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-githubdesktop'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
