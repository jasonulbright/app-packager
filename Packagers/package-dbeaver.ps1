<#
Vendor: DBeaver Corp
App: DBeaver Community
CMName: DBeaver Community
VendorUrl: https://dbeaver.io/
CPE: cpe:2.3:a:dbeaver:dbeaver:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/dbeaver/dbeaver/releases
DownloadPageUrl: https://dbeaver.io/download/
IconSource: Installer

.SYNOPSIS
    Packages DBeaver Community Edition for MECM.

.DESCRIPTION
    Downloads the latest DBeaver Community x64 setup EXE from GitHub releases,
    stages content to a versioned local folder with file-based version detection
    metadata, and creates an MECM Application with file-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The installer is an NSIS package. The /allusers flag is required for
    machine-wide installation; /currentuser installs to %LOCALAPPDATA%\DBeaver
    without elevation. DBeaver release tags have no v prefix.

    Install scope and the AI-feature toggle default to Packagers\packager-preferences.json
    under DBeaverInstallOptions, which is written by the Packager Preferences UI.
    -InstallScope and -DisableAI override the stored values.

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

.PARAMETER InstallScope
    System installs machine-wide to C:\Program Files\DBeaver via /allusers.
    User installs to %LOCALAPPDATA%\DBeaver via /currentuser. Blank uses the
    stored preference.

.PARAMETER DisableAI
    Appends -Dai.disabled=true to the installed dbeaver.ini after install.
    Omit to use the stored preference.

.PARAMETER GetLatestVersionOnly
    Queries the GitHub releases API for the latest DBeaver version, outputs
    the version string, and exits.

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
    [ValidateSet('','System','User')]
    [string]$InstallScope = '',
    [switch]$DisableAI,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


# Captured at script scope: $PSBoundParameters inside a function describes that
# function's call, not this script's, so an omitted -DisableAI must be recorded here.
$script:DisableAISpecified = $PSBoundParameters.ContainsKey('DisableAI')

Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force -ErrorAction Stop
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
# DBeaver tags have no v prefix (e.g. "25.3.5")
$GitHubApiUrl = "https://api.github.com/repos/dbeaver/dbeaver/releases/latest"

$VendorFolder = "DBeaver"
$AppFolder    = "DBeaver Community"

$BaseDownloadRoot = Join-Path $DownloadRoot "DBeaver"

# --- Functions ---


function Get-DBeaverInstallOptions {
    $resolved = [ordered]@{
        InstallScope = "System"
        DisableAI    = $false
    }

    try {
        $prefs = Get-PackagerPreferences
        if ($prefs -and $prefs.DBeaverInstallOptions) {
            $cfg = $prefs.DBeaverInstallOptions
            if ($cfg.InstallScope -in @('System','User')) { $resolved.InstallScope = [string]$cfg.InstallScope }
            if ($null -ne $cfg.DisableAI) { $resolved.DisableAI = [bool]$cfg.DisableAI }
        }
    }
    catch {
        Write-Log "Could not read DBeaverInstallOptions; using defaults: $($_.Exception.Message)" -Level WARN
    }

    if ($InstallScope -in @('System','User')) { $resolved.InstallScope = $InstallScope }
    if ($script:DisableAISpecified) { $resolved.DisableAI = [bool]$DisableAI }

    return [pscustomobject]$resolved
}


function Get-LatestDBeaverRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        # DBeaver tags have no v prefix
        $version = $release.tag_name.Trim()
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse version from GitHub release tag."
        }

        $asset = $release.assets | Where-Object { $_.name -match 'dbeaver-ce-.*-windows-x86_64\.exe$' } | Select-Object -First 1
        if (-not $asset) { throw "No Windows x64 setup EXE asset found in release." }

        Write-Log "Latest DBeaver version       : $version" -Quiet:$Quiet
        return @{ Version = $version; FileName = $asset.name; DownloadUrl = $asset.browser_download_url }
    }
    catch {
        Write-Log "Failed to get DBeaver version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDBeaver {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DBeaver Community - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $options = Get-DBeaverInstallOptions
    Write-Log "Install scope                : $($options.InstallScope)"
    Write-Log "Disable AI features          : $($options.DisableAI)"

    $releaseInfo = Get-LatestDBeaverRelease
    if (-not $releaseInfo) { throw "Could not resolve DBeaver version." }

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
        Write-Log "Downloading DBeaver..."
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
    # NSIS installer. /allusers installs machine-wide; /currentuser installs to
    # %LOCALAPPDATA%\DBeaver and needs no elevation.
    $isUserScope = ($options.InstallScope -eq 'User')

    if ($isUserScope) {
        $installArgs   = "/S /currentuser"
        $installArgList = '@(''/S'', ''/currentuser'')'
        $uninstallArgs = "/currentuser /S"
        $uninstallArgList = '@(''/currentuser'', ''/S'')'
        # Resolved in the deploying user's context, not at Stage time.
        $installRootExpr  = 'Join-Path $env:LOCALAPPDATA ''DBeaver'''
        $uninstallCommand = "%LOCALAPPDATA%\DBeaver\Uninstall.exe"
        $detectionPath    = "%LOCALAPPDATA%\DBeaver"
    }
    else {
        $installArgs   = "/allusers /S"
        $installArgList = '@(''/allusers'', ''/S'')'
        $uninstallArgs = "/S"
        $uninstallArgList = '@(''/allusers'', ''/S'')'
        $installRootExpr  = 'Join-Path $env:ProgramFiles ''DBeaver'''
        $uninstallCommand = "C:\Program Files\DBeaver\Uninstall.exe"
        $detectionPath    = "{0}\DBeaver" -f $env:ProgramFiles
    }

    $installLines = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList {0} -Wait -PassThru -NoNewWindow' -f $installArgList)
    )

    if ($options.DisableAI) {
        # DBEAVER_AI_DISABLED is DBeaver's environment-variable form of the
        # ini's -Dai.disabled=true. The variable survives upgrades (the
        # installer replaces dbeaver.ini every run) and setting it is
        # naturally idempotent. Scope follows the install scope: Machine for
        # a system deployment, User for a user-context deployment.
        $envScope = if ($isUserScope) { 'User' } else { 'Machine' }
        $installLines += @(
            'if ($proc.ExitCode -ne 0) { exit $proc.ExitCode }',
            ('[Environment]::SetEnvironmentVariable(''DBEAVER_AI_DISABLED'', ''true'', ''{0}'')' -f $envScope),
            'exit 0'
        )
    }
    else {
        $installLines += 'exit $proc.ExitCode'
    }

    $installPs1 = $installLines -join "`r`n"

    # NSIS uninstaller copies itself to temp and exits immediately.
    # Poll for dbeaver.exe removal to confirm uninstall completed.
    $uninstallPs1 = (
        ('$installRoot = {0}' -f $installRootExpr),
        '$exePath = Join-Path $installRoot ''dbeaver.exe''',
        ('$null = Start-Process -FilePath (Join-Path $installRoot ''Uninstall.exe'') -ArgumentList {0} -PassThru -NoNewWindow' -f $uninstallArgList),
        '$timeout = 120; $elapsed = 0',
        'while ((Test-Path -LiteralPath $exePath) -and $elapsed -lt $timeout) {',
        '    Start-Sleep -Seconds 2; $elapsed += 2',
        '}',
        'if (Test-Path -LiteralPath $exePath) { exit 1 }',
        'exit 0'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1 `
        -InstallBatExitCode '3010' `
        -UninstallBatExitCode '3010'

    # --- Write stage manifest ---
    $appName   = "DBeaver $version"
    $publisher = "DBeaver Corp"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : dbeaver.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = $appName
        Publisher        = $publisher
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = $installArgs
        UninstallCommand = $uninstallCommand
        UninstallArgs    = $uninstallArgs
        InstallScope     = $options.InstallScope
        DisableAI        = [bool]$options.DisableAI
        RunningProcess   = @("dbeaver")
        Detection        = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "dbeaver.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageDBeaver {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DBeaver Community - PACKAGE phase"
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
        $info = Get-LatestDBeaverRelease -Quiet
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
    Write-Log "DBeaver Community Auto-Packager starting"
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
        Invoke-StageDBeaver
    }
    elseif ($PackageOnly) {
        Invoke-PackageDBeaver
    }
    else {
        Invoke-StageDBeaver
        Invoke-PackageDBeaver
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-dbeaver'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
