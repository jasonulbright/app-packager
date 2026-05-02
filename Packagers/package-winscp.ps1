<#
Vendor: Martin Prikryl
App: WinSCP (x64)
CMName: WinSCP
VendorUrl: https://winscp.net/
CPE: cpe:2.3:a:winscp:winscp:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://winscp.net/eng/docs/history
DownloadPageUrl: https://winscp.net/eng/download.php

.SYNOPSIS
    Packages WinSCP (x64) for MECM.

.DESCRIPTION
    Downloads the latest WinSCP installer from winscp.net, stages content to a
    versioned local folder, temporarily installs locally to extract registry
    metadata (DisplayName, Publisher, uninstall key), creates stage manifest with
    registry-based detection (DisplayVersion), then uninstalls from the packaging
    machine.

    Supports two-phase operation:
      -StageOnly    Download, temp install for registry discovery, generate content
                    wrappers and stage manifest, then uninstall
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\WinSCP\WinSCP\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\WinSCP).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download, temp install for discovery, generate
    content wrappers and stage manifest, then uninstall.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available WinSCP version string and exits.

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
$VendorFolder = "WinSCP"
$AppFolder    = "WinSCP"

$BaseDownloadRoot = Join-Path $DownloadRoot "WinSCP"

# --- Functions ---


function Get-LatestWinSCPVersion {
    param([switch]$Quiet)

    $url = "https://winscp.net/eng/downloads.php"
    Write-Log "WinSCP downloads page        : $url" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $url) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP downloads page: $url" }

        $version = $null
        if ($html -match 'Download\s+WinSCP\s+([0-9]+\.[0-9]+\.[0-9]+)') {
            $version = $matches[1]
        }
        elseif ($html -match 'WinSCP-([0-9]+\.[0-9]+\.[0-9]+)-Setup\.exe') {
            $version = $matches[1]
        }

        if (-not $version) {
            throw "Could not parse latest WinSCP version from downloads page."
        }

        Write-Log "Latest WinSCP version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get WinSCP version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-WinSCPDownloadUrl {
    param([Parameter(Mandatory)][string]$Version)

    # SourceForge hosts WinSCP releases; URL redirects to a mirror
    $url = "https://sourceforge.net/projects/winscp/files/WinSCP/$Version/WinSCP-$Version-Setup.exe/download"
    Write-Log "SourceForge download URL     : $url"
    return $url
}

function Test-DownloadedInstaller {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -lt 1MB) { return $false }

    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 64
        [void]$fs.Read($buf, 0, $buf.Length)
        $head = [System.Text.Encoding]::ASCII.GetString($buf)
        if ($head -match '<!DOCTYPE|<html|<HTML') { return $false }
    }
    finally { $fs.Dispose() }

    return $true
}

function Install-WinSCPForDiscovery {
    param([Parameter(Mandatory)][string]$InstallerPath)

    Write-Log "Installing WinSCP locally for registry discovery..."
    Write-Log "Installer                    : $InstallerPath"

    $p = Start-Process -FilePath $InstallerPath -ArgumentList "/VERYSILENT /NORESTART /ALLUSERS" -Wait -PassThru -ErrorAction Stop
    Write-Log "Install exit code            : $($p.ExitCode)"

    if ($p.ExitCode -ne 0) {
        throw "Installer returned non-zero exit code: $($p.ExitCode)"
    }
}

function Find-WinSCPUninstallEntry {
    param([Parameter(Mandatory)][string]$ExpectedVersion)

    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $candidates = @()
    foreach ($p in $paths) {
        $items = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue
        foreach ($it in $items) {
            $dn = $it.DisplayName
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }
            if ($dn -match '^WinSCP') {
                $candidates += [pscustomobject]@{
                    KeyName              = ($it.PSPath -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE\\', 'HKLM:\')
                    DisplayName          = $dn
                    DisplayVersion       = $it.DisplayVersion
                    Publisher            = $it.Publisher
                    UninstallString      = $it.UninstallString
                    QuietUninstallString = $it.QuietUninstallString
                }
            }
        }
    }

    Write-Log "WinSCP uninstall candidates  : $($candidates.Count)"

    $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
    if ($match) { return $match }

    return ($candidates | Select-Object -First 1)
}

function Uninstall-WinSCPFromDiscovery {
    param(
        [string]$QuietUninstallString,
        [string]$UninstallString
    )

    $cmd = $QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($cmd)) { $cmd = $UninstallString }

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        Write-Log "No uninstall command discovered; leaving WinSCP installed on packaging machine." -Level WARN
        return
    }

    Write-Log "Uninstalling discovery install..."
    Write-Log "Uninstall command            : $cmd"

    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru -ErrorAction Stop
    Write-Log "Uninstall exit code          : $($p.ExitCode)"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWinSCP {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinSCP (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestWinSCPVersion
    if (-not $version) { throw "Could not resolve WinSCP version." }

    if ($version -notmatch '^\d+\.\d+\.\d+$') {
        throw "Parsed version '$version' does not look like x.y.z"
    }

    $installerFileName = "WinSCP-$version-Setup.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download from SourceForge ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        $downloadUrl = Get-WinSCPDownloadUrl -Version $version

        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe -ExtraCurlArgs @('-L', '--max-redirs', '10')

        if (-not (Test-DownloadedInstaller -Path $localExe)) {
            try { Remove-Item -LiteralPath $localExe -Force -ErrorAction SilentlyContinue } catch {}
            throw "Downloaded file did not validate as installer (too small or HTML content)."
        }

        Write-Log ("Download validated             ({0} bytes)" -f (Get-Item -LiteralPath $localExe).Length)
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
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/VERYSILENT'', ''/NORESTART'', ''/ALLUSERS'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        '$uninstaller = Join-Path $env:ProgramFiles ''WinSCP\unins000.exe''',
        'if (-not (Test-Path -LiteralPath $uninstaller)) {',
        '    $uninstaller = Join-Path ${env:ProgramFiles(x86)} ''WinSCP\unins000.exe''',
        '}',
        'if (Test-Path -LiteralPath $uninstaller) {',
        '    $proc = Start-Process -FilePath $uninstaller -ArgumentList @(''/VERYSILENT'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}',
        'Write-Warning ''WinSCP uninstall executable not found.''',
        'exit 0'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Temp install for registry discovery ---
    Write-Log ""
    Install-WinSCPForDiscovery -InstallerPath $localExe

    $uninstallEntry = Find-WinSCPUninstallEntry -ExpectedVersion $version
    if ($null -eq $uninstallEntry) {
        throw "Could not find any WinSCP uninstall entry after install."
    }

    Write-Log "Registry DisplayName         : $($uninstallEntry.DisplayName)"
    Write-Log "Registry DisplayVersion      : $($uninstallEntry.DisplayVersion)"
    Write-Log "Registry Publisher           : $($uninstallEntry.Publisher)"
    Write-Log "Registry Key                 : $($uninstallEntry.KeyName)"

    $regRelative = ($uninstallEntry.KeyName -replace '^HKLM:\\', '')
    if ([string]::IsNullOrWhiteSpace($regRelative)) {
        throw "Failed to compute registry relative key path."
    }

    $publisher = $uninstallEntry.Publisher
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Martin Prikryl" }

    $appName = $uninstallEntry.DisplayName

    # Uninstall discovery install
    Write-Log ""
    Uninstall-WinSCPFromDiscovery `
        -QuietUninstallString $uninstallEntry.QuietUninstallString `
        -UninstallString $uninstallEntry.UninstallString

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection RegKey             : $regRelative"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /ALLUSERS"
        UninstallArgs   = "/VERYSILENT /NORESTART"
        RunningProcess  = @("WinSCP")
        Detection       = @{
            Type          = "File"
            FilePath      = "C:\Program Files (x86)\WinSCP"
            FileName      = "WinSCP.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
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

function Invoke-PackageWinSCP {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinSCP (x64) - PACKAGE phase"
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
    Write-Log "Detection RegKey             : $($manifest.Detection.RegistryKeyRelative)"
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
        $v = Get-LatestWinSCPVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
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
    Write-Log "WinSCP (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageWinSCP
    }
    elseif ($PackageOnly) {
        Invoke-PackageWinSCP
    }
    else {
        Invoke-StageWinSCP
        Invoke-PackageWinSCP
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
