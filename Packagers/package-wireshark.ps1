<#
Vendor: Wireshark Foundation
App: Wireshark (x64)
CMName: Wireshark
VendorUrl: https://www.wireshark.org/
CPE: cpe:2.3:a:wireshark:wireshark:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.wireshark.org/docs/relnotes/
DownloadPageUrl: https://www.wireshark.org/download.html
IconSource: Installer

.SYNOPSIS
    Packages Wireshark (x64) for MECM.

.DESCRIPTION
    Downloads the latest Wireshark x64 installer from the official download
    server, stages content to a versioned local folder, reads the Add/Remove
    Programs values from the installer's compiled NSIS script, and creates a
    stage manifest with file-version detection on Wireshark.exe.

    Supports two-phase operation:
      -StageOnly    Download, read installer metadata, generate content wrappers
                    and stage manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Wireshark Foundation\Wireshark\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Wireshark).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download, read installer metadata,
    generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Wireshark version string and exits.

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
$WiresharkDownloadPage = "https://www.wireshark.org/download.html"
$WiresharkWin64Root    = "https://www.wireshark.org/download/win64"

$VendorFolder = "Wireshark Foundation"
$AppFolder    = "Wireshark"

$DesktopIconSetting     = "no"
$QuickLaunchIconSetting = "no"

$BaseDownloadRoot = Join-Path $DownloadRoot "Wireshark"

# --- Functions ---


function Get-LatestWiresharkVersion {
    param([switch]$Quiet)

    Write-Log "Wireshark download page      : $WiresharkDownloadPage" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $WiresharkDownloadPage) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Wireshark download page: $WiresharkDownloadPage" }

        if ($html -match 'Stable Release:\s*([0-9]+\.[0-9]+\.[0-9]+)') {
            $v = $matches[1]
            Write-Log "Latest Wireshark version     : $v" -Quiet:$Quiet
            return $v
        }

        throw "Could not parse Stable Release version from download page."
    }
    catch {
        Write-Log "Failed to get Wireshark version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-WiresharkInstallerMetadata {
    <#
    .SYNOPSIS
        Reads the Add/Remove Programs values the installer writes from its
        compiled NSIS script.
    .DESCRIPTION
        The header decode names the uninstall key, DisplayName, DisplayVersion
        and Publisher without running the installer, so staging needs neither
        elevation nor a temporary install on the packaging machine.
    #>
    param([Parameter(Mandatory)][string]$InstallerPath)

    $analysis = Get-InstallerAnalysis -Path $InstallerPath
    $arp = @{}
    if ($analysis.PackageMetadata -and $analysis.PackageMetadata.PSObject.Properties['ArpValues'] -and $analysis.PackageMetadata.ArpValues) {
        $arp = $analysis.PackageMetadata.ArpValues
    }
    $literal = { param($v) if ($null -ne $v -and ([string]$v) -notmatch '\$') { [string]$v } else { '' } }

    return [ordered]@{
        DisplayName          = (& $literal $arp['DisplayName'])
        DisplayVersion       = (& $literal $arp['DisplayVersion'])
        Publisher            = (& $literal $arp['Publisher'])
        UninstallRegistryKey = [string]$analysis.UninstallRegistryKey
        UninstallCommand     = [string]$analysis.UninstallCommand
    }
}

# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWireshark {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Wireshark (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestWiresharkVersion
    if (-not $version) { throw "Could not resolve Wireshark version." }

    $installerFileName = "Wireshark-${version}-x64.exe"
    $downloadUrl       = "${WiresharkWin64Root}/${installerFileName}"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
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
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'', ''/desktopicon={0}'', ''/quicklaunchicon={1}'') -Wait -PassThru -NoNewWindow' -f $DesktopIconSetting, $QuickLaunchIconSetting),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        '$app = Get-ChildItem `',
        '    ''HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'',',
        '    ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'' `',
        '    -ErrorAction SilentlyContinue |',
        '    Get-ItemProperty -ErrorAction SilentlyContinue |',
        '    Where-Object { $_.DisplayName -like ''Wireshark*'' } |',
        '    Select-Object -First 1',
        '',
        'if ($app) {',
        '    $cmd = if ($app.QuietUninstallString) { $app.QuietUninstallString }',
        '           elseif ($app.UninstallString) { $app.UninstallString }',
        '           else { $null }',
        '    if ($cmd) {',
        '        if ($cmd -match ''^"([^"]+)"\s*(.*)$'') {',
        '            $proc = Start-Process $Matches[1] -ArgumentList $Matches[2] -Wait -PassThru -NoNewWindow',
        '            exit $proc.ExitCode',
        '        }',
        '        $parts = $cmd.Split(@('' ''), 2)',
        '        $proc = Start-Process $parts[0] -ArgumentList $parts[1] -Wait -PassThru -NoNewWindow',
        '        exit $proc.ExitCode',
        '    }',
        '}',
        '$fallback = Join-Path $env:ProgramFiles ''Wireshark\uninstall.exe''',
        'if (Test-Path -LiteralPath $fallback) {',
        '    $proc = Start-Process -FilePath $fallback -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Add/Remove Programs values from the compiled installer script ---
    Write-Log ""
    Write-Log "Reading ARP values from the installer's NSIS header..."
    $registryData = Get-WiresharkInstallerMetadata -InstallerPath $localExe

    $appName   = $registryData.DisplayName
    $publisher = $registryData.Publisher

    if (-not $appName)   { $appName   = "Wireshark $version x64" }
    if (-not $publisher) { $publisher = "Wireshark Foundation" }

    if ($registryData.DisplayVersion -and ($registryData.DisplayVersion -ne $version)) {
        Write-Log "Script DisplayVersion '$($registryData.DisplayVersion)' differs from download version '$version'. Detection uses download version." -Level WARN
    }

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP DisplayName              : $appName"
    Write-Log "ARP Publisher                : $publisher"
    if ($registryData.UninstallRegistryKey) { Write-Log "ARP registry key             : $($registryData.UninstallRegistryKey)" }
    Write-Log "Detection                    : Wireshark.exe file version >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S /desktopicon=no /quicklaunchicon=no"
        UninstallArgs   = "/S"
        RunningProcess  = @("Wireshark")
        Detection       = @{
            Type          = "File"
            FilePath      = "C:\Program Files\Wireshark"
            FileName      = "Wireshark.exe"
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

function Invoke-PackageWireshark {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Wireshark (x64) - PACKAGE phase"
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
    Write-Log "Detection RegKey             : $($manifest.Detection.RegistryKeyRelative)"
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
        $v = Get-LatestWiresharkVersion -Quiet
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
    Write-Log "Wireshark (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "WiresharkDownloadPage        : $WiresharkDownloadPage"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageWireshark
    }
    elseif ($PackageOnly) {
        Invoke-PackageWireshark
    }
    else {
        Invoke-StageWireshark
        Invoke-PackageWireshark
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-wireshark'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
