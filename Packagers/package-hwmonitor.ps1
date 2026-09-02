<#
Vendor: CPUID
App: HWMonitor
CMName: HWMonitor
VendorUrl: https://www.cpuid.com/softwares/hwmonitor.html
CPE: cpe:2.3:a:cpuid:hwmonitor:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.cpuid.com/softwares/hwmonitor.html
DownloadPageUrl: https://www.cpuid.com/softwares/hwmonitor.html
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages HWMonitor (x64) for MECM.

.DESCRIPTION
    Reads the current version from the vendor product page, downloads the
    matching Inno Setup installer, stages content to a versioned local folder,
    and creates an MECM Application with file-existence detection on the
    installed HWMonitor_x64.exe.

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
    Content is staged under: <FileServerPath>\Applications\CPUID\HWMonitor\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available HWMonitor version string and exits.

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
$DownloadPageUrl = "https://www.cpuid.com/softwares/hwmonitor.html"
$DownloadBase    = "https://www.cpuid.com/downloads/hwmonitor"

$VendorFolder = "CPUID"
$AppFolder    = "HWMonitor"

$BaseDownloadRoot = Join-Path $DownloadRoot "HWMonitor"

$InstallDir = "{0}\CPUID\HWMonitor" -f $env:ProgramFiles

# --- Functions ---


function Assert-HwMonitorPayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor download host answers 200 with an HTML interstitial for
        unresolved links, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Resolve-HwMonitorInstallerUrl {
    <#
    .SYNOPSIS
        Returns the direct installer URL for a HWMonitor version.
    .DESCRIPTION
        www.cpuid.com/downloads/... serves an HTML interstitial, not the
        binary; the real payload lives on download.cpuid.com and is linked
        from that page. The link is read from the page so a host change is
        followed; the known host is the fallback.
    #>
    param(
        [Parameter(Mandatory)][string]$FileName,
        [switch]$Quiet
    )

    $fallback = "https://download.cpuid.com/hwmonitor/$FileName"
    $interstitial = "$DownloadBase/$FileName"

    try {
        $html = (curl.exe -L --fail --silent --show-error $interstitial) -join "`n"
        if ($LASTEXITCODE -eq 0) {
            $m = [regex]::Match($html, 'href\s*=\s*"(?<url>https?://[^"]*?' + [regex]::Escape($FileName) + ')"')
            if ($m.Success -and $m.Groups['url'].Value -ne $interstitial) {
                Write-Log "Resolved installer URL       : $($m.Groups['url'].Value)" -Quiet:$Quiet
                return $m.Groups['url'].Value
            }
        }
    }
    catch { }

    Write-Log "Falling back to known host   : $fallback" -Level WARN -Quiet:$Quiet
    return $fallback
}


function Get-LatestHwMonitorVersion {
    <#
    .SYNOPSIS
        Returns the highest HWMonitor version linked on the vendor product page.
    .DESCRIPTION
        The page links previous releases alongside the current one, so the
        highest [version] wins rather than the first match in document order.
    #>
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch HWMonitor download page: $DownloadPageUrl" }

        # Release files are hwmonitor_<version>.exe; the version carries two or
        # three parts (1.57, 1.65.1).
        $rx = [regex]'hwmonitor_(?<ver>\d+\.\d+(\.\d+)?)\.exe'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate any installer links on the download page."
        }

        $best = $rxMatches |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ } |
            Select-Object -Last 1

        if ([string]::IsNullOrWhiteSpace($best)) { throw "Version selection produced an empty result." }

        Write-Log "Latest HWMonitor version     : $best" -Quiet:$Quiet
        return $best
    }
    catch {
        Write-Log "Failed to get HWMonitor version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageHwMonitor {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HWMonitor (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestHwMonitorVersion
    if (-not $version) { throw "Could not resolve HWMonitor version." }

    $installerFileName = "hwmonitor_$version.exe"
    $downloadUrl       = Resolve-HwMonitorInstallerUrl -FileName $installerFileName

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

    Assert-HwMonitorPayloadIsExecutable -Path $localExe

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
        -InstallArgs "'/VERYSILENT', '/NORESTART'" `
        -UninstallCommand 'unused'

    # Inno Setup names the uninstaller unins###.exe by install order, so the
    # ARP UninstallString is the only value that names the right one.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
$cmd = $null
foreach ($root in $keys) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $hit = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.DisplayName -like 'CPUID HWMonitor*' -or $_.DisplayName -like 'HWMonitor*' } |
        Select-Object -First 1
    if ($hit -and $hit.UninstallString) { $cmd = $hit.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/NORESTART') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "HWMonitor (x64)"
    $publisher = "CPUID"

    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : HWMonitor_x64.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART"
        UninstallArgs   = "/VERYSILENT /NORESTART"
        RunningProcess  = @("HWMonitor", "HWMonitor_x64")
        Detection       = @{
            Type         = "File"
            FilePath     = $InstallDir
            FileName     = "HWMonitor_x64.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
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

function Invoke-PackageHwMonitor {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HWMonitor (x64) - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        $v = Get-LatestHwMonitorVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("HWMonitor GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HWMonitor (x64) Auto-Packager starting"
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
        Invoke-StageHwMonitor
    }
    elseif ($PackageOnly) {
        Invoke-PackageHwMonitor
    }
    else {
        Invoke-StageHwMonitor
        Invoke-PackageHwMonitor
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-hwmonitor'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
