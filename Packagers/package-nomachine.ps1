<#
Vendor: NoMachine
App: NoMachine
CMName: NoMachine
VendorUrl: https://www.nomachine.com/
CPE: cpe:2.3:a:nomachine:nomachine:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.nomachine.com/history
DownloadPageUrl: https://download.nomachine.com/download/?id=41&platform=windows
UpdateCadenceDays: 60

.SYNOPSIS
    Packages NoMachine (x64) for MECM.

.DESCRIPTION
    Resolves the current Windows x64 package from the vendor download page,
    downloads the Inno Setup installer, stages content to a versioned local
    folder with ARP detection metadata, and creates an MECM Application with
    registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection, write manifest
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
    Outputs only the latest available NoMachine version string and exits.

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
$DownloadPageUrl = "https://download.nomachine.com/download/?id=41&platform=windows"

$VendorFolder = "NoMachine"
$AppFolder    = "NoMachine"

$BaseDownloadRoot = Join-Path $DownloadRoot "NoMachine"

# Inno Setup appends _is1 to the AppId; NoMachine's AppId is the product name.
$ArpKeyRelative = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\NoMachine_is1"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A mirror redirect that lands on an HTML error page would otherwise stage
        as a valid-looking installer and fail at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestNoMachineRelease {
    <#
    .SYNOPSIS
        Returns the current NoMachine Windows x64 package and its download URL.
    .DESCRIPTION
        The vendor publishes no version feed, so the download page is parsed for
        the x64 installer filename. The page links a rotating mirror host; the
        URL is rebuilt against download.nomachine.com so a mirror rename does not
        break staging. The filename's trailing _<n> is the vendor's package
        revision and is not part of the product version.
    #>
    param([switch]$Quiet)

    Write-Log "NoMachine download page      : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the NoMachine download page." }

        $rx = [regex]'(?i)(?<file>nomachine-personal-edition_(?<ver>\d+\.\d+\.\d+)_(?<rev>\d+)_x64\.exe)'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate an x64 installer filename on the download page."
        }

        $candidates = foreach ($m in $rxMatches) {
            [pscustomobject]@{
                FileName = $m.Groups['file'].Value
                Version  = $m.Groups['ver'].Value
                Revision = [int]$m.Groups['rev'].Value
                Series   = ($m.Groups['ver'].Value -split '\.')[0..1] -join '.'
            }
        }

        $best = $candidates |
            Sort-Object @{ Expression = { [version]$_.Version } }, Revision -Descending |
            Select-Object -First 1

        $url = "https://download.nomachine.com/download/$($best.Series)/Windows/$($best.FileName)"

        Write-Log "Latest NoMachine version     : $($best.Version)" -Quiet:$Quiet
        Write-Log "Package revision             : $($best.Revision)" -Quiet:$Quiet
        Write-Log "Installer asset              : $($best.FileName)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            FileName    = $best.FileName
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get NoMachine version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function New-NoMachineUninstallContent {
    <#
    .SYNOPSIS
        Returns uninstall.ps1 content that runs the registered Inno uninstaller.
    .DESCRIPTION
        Inno Setup names the uninstaller unins###.exe by install order and writes
        it into the install directory, so the ARP UninstallString is the only
        value that names the right one. Both registry views are read because the
        32-bit and 64-bit packages share one ARP key name.
    #>

    return @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\NoMachine_is1',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\NoMachine_is1'
)
$cmd = $null
foreach ($key in $keys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $hit = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($hit -and $hit.UninstallString) { $cmd = $hit.UninstallString; break }
}
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageNoMachine {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NoMachine (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestNoMachineRelease
    if (-not $releaseInfo) { throw "Could not resolve NoMachine version." }

    $version       = $releaseInfo.Version
    $installerName = $releaseInfo.FileName

    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading installer (~100 MB, this can take several minutes)..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

    # Inno writes DisplayVersion from AppVersion, which the setup binary carries
    # as its ProductVersion; the filename's package revision is not part of it.
    $productVersion = (Get-Item -LiteralPath $localExe).VersionInfo.ProductVersion
    if ([string]::IsNullOrWhiteSpace($productVersion)) {
        throw "Installer carries no ProductVersion; cannot derive the ARP DisplayVersion."
    }
    $productVersion = $productVersion.Trim()

    Write-Log "Installer ProductVersion     : $productVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    Write-Log "ARP RegistryKey              : $ArpKeyRelative"
    Write-Log "ARP DisplayVersion           : $productVersion"
    Write-Log ""

    # --- Generate content wrappers ---
    # The installer loads a display driver and asks for a restart; /NORESTART
    # leaves the reboot to the deployment.
    $wrappers = New-ExeWrapperContent `
        -InstallerFileName $installerName `
        -InstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'" `
        -UninstallCommand 'unused'

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content (New-NoMachineUninstallContent)

    # --- Write stage manifest ---
    # The x64 package is built from a 32-bit Inno setup binary, so which registry
    # view holds the ARP key depends on the installer's 64-bit install mode; both
    # views are matched.
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "NoMachine"
        Publisher       = "NoMachine S.a.r.l."
        SoftwareVersion = $version
        InstallerFile   = $installerName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("nxplayer", "nxservice", "nxtray")
        Detection       = @{
            Type      = "Compound"
            Connector = "Or"
            Clauses   = @(
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $ArpKeyRelative
                    ValueName           = "DisplayVersion"
                    ExpectedValue       = $productVersion
                    Is64Bit             = $true
                },
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $ArpKeyRelative
                    ValueName           = "DisplayVersion"
                    ExpectedValue       = $productVersion
                    Is64Bit             = $false
                }
            )
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

function Invoke-PackageNoMachine {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NoMachine (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $ArpKeyRelative"
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
        $info = Get-LatestNoMachineRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("NoMachine GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NoMachine (x64) Auto-Packager starting"
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
        Invoke-StageNoMachine
    }
    elseif ($PackageOnly) {
        Invoke-PackageNoMachine
    }
    else {
        Invoke-StageNoMachine
        Invoke-PackageNoMachine
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-nomachine'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
