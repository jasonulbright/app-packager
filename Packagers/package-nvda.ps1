<#
Vendor: NV Access
App: NVDA
CMName: NVDA
VendorUrl: https://www.nvaccess.org/
ReleaseNotesUrl: https://www.nvaccess.org/post/
DownloadPageUrl: https://www.nvaccess.org/download/
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages NVDA (NonVisual Desktop Access) for MECM.

.DESCRIPTION
    Resolves the current stable release from the vendor's update-check endpoint,
    downloads the launcher EXE, stages content to a versioned local folder with
    ARP detection metadata, and creates an MECM Application with registry-based
    detection.

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
    Outputs only the latest available NVDA version string and exits.

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
$UpdateCheckUrl = "https://www.nvaccess.org/nvdaUpdateCheck?versionType=stable&autoCheck=false"

$VendorFolder = "NV Access"
$AppFolder    = "NVDA"

$BaseDownloadRoot = Join-Path $DownloadRoot "NVDA"

$ArpKeyRelative = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\NVDA"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A download redirect that lands on an HTML error page would otherwise
        stage as a valid-looking installer and fail at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestNvdaRelease {
    <#
    .SYNOPSIS
        Returns the current stable NVDA version and its launcher URL.
    .DESCRIPTION
        The vendor's update-check endpoint answers with a colon-delimited key
        list rather than JSON. Querying versionType=stable keeps beta and alpha
        channels out; the launcher URL is taken from the response instead of
        being composed, so a changed release layout surfaces as a fetch failure
        rather than a wrong download.
    #>
    param([switch]$Quiet)

    Write-Log "NVDA update check endpoint   : $UpdateCheckUrl" -Quiet:$Quiet

    try {
        $lines = curl.exe -L --fail --silent --show-error $UpdateCheckUrl
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the NVDA update-check endpoint." }

        $fields = @{}
        foreach ($line in $lines) {
            if ($line -match '^\s*(?<k>[A-Za-z]+)\s*:\s*(?<v>.+?)\s*$') {
                $fields[$matches['k']] = $matches['v']
            }
        }

        $version = $fields['version']
        $url     = $fields['launcherUrl']

        if ([string]::IsNullOrWhiteSpace($version)) { throw "Update-check response carried no version field." }
        if ([string]::IsNullOrWhiteSpace($url))     { throw "Update-check response carried no launcherUrl field." }
        if ($version -notmatch '^\d{4}\.\d+(\.\d+)?$') {
            throw "Update-check version is not a stable NVDA release string: $version"
        }
        if ($url -notmatch '(?i)^https://[^/]*nvaccess\.org/.*\.exe$') {
            throw "Update-check launcherUrl is not an nvaccess.org executable: $url"
        }

        Write-Log "Latest NVDA version          : $version" -Quiet:$Quiet
        Write-Log "Launcher URL                 : $url" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = [System.IO.Path]::GetFileName(([uri]$url).AbsolutePath)
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get NVDA version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function New-NvdaUninstallContent {
    <#
    .SYNOPSIS
        Returns uninstall.ps1 content that runs the registered NVDA uninstaller.
    .DESCRIPTION
        NVDA writes its uninstaller into the install directory rather than
        shipping one in content, and the launcher is a 32-bit process, so the
        ARP key lands in the 32-bit registry view on x64 machines. Both views
        are read so a future 64-bit build keeps uninstalling.
    #>

    return @'
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\NVDA',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\NVDA'
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
$proc = Start-Process -FilePath $exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageNvda {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVDA - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestNvdaRelease
    if (-not $releaseInfo) { throw "Could not resolve NVDA version." }

    $version       = $releaseInfo.Version
    $installerName = $releaseInfo.FileName

    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading launcher (~65 MB)..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

    # The installer writes DisplayVersion from its own four-field build version,
    # which the launcher carries as its FileVersion; the two-field release string
    # never appears in the registry.
    $fileVersion = (Get-Item -LiteralPath $localExe).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($fileVersion)) {
        throw "Launcher carries no FileVersion; cannot derive the ARP DisplayVersion."
    }
    $fileVersion = $fileVersion.Trim()

    Write-Log "Launcher FileVersion         : $fileVersion"
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
    Write-Log "ARP DisplayVersion           : $fileVersion"
    Write-Log ""

    # --- Generate content wrappers ---
    # --install-silent installs per-machine without starting NVDA afterwards.
    $wrappers = New-ExeWrapperContent `
        -InstallerFileName $installerName `
        -InstallArgs "'--install-silent'" `
        -UninstallCommand 'unused'

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content (New-NvdaUninstallContent)

    # --- Write stage manifest ---
    # The launcher is 32-bit, so its ARP key lands in the 32-bit view on x64
    # machines; the 64-bit clause covers a future 64-bit build.
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "NVDA"
        Publisher       = "NV Access"
        SoftwareVersion = $version
        InstallerFile   = $installerName
        InstallerType   = "EXE"
        InstallArgs     = "--install-silent"
        UninstallArgs   = "/S"
        RunningProcess  = @("nvda")
        Detection       = @{
            Type      = "Compound"
            Connector = "Or"
            Clauses   = @(
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $ArpKeyRelative
                    ValueName           = "DisplayVersion"
                    ExpectedValue       = $fileVersion
                    Is64Bit             = $false
                },
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $ArpKeyRelative
                    ValueName           = "DisplayVersion"
                    ExpectedValue       = $fileVersion
                    Is64Bit             = $true
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

function Invoke-PackageNvda {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVDA - PACKAGE phase"
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
        $info = Get-LatestNvdaRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("NVDA GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "NVDA Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "UpdateCheckUrl               : $UpdateCheckUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageNvda
    }
    elseif ($PackageOnly) {
        Invoke-PackageNvda
    }
    else {
        Invoke-StageNvda
        Invoke-PackageNvda
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-nvda'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
