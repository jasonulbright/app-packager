<#
Vendor: Proton AG
App: Proton VPN
CMName: Proton VPN
VendorUrl: https://protonvpn.com/
CPE: cpe:2.3:a:proton:protonvpn:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://protonvpn.com/download/windows-releases.json
DownloadPageUrl: https://protonvpn.com/download-windows
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Proton VPN (x64) for MECM.

.DESCRIPTION
    Reads the vendor's Windows release feed, downloads the newest Stable x64
    Inno Setup installer, stages content to a versioned local folder, and
    creates an MECM Application with registry-based detection on the Inno
    Setup uninstall key.

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
    Content is staged under: <FileServerPath>\Applications\Proton AG\Proton VPN\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\ProtonVPN).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Proton VPN version string and exits.

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
$ReleaseFeedUrl = "https://protonvpn.com/download/windows-releases.json"

$VendorFolder = "Proton AG"
$AppFolder    = "Proton VPN"

$BaseDownloadRoot = Join-Path $DownloadRoot "ProtonVPN"

$InstallDir = "{0}\Proton\VPN" -f $env:ProgramFiles

$UninstallKeyName = "Proton VPN_is1"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A CDN edge that answers 200 with an HTML error body would otherwise
        stage as a valid-looking EXE and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestProtonVpnRelease {
    <#
    .SYNOPSIS
        Returns the newest stable Proton VPN version and its x64 installer URL.
    .DESCRIPTION
        The feed carries an EarlyAccess category alongside Stable, so only the
        Stable category is read. Its releases are sorted as versions rather
        than trusted in document order, and the URL is required to carry the
        x64 suffix so the arm64 sibling cannot be staged by accident.
    #>
    param([switch]$Quiet)

    Write-Log "Release feed URL             : $ReleaseFeedUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleaseFeedUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Proton VPN release feed: $ReleaseFeedUrl" }

        $feed = ConvertFrom-Json $json

        $stable = $feed.Categories | Where-Object { $_.Name -eq 'Stable' } | Select-Object -First 1
        if (-not $stable -or -not $stable.Releases) {
            throw "Release feed contains no Stable category."
        }

        $candidates = $stable.Releases |
            Where-Object { $_.Version -match '^\d+(\.\d+)+$' -and $_.File -and $_.File.Url -match '_x64\.exe$' }

        if (-not $candidates) {
            throw "No stable release in the feed carries an x64 installer URL."
        }

        $best = $candidates | Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1

        $fileName = ([uri]$best.File.Url).Segments[-1]

        Write-Log "Latest Proton VPN version    : $($best.Version)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            DownloadUrl = $best.File.Url
            FileName    = $fileName
        }
    }
    catch {
        Write-Log "Failed to get Proton VPN version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageProtonVpn {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Proton VPN (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestProtonVpnRelease
    if (-not $releaseInfo) { throw "Could not resolve Proton VPN version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -notlike "$version*") {
        throw "Downloaded installer reports version '$fileVersion' but the feed announced '$version'; refusing to stage a mismatched payload."
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
    # /silent is the switch the vendor's own release feed passes to this
    # installer; the Inno uninstaller takes the quieter /VERYSILENT.
    $uninstallExe = Join-Path $InstallDir "unins000.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/silent', '/norestart'" `
        -UninstallCommand $uninstallExe `
        -UninstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $UninstallKeyName

    Write-Log ""
    Write-Log "Detection RegistryKey        : $arpRegistryKey"
    Write-Log "Detection DisplayVersion     : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Proton VPN"
        Publisher        = "Proton AG"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/silent /norestart"
        UninstallArgs    = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("ProtonVPN", "ProtonVPN.Client")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $true
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

function Invoke-PackageProtonVpn {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Proton VPN (x64) - PACKAGE phase"
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
        $info = Get-LatestProtonVpnRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Proton VPN GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Proton VPN (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseFeedUrl               : $ReleaseFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageProtonVpn
    }
    elseif ($PackageOnly) {
        Invoke-PackageProtonVpn
    }
    else {
        Invoke-StageProtonVpn
        Invoke-PackageProtonVpn
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-protonvpn'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
