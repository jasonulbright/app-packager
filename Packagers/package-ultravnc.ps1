<#
Vendor: uvnc bvba
App: UltraVNC
CMName: UltraVNC
VendorUrl: https://uvnc.com/
CPE: cpe:2.3:a:uvnc:ultravnc:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://uvnc.com/downloads/ultravnc.html
DownloadPageUrl: https://uvnc.com/downloads/ultravnc.html
UpdateCadenceDays: 90

.SYNOPSIS
    Packages UltraVNC (x64) for MECM.

.DESCRIPTION
    Resolves the newest UltraVNC release from the vendor download index,
    follows the download handler to the mirror URL, downloads the x64 Inno
    Setup installer, stages content to a versioned local folder, and creates an
    MECM Application with registry-based detection on the Inno Setup uninstall
    key.

    The staged install selects the Viewer and Server components but registers
    no service and starts nothing: an UltraVNC server that is reachable before
    a password is set in ultravnc.ini accepts unauthenticated sessions.
    Configure ultravnc.ini and register the service separately.

    The vendor also publishes an MSI build of the same release; it is an MSI
    wrapper around this installer whose ProductName carries the wrapper
    vendor's unregistered-copy notice, so the native installer is staged
    instead.

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
    Content is staged under: <FileServerPath>\Applications\uvnc bvba\UltraVNC\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\UltraVNC).
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
    Outputs only the latest available UltraVNC version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$SiteRoot        = "https://uvnc.com"
$DownloadIndexUrl = "$SiteRoot/downloads/ultravnc.html"

$VendorFolder = "uvnc bvba"
$AppFolder    = "UltraVNC"

$BaseDownloadRoot = Join-Path $DownloadRoot "UltraVNC"

$InstallDir = "{0}\uvnc bvba\UltraVNC" -f $env:ProgramFiles

# The Inno script sets AppID to a fixed literal that is neither the product
# name nor the version, so the uninstall key is stable across releases; the x64
# build installs in 64-bit mode, so the key is not redirected.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Ultravnc2_is1"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor's download handler answers 200 with an HTML redirect stub,
        so a mirror that has dropped the file would otherwise stage as a
        valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestUltraVncRelease {
    <#
    .SYNOPSIS
        Returns the newest UltraVNC release version and its x64 installer URL.
    .DESCRIPTION
        Three vendor pages are chained: the download index lists one detail page
        per release with the version in the slug, the detail page links the
        per-file summary pages, and the matching download handler answers with
        an HTML stub whose script assignment carries the mirror URL. Index
        entries are sorted as versions rather than trusted in document order.
    #>
    param([switch]$Quiet)

    Write-Log "UltraVNC download index      : $DownloadIndexUrl" -Quiet:$Quiet

    try {
        $indexHtml = (curl.exe -L --fail --silent --show-error $DownloadIndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch UltraVNC download index: $DownloadIndexUrl" }

        $indexRx = [regex]'href\s*=\s*"(?<href>/downloads/ultravnc/\d+-ultravnc-(?<a>\d+)-(?<b>\d+)-(?<c>\d+)-(?<d>\d+)\.html)"'
        $entries = foreach ($m in $indexRx.Matches($indexHtml)) {
            [pscustomobject]@{
                Href    = $m.Groups['href'].Value
                Version = "{0}.{1}.{2}.{3}" -f $m.Groups['a'].Value, $m.Groups['b'].Value, $m.Groups['c'].Value, $m.Groups['d'].Value
            }
        }

        if (-not $entries -or @($entries).Count -lt 1) {
            throw "Could not locate any release entries on the UltraVNC download index."
        }

        $best = @($entries) | Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1
        $version = $best.Version
        $compact = $version -replace '\.', ''

        $detailUrl = $SiteRoot + $best.Href
        Write-Log "Release detail page          : $detailUrl" -Quiet:$Quiet

        $detailHtml = (curl.exe -L --fail --silent --show-error $detailUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch UltraVNC release detail page: $detailUrl" }

        $summaryRx = [regex]('href\s*=\s*"(?<href>/component/jdownloads/summary/[^"]*?' + $compact + '-x64-setup\.html[^"]*)"')
        $summaryMatch = $summaryRx.Match($detailHtml)
        if (-not $summaryMatch.Success) {
            throw "Could not locate the x64 setup entry for UltraVNC $version on $detailUrl"
        }

        # The summary page and its download handler differ only in this segment.
        $sendUrl = $SiteRoot + ($summaryMatch.Groups['href'].Value -replace '/jdownloads/summary/', '/jdownloads/send/')

        $stub = (curl.exe -L --fail --silent --show-error $sendUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch UltraVNC download handler: $sendUrl" }

        $urlMatch = [regex]::Match($stub, "location\.href\s*=\s*'(?<url>https?://[^']+\.exe)'")
        if (-not $urlMatch.Success) {
            throw "UltraVNC download handler did not return an installer URL: $sendUrl"
        }

        $downloadUrl = $urlMatch.Groups['url'].Value

        Write-Log "Latest UltraVNC version      : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = (Split-Path -Leaf ($downloadUrl -split '\?')[0])
        }
    }
    catch {
        Write-Log "Failed to get UltraVNC version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageUltraVnc {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UltraVNC (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestUltraVncRelease
    if (-not $releaseInfo) { throw "Could not resolve UltraVNC version." }

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
    # An empty task list deselects the service-install and service-start tasks
    # that the payload otherwise pre-checks.
    $uninstallExe = Join-Path $InstallDir "unins000.exe"
    $installArgs = "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-', '/COMPONENTS=UltraVNC_Viewer,UltraVNC_Server', '/TASKS='"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand $uninstallExe -UninstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "UltraVNC"
        Publisher        = "uvnc bvba"
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /COMPONENTS=UltraVNC_Viewer,UltraVNC_Server /TASKS="
        UninstallArgs    = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("vncviewer", "winvnc")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
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

function Invoke-PackageUltraVnc {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UltraVNC (x64) - PACKAGE phase"
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
        $info = Get-LatestUltraVncRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("UltraVNC GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "UltraVNC (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadIndexUrl             : $DownloadIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageUltraVnc
    }
    elseif ($PackageOnly) {
        Invoke-PackageUltraVnc
    }
    else {
        Invoke-StageUltraVnc
        Invoke-PackageUltraVnc
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-ultravnc'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
