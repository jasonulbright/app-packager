<#
Vendor: Typora
App: Typora
CMName: Typora
VendorUrl: https://typora.io/
CPE: cpe:2.3:a:typora:typora:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://typora.io/releases/windows
DownloadPageUrl: https://typora.io/#windows
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Typora (x64) for MECM.

.DESCRIPTION
    Reads the current stable version from the vendor's Windows release channel
    page, downloads the x64 Inno Setup installer from the vendor's static
    download URL, stages content to a versioned local folder, and creates an
    MECM Application with file-version-based detection on the installed
    Typora.exe.

    Typora is commercial software: the installer is served without
    authentication, but each seat requires a purchased license key entered
    after installation. Deploy only to machines covered by owned licenses.

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
    Content is staged under: <FileServerPath>\Applications\Typora\Typora\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Typora).
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
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Typora version string and exits.

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
$ReleaseChannelUrl = "https://typora.io/releases/windows"
$ExeDownloadUrl    = "https://download.typora.io/windows/typora-setup-x64.exe"
$ExeFileName       = "typora-setup-x64.exe"

$VendorFolder = "Typora"
$AppFolder    = "Typora"

$BaseDownloadRoot = Join-Path $DownloadRoot "Typora"

$InstallDir = "{0}\Typora" -f $env:ProgramFiles

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


function Get-TyporaStableVersion {
    <#
    .SYNOPSIS
        Returns the newest Typora version listed on the Windows release channel page.
    .DESCRIPTION
        The page carries one h2 heading per release, newest first. Headings are
        collected and sorted as versions rather than trusted in document order,
        so a re-ordered or back-filled entry cannot pin the packager to an older
        release.
    #>
    param([switch]$Quiet)

    Write-Log "Typora release channel       : $ReleaseChannelUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $ReleaseChannelUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Typora release channel page: $ReleaseChannelUrl" }

        $rx = [regex]'<h2[^>]*>\s*(?<ver>\d+\.\d+\.\d+)\s*</h2>'
        $found = $rx.Matches($html) |
            ForEach-Object { $_.Groups['ver'].Value } |
            Sort-Object -Unique { [version]$_ }

        if (-not $found -or @($found).Count -lt 1) {
            throw "Could not locate any release headings on the Typora release channel page."
        }

        $version = @($found) | Sort-Object { [version]$_ } | Select-Object -Last 1

        Write-Log "Latest Typora version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Typora version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTypora {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Typora (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-TyporaStableVersion
    if (-not $version) { throw "Could not resolve Typora version." }

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $ExeFileName"
    Write-Log ""

    # --- Download ---
    # The vendor serves one static URL that always carries the current release,
    # so a stale local copy from an earlier run would silently stage the wrong
    # version; the download is unconditional and the payload's own file version
    # is checked against the release page.
    $localExe = Join-Path $BaseDownloadRoot $ExeFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Download URL                 : $ExeDownloadUrl"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = (Get-Item -LiteralPath $localExe).VersionInfo.FileVersion
    if ($fileVersion) { $fileVersion = $fileVersion.Trim() }
    if ([string]::IsNullOrWhiteSpace($fileVersion)) {
        throw "Downloaded installer carries no file version; cannot confirm the staged release."
    }
    if ($fileVersion -ne $version) {
        throw "Installer file version ($fileVersion) does not match the release page version ($version)."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $ExeFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # /ALLUSERS and an explicit /DIR pin the Inno payload to a machine-wide
    # install; without them it can land in the invoking user's profile.
    $uninstallExe = Join-Path $InstallDir "unins000.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $ExeFileName `
        -InstallArgs ("'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-', '/ALLUSERS', '/DIR={0}'" -f $InstallDir) `
        -UninstallCommand $uninstallExe -UninstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : Typora.exe"
    Write-Log "Detection version            : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Typora"
        Publisher        = "Typora"
        SoftwareVersion  = $version
        InstallerFile    = $ExeFileName
        InstallerType    = "EXE"
        InstallArgs      = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /ALLUSERS /DIR=$InstallDir"
        UninstallArgs    = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("Typora")
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "Typora.exe"
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

function Invoke-PackageTypora {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Typora (x64) - PACKAGE phase"
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
        $v = Get-TyporaStableVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Typora GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Typora (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseChannelUrl            : $ReleaseChannelUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTypora
    }
    elseif ($PackageOnly) {
        Invoke-PackageTypora
    }
    else {
        Invoke-StageTypora
        Invoke-PackageTypora
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-typora'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
