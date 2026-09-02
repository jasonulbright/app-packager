<#
Vendor: TeamSpeak Systems GmbH
App: TeamSpeak 3 Client
CMName: TeamSpeak 3 Client
VendorUrl: https://teamspeak.com/
CPE: cpe:2.3:a:teamspeak:teamspeak3_client:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://teamspeak.com/en/downloads/
DownloadPageUrl: https://teamspeak.com/en/downloads/
UpdateCadenceDays: 180

.SYNOPSIS
    Packages the TeamSpeak 3 Client (x64) for MECM.

.DESCRIPTION
    Resolves the current stable client release from the vendor download page,
    downloads the win64 setup, stages content to a versioned local folder, and
    creates an MECM Application with file-version detection.

    The vendor's download page offers a TeamSpeak 6 client only from its
    pre-release path, so this packager tracks the release path, which currently
    carries the TeamSpeak 3 client.

    The payload is an NSIS setup: /S installs silently and skips the optional
    third-party offer page, and /D sets the target directory and must stay the
    last argument.

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
    Content is staged under: <FileServerPath>\Applications\TeamSpeak Systems GmbH\TeamSpeak 3 Client\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TeamSpeak3Client).
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
    create MECM application with file-version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available TeamSpeak 3 Client version string and exits.

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
$DownloadPageUrl = "https://teamspeak.com/en/downloads/"

$VendorFolder = "TeamSpeak Systems GmbH"
$AppFolder    = "TeamSpeak 3 Client"

$BaseDownloadRoot = Join-Path $DownloadRoot "TeamSpeak3Client"

$InstallDir  = "{0}\TeamSpeak 3 Client" -f $env:ProgramFiles
$ClientExe   = "ts3client_win64.exe"

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


function ConvertTo-FourPartVersion {
    <#
    .SYNOPSIS
        Returns a version string padded to four numeric parts.
    .DESCRIPTION
        The client binary stamps a four-part FileVersion while the release
        number carries three, and a file-version detection clause compares the
        two literally.
    #>
    param([Parameter(Mandatory)][string]$Version)

    $v = [version]($Version -replace '[,\s]+', '.')
    $build    = if ($v.Build    -lt 0) { 0 } else { $v.Build }
    $revision = if ($v.Revision -lt 0) { 0 } else { $v.Revision }
    return ("{0}.{1}.{2}.{3}" -f $v.Major, $v.Minor, $build, $revision)
}


function Get-LatestTeamSpeak3ClientRelease {
    <#
    .SYNOPSIS
        Returns the current stable TeamSpeak 3 Client version and its setup URL.
    .DESCRIPTION
        The download page also links a pre-release client from a separate
        pre_releases path; the pattern is anchored on the releases path so a
        beta never wins. The page lists the same asset several times, so the
        highest version is selected rather than the first link.
    #>
    param([switch]$Quiet)

    Write-Log "TeamSpeak download page      : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch TeamSpeak download page: $DownloadPageUrl" }

        $rx = [regex]'https://files\.teamspeak-services\.com/releases/client/(?<ver>\d+(?:\.\d+)+)/TeamSpeak3-Client-win64-(?<ver2>\d+(?:\.\d+)+)\.exe'
        $rxMatches = $rx.Matches($html)
        if (-not $rxMatches -or $rxMatches.Count -lt 1) {
            throw "Could not locate a win64 client setup link on the download page."
        }

        $candidates = foreach ($m in $rxMatches) {
            if ($m.Groups['ver'].Value -ne $m.Groups['ver2'].Value) { continue }
            [pscustomobject]@{
                Version = $m.Groups['ver'].Value
                Url     = $m.Value
            }
        }
        if (-not $candidates) { throw "No client setup link had a matching path and filename version." }

        $best = $candidates | Sort-Object { [version]$_.Version } | Select-Object -Last 1

        Write-Log "Latest TeamSpeak 3 Client    : $($best.Version)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            FileName    = "TeamSpeak3-Client-win64-$($best.Version).exe"
            DownloadUrl = $best.Url
        }
    }
    catch {
        Write-Log "Failed to get TeamSpeak 3 Client version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeamSpeak3Client {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamSpeak 3 Client (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestTeamSpeak3ClientRelease
    if (-not $releaseInfo) { throw "Could not resolve TeamSpeak 3 Client version." }

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

    # The setup binary carries the same FileVersion the installed client
    # binary is stamped with, so the detection value comes from the payload
    # rather than from the release number on the page.
    $stampedVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($stampedVersion)) {
        throw "Setup binary carries no file version; cannot derive the detection value."
    }
    $detectionVersion = ConvertTo-FourPartVersion -Version $stampedVersion.Trim()

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
    # NSIS requires /D last, unquoted, and takes the rest of the command line
    # as the path, so a directory with spaces must not be quoted here.
    $installArgs = "'/S', '/D=$InstallDir'"
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand (Join-Path $InstallDir 'Uninstall.exe') `
        -UninstallArgs "'/S'"

    # A missing uninstaller means the product is already gone; the generated
    # wrapper would fail the deployment on a Start-Process path error.
    $uninstallContent = @"
`$exe = '$(Join-Path $InstallDir 'Uninstall.exe')'
if (-not (Test-Path -LiteralPath `$exe)) { exit 0 }
`$proc = Start-Process -FilePath `$exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Install directory            : $InstallDir"
    Write-Log "Detection file               : $ClientExe"
    Write-Log "Detection version            : $detectionVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "TeamSpeak 3 Client"
        Publisher       = "TeamSpeak Systems GmbH"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S /D=$InstallDir"
        UninstallArgs   = "/S"
        RunningProcess  = @("ts3client_win64")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = $ClientExe
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $detectionVersion
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

function Invoke-PackageTeamSpeak3Client {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamSpeak 3 Client (x64) - PACKAGE phase"
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
    Write-Log "Detection File               : $($manifest.Detection.FilePath)\$($manifest.Detection.FileName)"
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
        $info = Get-LatestTeamSpeak3ClientRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("TeamSpeak 3 Client GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamSpeak 3 Client (x64) Auto-Packager starting"
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
        Invoke-StageTeamSpeak3Client
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeamSpeak3Client
    }
    else {
        Invoke-StageTeamSpeak3Client
        Invoke-PackageTeamSpeak3Client
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-teamspeak3client'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
