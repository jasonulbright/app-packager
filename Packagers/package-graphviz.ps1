<#
Vendor: Graphviz
App: Graphviz
CMName: Graphviz
VendorUrl: https://graphviz.org/
CPE: cpe:2.3:a:graphviz:graphviz:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://gitlab.com/graphviz/graphviz/-/blob/main/CHANGELOG.md
DownloadPageUrl: https://graphviz.org/download/
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Graphviz (win64) for MECM.

.DESCRIPTION
    Resolves the newest stable release from the Graphviz GitLab releases API,
    downloads the matching win64 NSIS installer, stages content to a versioned
    local folder with ARP detection metadata, and creates an MECM Application
    with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: The release feed also carries Debug builds and development snapshots;
    only the Release win64 asset of a three-part stable tag is staged.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Graphviz\Graphviz\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Graphviz).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers and
    stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Graphviz version string and exits.

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
$ReleasesApiUrl = "https://gitlab.com/api/v4/projects/4207231/releases?per_page=10"

$VendorFolder = "Graphviz"
$AppFolder    = "Graphviz"

$BaseDownloadRoot = Join-Path $DownloadRoot "Graphviz"

# The CPack NSIS installer registers under its install directory name and
# installs to %ProgramFiles%\Graphviz; the NSIS stub is 32-bit, so its ARP
# entry lands in the registry's 32-bit view.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Graphviz"
$InstallDir     = "{0}\Graphviz" -f $env:ProgramFiles

# --- Functions ---


function Assert-GraphvizPayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The package registry answers 200 with a JSON error body for an asset
        that is no longer published, which would otherwise stage as a
        valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestGraphvizRelease {
    <#
    .SYNOPSIS
        Returns the newest stable Graphviz release version, filename, and URL.
    .DESCRIPTION
        Release entries are returned in publication order and include
        development tags (16.0.1~dev.20260814) plus Debug-configuration assets,
        so tags are filtered to three-part stable versions, assets to the
        Release win64 installer, and the result sorted as versions.
    #>
    param([switch]$Quiet)

    Write-Log "Graphviz releases API        : $ReleasesApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleasesApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the Graphviz releases API: $ReleasesApiUrl" }

        $releases = ConvertFrom-Json $json
        if (-not $releases) { throw "Graphviz releases API returned no releases." }

        $candidates = foreach ($r in $releases) {
            if ($r.tag_name -notmatch '^\d+\.\d+\.\d+$') { continue }

            $asset = $r.assets.links |
                Where-Object { $_.name -match '^windows_\d+_cmake_Release_graphviz-install-.+-win64\.exe$' } |
                Select-Object -First 1
            if (-not $asset) { continue }

            [pscustomobject]@{
                Version     = $r.tag_name
                FileName    = $asset.name
                DownloadUrl = $asset.direct_asset_url
            }
        }

        if (-not $candidates -or @($candidates).Count -lt 1) {
            throw "No stable release carried a Release win64 installer asset."
        }

        $best = @($candidates) | Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1

        Write-Log "Latest Graphviz version      : $($best.Version)" -Quiet:$Quiet
        Write-Log "Installer filename           : $($best.FileName)" -Quiet:$Quiet

        return $best
    }
    catch {
        Write-Log "Failed to get Graphviz version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageGraphviz {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Graphviz (win64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestGraphvizRelease
    if (-not $releaseInfo) { throw "Could not resolve the latest Graphviz release." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

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

    Assert-GraphvizPayloadIsExecutable -Path $localExe

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedExe"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/S'" `
        -UninstallCommand (Join-Path $InstallDir 'Uninstall.exe') `
        -UninstallArgs "'/S'"

    # The installer's own ARP entry names the uninstaller, which is the only
    # value that stays correct when a machine carries a non-default install
    # directory. The 32-bit view is read first because the NSIS stub writes
    # there.
    $uninstallContent = @'
$keys = @(
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Graphviz',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Graphviz'
)
foreach ($key in $keys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $cmd = (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue).UninstallString
    if (-not $cmd) { continue }
    if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
    if (-not (Test-Path -LiteralPath $exe)) { continue }
    $proc = Start-Process -FilePath $exe -ArgumentList @('/S') -Wait -PassThru -NoNewWindow
    exit $proc.ExitCode
}
exit 0
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "ARP RegistryKey              : $ArpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log "Install directory            : $InstallDir"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Graphviz"
        Publisher       = "Graphviz"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("dot", "gvedit")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $false
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

function Invoke-PackageGraphviz {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Graphviz (win64) - PACKAGE phase"
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
        $info = Get-LatestGraphvizRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Graphviz GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Graphviz (win64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleasesApiUrl               : $ReleasesApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageGraphviz
    }
    elseif ($PackageOnly) {
        Invoke-PackageGraphviz
    }
    else {
        Invoke-StageGraphviz
        Invoke-PackageGraphviz
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-graphviz'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
