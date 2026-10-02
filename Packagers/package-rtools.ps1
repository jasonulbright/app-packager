<#
Vendor: The R Foundation
App: Rtools
CMName: Rtools
VendorUrl: https://cran.r-project.org/bin/windows/Rtools/
ReleaseNotesUrl: https://cran.r-project.org/bin/windows/Rtools/
DownloadPageUrl: https://cran.r-project.org/bin/windows/Rtools/
IconSource: None
WsusSupport: Yes

.SYNOPSIS
    Packages Rtools (x64) for ConfigMgr.

.DESCRIPTION
    Reads the CRAN Rtools index for the newest Rtools line (for example
    Rtools45 for R 4.5), downloads its 64-bit Intel installer, stages content
    to a versioned local folder, and creates a ConfigMgr Application.

    The package version is the installer FileVersion, for example
    4.5.6768.6492: the Rtools line, the toolchain revision and the base
    revision. The installer registers "Rtools<NN>_is1" in Add/Remove Programs
    with DisplayVersion <line>.<toolchain revision>, and detection compares
    that value as a version. The key name stays the same for every build of
    one line, so a newer build upgrades in place.

    Rtools installs to C:\rtools<NN>. Each line installs beside the other
    lines. R itself is package-r.ps1.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\The R Foundation\Rtools\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Rtools).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry version detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Rtools version string and exits.

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
$IndexUrl = "https://cran.r-project.org/bin/windows/Rtools/"

$VendorFolder = "The R Foundation"
$AppFolder    = "Rtools"

$BaseDownloadRoot = Join-Path $DownloadRoot "Rtools"

# --- Functions ---


function Get-LatestRtoolsRelease {
    <#
    .SYNOPSIS
        Finds the newest Rtools line on the CRAN index and its x64 installer.
        Returns a PSCustomObject with Line, Version, DisplayVersion, FileName,
        and DownloadUrl.
    #>
    param([switch]$Quiet)

    Write-Log "Rtools index                 : $IndexUrl" -Quiet:$Quiet

    try {
        $index = (curl.exe -L --fail --silent --show-error $IndexUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Rtools index: $IndexUrl" }

        # Lines are listed as rtools<major><minor>/rtools.html (rtools45 = 4.5).
        $lines = @([regex]::Matches($index, 'href="rtools(\d)(\d+)/rtools\.html"') | ForEach-Object {
                [pscustomobject]@{ Folder = 'rtools{0}{1}' -f $_.Groups[1].Value, $_.Groups[2].Value; Line = '{0}.{1}' -f $_.Groups[1].Value, $_.Groups[2].Value }
            })
        if ($lines.Count -eq 0) { throw "No Rtools line found on the index." }
        $newest = $lines | Sort-Object { [version]$_.Line } -Descending | Select-Object -First 1

        $pageUrl = '{0}{1}/rtools.html' -f $IndexUrl, $newest.Folder
        $page = (curl.exe -L --fail --silent --show-error (ConvertTo-SafeCurlUrl $pageUrl)) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Rtools page: $pageUrl" }

        # The aarch64 installer carries "-aarch64-" in its name, so this pattern
        # selects the 64-bit Intel installer only.
        $pattern = 'href="files/({0}-(\d+)-(\d+)\.exe)"' -f [regex]::Escape($newest.Folder)
        $match = [regex]::Match($page, $pattern)
        if (-not $match.Success) { throw "No x64 installer link found on $pageUrl" }

        $fileName = $match.Groups[1].Value
        $version = '{0}.{1}.{2}' -f $newest.Line, $match.Groups[2].Value, $match.Groups[3].Value

        Write-Log "Latest Rtools version        : $version" -Quiet:$Quiet

        return [PSCustomObject]@{
            Line           = $newest.Line
            Folder         = $newest.Folder
            Version        = $version
            DisplayVersion = '{0}.{1}' -f $newest.Line, $match.Groups[2].Value
            FileName       = $fileName
            DownloadUrl    = '{0}{1}/files/{2}' -f $IndexUrl, $newest.Folder, $fileName
        }
    }
    catch {
        Write-Log "Failed to get Rtools version info: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageRtools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Rtools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-LatestRtoolsRelease
    if (-not $release) { throw "Could not resolve the Rtools version." }

    $version           = $release.Version
    $installerFileName = $release.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Rtools installer..."
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    $fileVersion = ([string](Get-Item -LiteralPath $localExe).VersionInfo.FileVersion).Trim()
    Write-Log "Installer FileVersion        : $fileVersion"
    if ($fileVersion -and $fileVersion -ne $version) {
        throw "The installer reports version $fileVersion, but the CRAN file name says $version."
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
    $installDir   = 'C:\{0}' -f $release.Folder
    $uninstallCmd = Join-Path $installDir 'unins000.exe'

    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-'" `
        -UninstallCommand $uninstallCmd `
        -UninstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $arpKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{0}_is1' -f ($release.Folder.Substring(0, 1).ToUpperInvariant() + $release.Folder.Substring(1))

    $appName   = "Rtools $version"
    $publisher = "The R Foundation"

    Write-Log ""
    Write-Log "Detection                    : HKLM\$arpKey\DisplayVersion >= $($release.DisplayVersion)"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallCommand = $uninstallCmd
        RunningProcess  = @()
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $release.DisplayVersion
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

function Invoke-PackageRtools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Rtools (x64) - PACKAGE phase"
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

    # --- ConfigMgr application ---
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
        $rel = Get-LatestRtoolsRelease -Quiet
        if (-not $rel) { exit 1 }
        Write-Output $rel.Version
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
    Write-Log "Rtools (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "IndexUrl                     : $IndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageRtools
    }
    elseif ($PackageOnly) {
        Invoke-PackageRtools
    }
    else {
        Invoke-StageRtools
        Invoke-PackageRtools
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-rtools'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
