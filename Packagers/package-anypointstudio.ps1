<#
Vendor: MuleSoft
App: Anypoint Studio
CMName: Anypoint Studio
VendorUrl: https://www.mulesoft.com/platform/studio
ReleaseNotesUrl: https://docs.mulesoft.com/release-notes/studio/anypoint-studio
DownloadPageUrl: https://www.mulesoft.com/lp/dl/anypoint-mule-studio
IconSource: None
UpdateCadenceDays: 60

.SYNOPSIS
    Packages MuleSoft Anypoint Studio (x64) for ConfigMgr.

.DESCRIPTION
    Reads the latest Windows build from the MuleSoft downloads manifest,
    downloads the ZIP, verifies it against the manifest's SHA-256, stages
    content to a versioned local folder, and creates a ConfigMgr Application
    with file-existence detection.

    Anypoint Studio ships as a ZIP with no installer. The install wrapper
    removes any earlier C:\AnypointStudio and extracts the ZIP to C:\, the
    location the vendor's install guide names; folders created under C:\
    inherit write access for users, which Studio expects for its own
    folder. The uninstall wrapper removes the folder. Workspaces live under
    the user profile and are not touched.

    Detection is the feature.xml of the versioned Studio feature folder
    (features\org.mule.tooling.studio_<build>), read from the ZIP at Stage.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest Anypoint Studio version from the downloads manifest and exits.

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
    [int]$EstimatedRuntimeMins = 20,
    [int]$MaximumRuntimeMins = 60,
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
$ManifestUrl  = "https://www.mulesoft.com/downloads/manifest.json"
$DownloadPage = "https://www.mulesoft.com/lp/dl/anypoint-mule-studio"
$InstallDir   = "C:\AnypointStudio"

# The download CDN answers 403 unless a request carries a browser
# User-Agent and the download page as Referer; either one alone is refused.
$CdnCurlArgs = @(
    '-A', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
    '-e', $DownloadPage
)

$VendorFolder = "MuleSoft"
$AppFolder    = "Anypoint Studio"

$BaseDownloadRoot = Join-Path $DownloadRoot "AnypointStudio"

# --- Functions ---


function Get-AnypointStudioReleaseFromManifest {
    <#
    .SYNOPSIS
        Picks the latest Windows Studio entry from the downloads manifest.
    .DESCRIPTION
        The manifest is a JSON array of { name, version, os, source,
        integrity, packaging }; version is 'latest' or 'previous', and the
        release number appears only in the source file name.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Json)

    # Windows PowerShell 5.1 ConvertFrom-Json emits a JSON array as one
    # object; piping the variable enumerates the entries.
    $parsed = ConvertFrom-Json -InputObject $Json
    $entry = $parsed | Where-Object { $_.name -eq 'studio' -and $_.version -eq 'latest' -and $_.os -eq 'windows' } | Select-Object -First 1
    if (-not $entry) { throw "The downloads manifest has no latest Windows Studio entry." }

    $fileName = Split-Path -Path ([uri][string]$entry.source).AbsolutePath -Leaf
    $m = [regex]::Match($fileName, '^AnypointStudio-(\d+\.\d+\.\d+)-win64\.zip$')
    if (-not $m.Success) { throw "Unexpected Windows Studio file name in the downloads manifest: $fileName" }

    $sha256 = ([string]$entry.integrity).Trim().ToUpperInvariant()
    if ($sha256 -notmatch '^[0-9A-F]{64}$') { throw "The downloads manifest carries no SHA-256 for $fileName." }

    return [pscustomobject]@{
        Version     = $m.Groups[1].Value
        FileName    = $fileName
        DownloadUrl = [string]$entry.source
        Sha256      = $sha256
    }
}


function Get-LatestAnypointStudioRelease {
    param([switch]$Quiet)

    Write-Log "Downloads manifest           : $ManifestUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error @CdnCurlArgs $ManifestUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to read the MuleSoft downloads manifest." }

        $release = Get-AnypointStudioReleaseFromManifest -Json $json
        Write-Log "Latest Anypoint Studio       : $($release.Version)" -Quiet:$Quiet
        return $release
    }
    catch {
        Write-Log "Failed to get Anypoint Studio version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-AnypointStudioFeatureVersion {
    <#
    .SYNOPSIS
        Reads the full Studio build version (for example 7.25.0.202605141228)
        from the feature folder name inside the ZIP.
    #>
    param([Parameter(Mandatory)][string]$ZipPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $hasExe = $false
        $featureVersion = $null
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -eq 'AnypointStudio/AnypointStudio.exe') { $hasExe = $true }
            $m = [regex]::Match($entry.FullName, '^AnypointStudio/features/org\.mule\.tooling\.studio_([^/]+)/feature\.xml$')
            if ($m.Success) { $featureVersion = $m.Groups[1].Value }
        }
        if (-not $hasExe) { throw "The ZIP has no AnypointStudio\AnypointStudio.exe: $ZipPath" }
        if (-not $featureVersion) { throw "The ZIP has no org.mule.tooling.studio feature: $ZipPath" }
        return $featureVersion
    }
    finally { $zip.Dispose() }
}


function Save-AnypointStudioIcon {
    # The ZIP is not an executable, so the icon comes from the launcher inside it.
    param([Parameter(Mandatory)][string]$ZipPath, [Parameter(Mandatory)][string]$StageRoot)

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('AnypointStudio-' + [guid]::NewGuid().ToString('N') + '.exe')
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        try {
            $entry = $zip.GetEntry('AnypointStudio/AnypointStudio.exe')
            if (-not $entry) { return $null }
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $temp, $true)
        }
        finally { $zip.Dispose() }
        $result = Get-InstallerIcon -Path $temp -OutputPath (Join-Path $StageRoot 'app-icon.ico')
        if ($result) { return (Split-Path -Leaf $result.Path) }
        return $null
    }
    catch {
        Write-Log "Icon extraction failed       : $($_.Exception.Message)" -Level WARN
        return $null
    }
    finally {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAnypointStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anypoint Studio (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestAnypointStudioRelease
    if (-not $releaseInfo) { throw "Could not resolve Anypoint Studio version." }

    $version     = $releaseInfo.Version
    $zipFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localZip = Join-Path $BaseDownloadRoot $zipFileName
    Write-Log "Local ZIP path               : $localZip"

    if (-not (Test-Path -LiteralPath $localZip)) {
        Write-Log "Downloading Anypoint Studio..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localZip -ExtraCurlArgs $CdnCurlArgs
    }
    else {
        Write-Log "Local ZIP exists. Skipping download."
    }

    $actualSha256 = (Get-FileHash -LiteralPath $localZip -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
    if ($actualSha256 -ne $releaseInfo.Sha256) {
        Remove-Item -LiteralPath $localZip -Force -ErrorAction SilentlyContinue
        throw ("SHA-256 mismatch for {0}: expected {1}, got {2}. The download was removed." -f $releaseInfo.FileName, $releaseInfo.Sha256, $actualSha256)
    }
    Write-Log "SHA-256 verified             : $actualSha256"

    $featureVersion = Get-AnypointStudioFeatureVersion -ZipPath $localZip
    Write-Log "Studio build                 : $featureVersion"

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedZip = Join-Path $localContentPath $zipFileName
    if (-not (Test-Path -LiteralPath $stagedZip)) {
        Copy-Item -LiteralPath $localZip -Destination $stagedZip -Force -ErrorAction Stop
        Write-Log "Copied ZIP to staged folder  : $stagedZip"
    }
    else {
        Write-Log "Staged ZIP exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # An earlier version is removed first: extracting over it would leave
    # its plugins beside the new ones.
    $installPs1 = (
        "`$ErrorActionPreference = 'Stop'",
        ("`$zipPath = Join-Path `$PSScriptRoot '{0}'" -f $zipFileName),
        ("`$installDir = '{0}'" -f $InstallDir),
        'try {',
        '    if (Test-Path -LiteralPath $installDir) { Remove-Item -LiteralPath $installDir -Recurse -Force }',
        '    Add-Type -AssemblyName System.IO.Compression.FileSystem',
        '    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, (Split-Path -Path $installDir -Parent))',
        '    exit 0',
        '}',
        'catch {',
        '    Write-Error $_.Exception.Message',
        '    exit 1',
        '}'
    ) -join "`r`n"

    $uninstallPs1 = (
        "`$ErrorActionPreference = 'Stop'",
        ("`$installDir = '{0}'" -f $InstallDir),
        'try {',
        '    if (Test-Path -LiteralPath $installDir) { Remove-Item -LiteralPath $installDir -Recurse -Force }',
        '    exit 0',
        '}',
        'catch {',
        '    Write-Error $_.Exception.Message',
        '    exit 1',
        '}'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $detectionPath = Join-Path $InstallDir ("features\org.mule.tooling.studio_{0}" -f $featureVersion)

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : feature.xml"
    Write-Log ""

    $manifestData = @{
        AppName         = "Anypoint Studio"
        Publisher       = "MuleSoft"
        SoftwareVersion = $version
        InstallerFile   = $zipFileName
        InstallerType   = "EXE"
        InstallArgs     = ""
        UninstallArgs   = ""
        RunningProcess  = @("AnypointStudio")
        Detection       = @{
            Type         = "File"
            FilePath     = $detectionPath
            FileName     = "feature.xml"
            PropertyType = "Existence"
            Is64Bit      = $true
        }
    }
    $icon = Save-AnypointStudioIcon -ZipPath $localZip -StageRoot $localContentPath
    if ($icon) {
        $manifestData['Icon'] = $icon
        Write-Log "Staged icon                  : $icon"
    }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageAnypointStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anypoint Studio (x64) - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
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
        $info = Get-LatestAnypointStudioRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Anypoint Studio GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Anypoint Studio (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ManifestUrl                  : $ManifestUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAnypointStudio
    }
    elseif ($PackageOnly) {
        Invoke-PackageAnypointStudio
    }
    else {
        Invoke-StageAnypointStudio
        Invoke-PackageAnypointStudio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-anypointstudio'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
