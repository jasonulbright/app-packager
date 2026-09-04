<#
Vendor: FreeCAD Team
App: FreeCAD
CMName: FreeCAD
VendorUrl: https://www.freecad.org/
CPE: cpe:2.3:a:freecad:freecad:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/FreeCAD/FreeCAD/releases
DownloadPageUrl: https://www.freecad.org/downloads.php
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages FreeCAD (x64) for MECM.

.DESCRIPTION
    Resolves the latest stable FreeCAD release from the GitHub releases API,
    downloads the Windows x86_64 NSIS installer, stages content to a versioned
    local folder with ARP detection metadata, and creates an MECM Application
    with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive detection metadata, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\FreeCAD Team\FreeCAD\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available FreeCAD version string and exits.

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
    [int]$EstimatedRuntimeMins = 20,
    [int]$MaximumRuntimeMins = 45,
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
$GitHubApiUrl = "https://api.github.com/repos/FreeCAD/FreeCAD/releases/latest"

$VendorFolder = "FreeCAD Team"
$AppFolder    = "FreeCAD"

$BaseDownloadRoot = Join-Path $DownloadRoot "FreeCAD"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A release-asset redirect that lands on an HTML error page would
        otherwise stage as a valid-looking installer and fail at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-FreeCadUninstallKeyName {
    <#
    .SYNOPSIS
        Returns the ARP key name the FreeCAD installer registers for a version.
    .DESCRIPTION
        The installer keys on FreeCAD<major><minor><patch> with no separators,
        so releases install side by side and the detection key changes with
        every patch release.
    #>
    param([Parameter(Mandatory)][string]$Version)

    $parts = @($Version -split '\.')
    if ($parts.Count -lt 3) { throw "FreeCAD version does not carry three fields: $Version" }
    return ("FreeCAD{0}{1}{2}" -f $parts[0], $parts[1], $parts[2])
}


function Get-FreeCadSeriesDirName {
    <#
    .SYNOPSIS
        Returns the install directory leaf the FreeCAD installer uses.
    .DESCRIPTION
        The installer targets "FreeCAD <major>.<minor>" under Program Files,
        which is where the uninstaller is written.
    #>
    param([Parameter(Mandatory)][string]$Version)

    $parts = @($Version -split '\.')
    if ($parts.Count -lt 2) { throw "FreeCAD version does not carry two fields: $Version" }
    return ("FreeCAD {0}.{1}" -f $parts[0], $parts[1])
}


function Get-LatestFreeCadRelease {
    <#
    .SYNOPSIS
        Returns the newest stable FreeCAD release and its Windows installer asset.
    .DESCRIPTION
        The project publishes weekly development builds as prereleases in the
        same repository, so a prerelease flag or a non-numeric tag is treated
        as a resolution failure rather than staged as stable. The 7z portable
        bundle ships alongside the installer; the asset match requires the
        installer so the portable payload is never staged.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        if ($release.prerelease) { throw "Latest FreeCAD release is flagged prerelease; refusing to stage." }

        $version = $release.tag_name -replace '^v', ''
        if ($version -notmatch '^\d+\.\d+\.\d+$') {
            throw "GitHub release tag is not a stable three-field version: $($release.tag_name)"
        }

        $asset = $release.assets |
            Where-Object { $_.name -match '(?i)^FreeCAD_.*-Windows-x86_64-.*installer\.exe$' } |
            Select-Object -First 1
        if (-not $asset) { throw "No Windows x86_64 installer asset found in the latest release." }

        Write-Log "Latest FreeCAD version       : $version" -Quiet:$Quiet
        Write-Log "Installer asset              : $($asset.name)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get FreeCAD version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageFreeCad {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "FreeCAD (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestFreeCadRelease
    if (-not $releaseInfo) { throw "Could not resolve FreeCAD version." }

    $version       = $releaseInfo.Version
    $installerName = $releaseInfo.FileName

    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ---
    $localInstaller = Join-Path $BaseDownloadRoot $installerName
    Write-Log "Local installer path         : $localInstaller"

    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Write-Log "Downloading installer (~500 MB, this can take several minutes)..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localInstaller
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localInstaller

    $fileVersion = (Get-Item -LiteralPath $localInstaller).VersionInfo.FileVersion
    Write-Log "Installer FileVersion        : $fileVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedInstaller = Join-Path $localContentPath $installerName
    if (-not (Test-Path -LiteralPath $stagedInstaller)) {
        Copy-Item -LiteralPath $localInstaller -Destination $stagedInstaller -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedInstaller"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Detection metadata ---
    $uninstallKeyName = Get-FreeCadUninstallKeyName -Version $version
    $seriesDirName    = Get-FreeCadSeriesDirName -Version $version
    $installDir       = Join-Path $env:ProgramFiles $seriesDirName
    $uninstallerPath  = Join-Path $installDir "Uninstall-FreeCAD.exe"
    $arpRegistryKey   = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + $uninstallKeyName

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $version"
    Write-Log "Install directory            : $installDir"
    Write-Log ""

    # --- Generate content wrappers ---
    # NSIS MultiUser installer: /AllUsers picks the per-machine mode that would
    # otherwise be chosen on a wizard page, /S runs it silently. The uninstaller
    # is written into the install directory, so uninstall runs from there rather
    # than from content.
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerName `
        -InstallArgs "'/AllUsers','/S'" `
        -UninstallCommand $uninstallerPath `
        -UninstallArgs "'/AllUsers','/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "FreeCAD"
        Publisher        = "FreeCAD Team"
        SoftwareVersion  = $version
        InstallerFile    = $installerName
        InstallerType    = "EXE"
        InstallArgs      = "/AllUsers /S"
        UninstallCommand = $uninstallerPath
        UninstallArgs    = "/AllUsers /S"
        RunningProcess   = @("FreeCAD","FreeCADCmd")
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $true
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

function Invoke-PackageFreeCad {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "FreeCAD (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
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
        $info = Get-LatestFreeCadRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("FreeCAD GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "FreeCAD (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "GitHubApiUrl                 : $GitHubApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageFreeCad
    }
    elseif ($PackageOnly) {
        Invoke-PackageFreeCad
    }
    else {
        Invoke-StageFreeCad
        Invoke-PackageFreeCad
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-freecad'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
