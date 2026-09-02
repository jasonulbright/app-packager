<#
Vendor: TGRMN Software
App: Bulk Rename Utility
CMName: Bulk Rename Utility
VendorUrl: https://www.bulkrenameutility.co.uk/
CPE: cpe:2.3:a:tgrmn:bulk_rename_utility:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.bulkrenameutility.co.uk/Downloads/BRUChangelog.pdf
DownloadPageUrl: https://www.bulkrenameutility.co.uk/Download.php
UpdateCadenceDays: 180

.SYNOPSIS
    Packages Bulk Rename Utility (x64) for MECM.

.DESCRIPTION
    Resolves the newest build from the vendor's unversioned download endpoint,
    which redirects to a versioned installer filename, downloads it, stages
    content to a versioned local folder, and creates an MECM Application with
    script-based detection on the ARP entry.

    The payload is an Inno Setup installer, so the silent switches are
    /VERYSILENT /SUPPRESSMSGBOXES /NORESTART and the uninstaller is resolved
    from the ARP entry rather than assumed at a fixed path.

    The download is free for personal use only; business, government and
    educational use requires a purchased license. A license entitlement must be
    in place before this payload is deployed in an organizational environment.

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
    Content is staged under: <FileServerPath>\Applications\TGRMN Software\Bulk Rename Utility\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\BulkRenameUtility).
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
    create MECM application with script-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Bulk Rename Utility version string and exits.

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
$DownloadPageUrl  = "https://www.bulkrenameutility.co.uk/Download.php"
$SetupRedirectUrl = "https://www.bulkrenameutility.co.uk/Down/BRU_setup.exe"

$VendorFolder = "TGRMN Software"
$AppFolder    = "Bulk Rename Utility"

$BaseDownloadRoot = Join-Path $DownloadRoot "BulkRenameUtility"

$InstallerFileName = "BRU_setup.exe"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The download host answers a retired build path with a 200 or 410 HTML
        body, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestBulkRenameUtilityRelease {
    <#
    .SYNOPSIS
        Returns the newest Bulk Rename Utility version and its installer URL.
    .DESCRIPTION
        The download page drives its button from script rather than a static
        link, so the version is taken from the redirect target of the vendor's
        unversioned endpoint (BRU_setup_<version>.exe), which is the build the
        vendor is actually serving.
    #>
    param([switch]$Quiet)

    Write-Log "BRU download endpoint        : $SetupRedirectUrl" -Quiet:$Quiet

    try {
        $resolved = (curl.exe -sIL --fail --show-error -A "Mozilla/5.0" -o NUL -w "%{url_effective}" $SetupRedirectUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve the BRU download redirect: $SetupRedirectUrl" }

        $m = [regex]::Match($resolved, 'BRU_setup_(?<ver>\d+(\.\d+)+)\.exe')
        if (-not $m.Success) {
            throw "Redirect target does not carry a versioned installer filename: $resolved"
        }

        $version = $m.Groups['ver'].Value

        Write-Log "Latest BRU version           : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $InstallerFileName
            DownloadUrl = $resolved
        }
    }
    catch {
        Write-Log "Failed to get Bulk Rename Utility version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBulkRenameUtility {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bulk Rename Utility (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestBulkRenameUtilityRelease
    if (-not $releaseInfo) { throw "Could not resolve Bulk Rename Utility version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Download ---
    # The staged filename is fixed, so a cached copy from an earlier release
    # would shadow the new one; the payload is re-fetched every stage.
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe

    Assert-PayloadIsExecutable -Path $localExe

    $productVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).ProductVersion
    if ($productVersion) { $productVersion = $productVersion.Trim() }
    Write-Log "Installer ProductVersion     : $productVersion"
    if ($productVersion -and $productVersion -ne $version) {
        throw "Downloaded installer reports version '$productVersion' but the download endpoint served '$version'; refusing to stage a mismatched payload."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    # --- Generate content wrappers ---
    # The Inno uninstaller is written next to an operator-selectable target, so
    # uninstall resolves it from the ARP entry instead of a fixed path.
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallScript = @'
$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    foreach ($sub in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
        $entry = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction SilentlyContinue
        if (-not $entry -or [string]$entry.DisplayName -notlike 'Bulk Rename Utility*') { continue }
        if (-not $entry.UninstallString) { continue }
        # UninstallString is the uninstaller path, quoted or bare; only the
        # executable is wanted so the silent switches can be re-supplied here.
        $exe = ($entry.UninstallString -replace '^"([^"]+)".*$', '$1').Trim('"')
        if (-not (Test-Path -LiteralPath $exe)) { continue }
        $proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru -NoNewWindow
        exit $proc.ExitCode
    }
}
Write-Error 'Bulk Rename Utility uninstall entry not found.'
exit 1
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # The Inno ARP key name carries the product's app id, which the vendor has
    # changed across major versions, so detection matches on DisplayName and
    # compares DisplayVersion.
    $detectionScript = @"
`$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
`$wanted = [version]'$version'
foreach (`$root in `$roots) {
    if (-not (Test-Path -LiteralPath `$root)) { continue }
    foreach (`$sub in Get-ChildItem -LiteralPath `$root -ErrorAction SilentlyContinue) {
        `$entry = Get-ItemProperty -LiteralPath `$sub.PSPath -ErrorAction SilentlyContinue
        if (-not `$entry -or [string]`$entry.DisplayName -notlike 'Bulk Rename Utility*') { continue }
        `$found = `$null
        if (-not [version]::TryParse([string]`$entry.DisplayVersion, [ref]`$found)) { continue }
        if (`$found -ge `$wanted) { Write-Output 'Installed'; exit 0 }
    }
}
exit 0
"@

    Write-Log ""
    Write-Log "Detection                    : ARP DisplayName 'Bulk Rename Utility*' with DisplayVersion >= $version"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Bulk Rename Utility"
        Publisher       = "TGRMN Software"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("BulkRenameUtility")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
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

function Invoke-PackageBulkRenameUtility {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bulk Rename Utility (x64) - PACKAGE phase"
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
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
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
        $info = Get-LatestBulkRenameUtilityRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Bulk Rename Utility GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Bulk Rename Utility (x64) Auto-Packager starting"
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
        Invoke-StageBulkRenameUtility
    }
    elseif ($PackageOnly) {
        Invoke-PackageBulkRenameUtility
    }
    else {
        Invoke-StageBulkRenameUtility
        Invoke-PackageBulkRenameUtility
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-bulkrenameutility'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
