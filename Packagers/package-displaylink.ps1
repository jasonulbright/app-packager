<#
Vendor: DisplayLink
App: DisplayLink Graphics
CMName: DisplayLink Graphics
VendorUrl: https://www.synaptics.com/products/displaylink-graphics
CPE: cpe:2.3:a:displaylink:displaylink_graphics:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://support.displaylink.com/knowledgebase/topics/106669-windows-release-notes
DownloadPageUrl: https://www.synaptics.com/products/displaylink-graphics/downloads/windows
UpdateCadenceDays: 120

.SYNOPSIS
    Packages the DisplayLink USB Graphics Software (dock driver) for MECM.

.DESCRIPTION
    Resolves the newest public Windows driver release from the vendor download
    index, downloads the self-extracting installer, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection.

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
    Content is staged under: <FileServerPath>\Applications\DisplayLink\DisplayLink Graphics\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available DisplayLink driver version string and exits.

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
$SiteRoot        = "https://www.synaptics.com"
$DownloadIndexUrl = "$SiteRoot/products/displaylink-graphics/downloads/windows"

$VendorFolder = "DisplayLink"
$AppFolder    = "DisplayLink Graphics"

$BaseDownloadRoot = Join-Path $DownloadRoot "DisplayLink Graphics"
$InstallerFileName = "DisplayLink-Installer.exe"
$CoreSoftwareDir   = Join-Path $env:ProgramFiles "DisplayLink Core Software"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The vendor site answers 200 with an HTML page when a file path no
        longer resolves, which would otherwise stage as a valid-looking
        installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 2) } finally { $stream.Dispose() }

    if ($read -lt 2 -or $buffer[0] -ne 0x4D -or $buffer[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Resolve-DisplayLinkInstallerUrl {
    <#
    .SYNOPSIS
        Returns the download URL of the newest public Windows driver release.
    .DESCRIPTION
        The download index links one node page per release, named
        windows-<major>.<minor>-m<milestone>-public. Release order is not
        document order, so the candidates are sorted as versions with the
        milestone as a third field. The node page carries the direct file path
        under /sites/default/files/exe_files; the corporate MSI bundles behind
        the same index sit behind a licence-acceptance form and are not
        machine-fetchable, so the public self-extracting installer is used.
    #>
    param([switch]$Quiet)

    Write-Log "DisplayLink download index   : $DownloadIndexUrl" -Quiet:$Quiet

    $indexHtml = (curl.exe -L --fail --silent --show-error $DownloadIndexUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the DisplayLink download index." }

    $rx = [regex]'downloads/windows-(?<major>\d+)\.(?<minor>\d+)-m(?<milestone>\d+)-public'
    $rxMatches = $rx.Matches($indexHtml)
    if ($rxMatches.Count -lt 1) {
        throw "Could not locate any public Windows driver releases in the download index."
    }

    $candidates = foreach ($m in $rxMatches) {
        [pscustomobject]@{
            Slug  = $m.Value
            Order = [version]("{0}.{1}.{2}" -f $m.Groups['major'].Value, $m.Groups['minor'].Value, $m.Groups['milestone'].Value)
            Label = "{0}.{1} M{2}" -f $m.Groups['major'].Value, $m.Groups['minor'].Value, $m.Groups['milestone'].Value
        }
    }

    $best = $candidates | Sort-Object Order -Descending | Select-Object -First 1
    $nodeUrl = "$SiteRoot/products/displaylink-graphics/$($best.Slug)"

    Write-Log "Newest public release        : $($best.Label)" -Quiet:$Quiet
    Write-Log "Release page                 : $nodeUrl" -Quiet:$Quiet

    $nodeHtml = (curl.exe -L --fail --silent --show-error "$nodeUrl`?filetype=exe") -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the DisplayLink release page: $nodeUrl" }

    $fileRx = [regex]'(?i)(?<href>/sites/default/files/exe_files/[^"'']+\.exe)'
    $fileMatch = $fileRx.Match($nodeHtml)
    if (-not $fileMatch.Success) {
        throw "Release page carried no direct installer path: $nodeUrl"
    }

    $url = $SiteRoot + $fileMatch.Groups['href'].Value
    Write-Log "Resolved installer URL       : $url" -Quiet:$Quiet

    return [pscustomobject]@{
        Label       = $best.Label
        DownloadUrl = $url
    }
}


function Get-DisplayLinkStagedInstaller {
    <#
    .SYNOPSIS
        Downloads the newest installer if needed and returns its path and version.
    .DESCRIPTION
        The release label on the site ("11.5 M0") is not the version the
        product registers, so the four-field build number is read from the
        downloaded installer's file version. The download is cached under the
        release label because the vendor filename is not stable across
        releases.
    #>
    param([switch]$Quiet)

    $release = Resolve-DisplayLinkInstallerUrl -Quiet:$Quiet

    $cacheDir = Join-Path $BaseDownloadRoot ($release.Label -replace '[\\/:*?"<>|]', '_')
    Initialize-Folder -Path $cacheDir

    $localInstaller = Join-Path $cacheDir $InstallerFileName
    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Write-Log "Downloading installer (~65 MB)..." -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localInstaller -Quiet:$Quiet
    }
    else {
        Write-Log "Cached installer exists. Skipping download." -Quiet:$Quiet
    }

    Assert-ExePayload -Path $localInstaller

    # The version resource uses comma separators ("11, 5, 5742, 0").
    $raw = (Get-Item -LiteralPath $localInstaller).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "Installer carries no file version resource: $localInstaller" }

    $version = ($raw -replace '[^\d.,]', '') -replace ',', '.' -replace '\s', ''
    if ($version -notmatch '^\d+(\.\d+){1,3}$') {
        throw "Installer file version is not a numeric version: $raw"
    }

    return [pscustomobject]@{
        Label     = $release.Label
        Version   = $version
        LocalPath = $localInstaller
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDisplayLink {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DisplayLink Graphics - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $installer = Get-DisplayLinkStagedInstaller
    $version   = $installer.Version

    Write-Log "Release label                : $($installer.Label)"
    Write-Log "Driver version               : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedInstaller = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedInstaller)) {
        Copy-Item -LiteralPath $installer.LocalPath -Destination $stagedInstaller -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedInstaller"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # Vendor switches: -silent suppresses the progress dialogs, -noreboot keeps
    # the installer from restarting the machine so the deployment reports the
    # reboot instead. Uninstall runs the same installer with -uninstall, so the
    # uninstall command stays inside content rather than depending on an ARP
    # entry whose product code changes with every release.
    $installPs1 = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('-silent','-noreboot') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallPs1 = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @('-uninstall','-silent','-noreboot') -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    # Detection reads the core service binary rather than an ARP key: the
    # installer registers under a per-release MSI product code that cannot be
    # derived from the downloaded payload.
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "DisplayLink Graphics"
        Publisher        = "DisplayLink Corp."
        SoftwareVersion  = $version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = "-silent -noreboot"
        UninstallCommand = $InstallerFileName
        UninstallArgs    = "-uninstall -silent -noreboot"
        RunningProcess   = @("DisplayLinkManager","DisplayLinkUI")
        Detection        = @{
            Type          = "File"
            FilePath      = $CoreSoftwareDir
            FileName      = "DisplayLinkManager.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageDisplayLink {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DisplayLink Graphics - PACKAGE phase"
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
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
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
        Initialize-Folder -Path $BaseDownloadRoot
        $installer = Get-DisplayLinkStagedInstaller -Quiet
        Write-Output $installer.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("DisplayLink GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "DisplayLink Graphics Auto-Packager starting"
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
        Invoke-StageDisplayLink
    }
    elseif ($PackageOnly) {
        Invoke-PackageDisplayLink
    }
    else {
        Invoke-StageDisplayLink
        Invoke-PackageDisplayLink
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-displaylink'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
