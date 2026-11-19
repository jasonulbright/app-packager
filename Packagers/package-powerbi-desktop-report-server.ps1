<#
Vendor: Microsoft
App: Microsoft Power BI Desktop (optimized for Power BI Report Server) (x64)
CMName: Microsoft Power BI Desktop (optimized for Power BI Report Server)
VendorUrl: https://www.microsoft.com/power-platform/products/power-bi/report-server
ReleaseNotesUrl: https://learn.microsoft.com/power-bi/report-server/changelog
DownloadPageUrl: https://www.microsoft.com/download/details.aspx?id=105945
SecurityAdvisoryUrl: https://msrc.microsoft.com/update-guide
IconSource: External
UpdateCadenceDays: 30
WsusSupport: No (DetectionNotMappable)
ArpName: Microsoft Power BI Desktop (* 20??) (x64)
ArpArch: x64

.SYNOPSIS
    Packages Power BI Desktop (optimized for Power BI Report Server), x64, for ConfigMgr.

.DESCRIPTION
    This is the client application that report authors install to create Power BI reports
    for a Power BI Report Server. It is not the report server. Microsoft Learn ("Install
    Power BI Desktop for Power BI Report Server") says that this edition is separate from
    Power BI Desktop for the Power BI service and that both can be installed on one
    computer. This packager is separate from the Power BI Desktop packager
    (package-powerbi-desktop.ps1) in the same way: its own installer, its own install
    folder (Program Files\Microsoft Power BI Desktop RS) and its own entries in Add or
    Remove Programs.

    The version comes from the change log of Power BI Report Server on Microsoft Learn: the
    newest entry under "Power BI Desktop (optimized for Power BI Report Server)". The
    installer comes from the Download Center page of Power BI Report Server (the target of
    "Advanced download options" on the Power BI Report Server home page). That page lists
    the build of the report server, not the version of Power BI Desktop, so the version of
    the installer is read from the installer itself. The stage stops when the installer
    does not report the version that the change log names (for example when the change log
    and the Download Center page are not updated at the same time), and it stages nothing.

    The installer is a Windows Installer bundle (PBIDesktopSetupRS_x64.exe, x64 only). The
    download is checked for a valid Authenticode signature of Microsoft Corporation. The
    vendor publishes no digest for the file, so the signature is the check. The Learn page
    "Install Power BI Desktop for Power BI Report Server" names an MSI file; the Download
    Center page offers the EXE bundle, and the EXE is what this packager stages.

    Microsoft documents the command-line options of the Power BI Desktop installer on the
    Learn page "Download Power BI Desktop" (section "Use command-line options during
    installation"): -quiet, -norestart, -uninstall, ACCEPT_EULA=1 and
    DISABLE_UPDATE_NOTIFICATION=1. The install and the uninstall run as SYSTEM on a lab
    computer with exit code 0.

    Detection reads the DisplayVersion value of the Add or Remove Programs entry of the
    bundle. The key is named with the bundle id that the installer stores in its header
    (the 32-bit registry view, like every Windows Installer bundle), so the key belongs to
    this edition only and the other edition cannot satisfy it. The rule accepts the
    packaged version or a newer one.

    Notes:
      - The file association of .pbix files goes to the edition that was installed last
        (Microsoft Learn, "Install Power BI Desktop for Power BI Report Server").
      - Both editions run as PBIDesktop.exe. The ConfigMgr running process check of the
        deployment type therefore asks the user to close a running Power BI Desktop of
        either edition.
      - There is no CPE for this edition. The NVD dictionary has no entry for Power BI
        Desktop. The nearest entry, microsoft:power_bi_report_server, is versioned with
        the build of the report server (15.0.x), not with the version of this application
        (2.x), so a lookup with it would compare unrelated version numbers. The CVE column
        shows "CPE missing" for this application.

    Supports two-phase operation:
      -StageOnly    Download the EXE, write wrappers and manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr app

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (Package phase)
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
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

# The page that lists the versions of Power BI Report Server and of Power BI Desktop (optimized for Power BI Report Server).
$ChangeLogUrl = "https://learn.microsoft.com/en-us/power-bi/report-server/changelog"
# The Download Center page of Power BI Report Server ("Advanced download options" on the Power BI Report Server home page).
$DownloadCenterUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=105945"
$InstallerFileName = "PBIDesktopSetupRS_x64.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "Power BI Desktop for Report Server"
$BaseDownloadRoot = Join-Path $DownloadRoot "PowerBIDesktopRS"

function Get-PowerBIReportServerDesktopVersion {
    # The newest version under "Power BI Desktop (optimized for Power BI Report Server)" in the change log, or $null.
    param([switch]$Quiet)

    Write-Log "Change log URL               : $ChangeLogUrl" -Quiet:$Quiet
    try {
        $html = (curl.exe -L --fail --silent --show-error --max-time 60 $ChangeLogUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to read the change log of Power BI Report Server." }

        $match = [regex]::Match(
            $html,
            'Power BI Desktop \(optimized for Power BI Report Server\)</h3>\s*<ul>\s*<li>\s*<p>\s*<em>Version:\s*(?<v>\d+\.\d+\.\d+\.\d+)'
        )
        if (-not $match.Success) { throw "Could not find a Power BI Desktop version in the change log." }

        $version = $match.Groups['v'].Value
        Write-Log "Latest Power BI Desktop (Report Server) version : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get the Power BI Desktop (Report Server) version: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function Get-PowerBIReportServerDownloadInfo {
    # Version (change log) and installer URL (Download Center page), or $null.
    param([switch]$Quiet)

    $version = Get-PowerBIReportServerDesktopVersion -Quiet:$Quiet
    if (-not $version) { return $null }

    Write-Log "Download Center URL          : $DownloadCenterUrl" -Quiet:$Quiet
    try {
        $html = (curl.exe -L --fail --silent --show-error --max-time 60 $DownloadCenterUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Microsoft Download Center." }

        $jsonMatch = [regex]::Match(
            $html,
            'window\.__DLCDetails__=(?<json>\{.*?\})(?=</script>)',
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )
        if (-not $jsonMatch.Success) { throw "Could not find Download Center metadata block." }

        $details = $jsonMatch.Groups['json'].Value | ConvertFrom-Json -ErrorAction Stop
        $file = @($details.dlcDetailsView.downloadFile | Where-Object { $_.name -eq $InstallerFileName } | Select-Object -First 1)
        if (-not $file) { throw "Could not find $InstallerFileName in Download Center metadata." }

        $url = [string]$file.url
        if ([string]::IsNullOrWhiteSpace($url)) { throw "The Download Center metadata has no URL for $InstallerFileName." }

        return [pscustomobject]@{
            Version     = $version
            FileName    = [string]$file.name
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get Power BI Desktop (Report Server) metadata: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function Get-PowerBIReportServerInstallerVersion {
    # The file version that the installer itself reports (the version resource), or an empty string.
    param([Parameter(Mandatory)][string]$Path)

    return [string](Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.FileVersion
}

function New-PowerBIReportServerInstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { Write-Error "Missing Power BI Desktop (Report Server) installer"; exit 2 }',
        '$args = @(''-quiet'', ''-norestart'', ''ACCEPT_EULA=1'', ''DISABLE_UPDATE_NOTIFICATION=1'')',
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-PowerBIReportServerUninstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { exit 0 }',
        '$args = @(''-uninstall'', ''-quiet'', ''-norestart'')',
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function Invoke-StagePowerBIReportServerDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Power BI Desktop (optimized for Power BI Report Server) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $info = Get-PowerBIReportServerDownloadInfo
    if (-not $info) { throw "Could not resolve the Power BI Desktop (Report Server) download info." }

    $version = $info.Version
    $installerName = $info.FileName
    $downloadUrl = $info.DownloadUrl
    $localInstaller = Join-Path $BaseDownloadRoot $installerName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Local installer              : $localInstaller"
    Write-Log ""

    # The file name does not carry the version. A cached file counts only when it reports the version
    # of the change log; otherwise it is removed and downloaded again once.
    if ((Test-Path -LiteralPath $localInstaller) -and ((Get-PowerBIReportServerInstallerVersion -Path $localInstaller) -ne $version)) {
        Write-Log "The local installer is not version $version. Downloading it again." -Level WARN
        Remove-Item -LiteralPath $localInstaller -Force -ErrorAction Stop
    }
    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    $signer = Assert-InstallerSignature -Path $localInstaller -ExpectedSigner 'Microsoft Corporation' -RemoveOnFailure
    Write-Log "Authenticode signer          : $signer"

    $fileVersion = Get-PowerBIReportServerInstallerVersion -Path $localInstaller
    if ($fileVersion -ne $version) {
        Remove-Item -LiteralPath $localInstaller -Force -ErrorAction SilentlyContinue
        throw ("The installer {0} reports version {1}, but the change log of Power BI Report Server names version {2}. The Download Center page and the change log may not be updated at the same time. The file was removed and nothing was staged." -f $installerName, $fileVersion, $version)
    }
    Write-Log "Installer file version       : $fileVersion"

    $bundleId = Get-BurnBundleId -Path $localInstaller
    if (-not $bundleId) { throw "Could not read the bundle id from $localInstaller." }
    $arpKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$bundleId"

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localInstaller -Destination (Join-Path $localContentPath $installerName) -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-PowerBIReportServerInstallWrapper -InstallerFile $installerName) `
        -UninstallPs1Content (New-PowerBIReportServerUninstallWrapper -InstallerFile $installerName)

    Write-Log ""
    Write-Log "Detection key                : $arpKey DisplayVersion >= $version (32-bit view)"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Microsoft Power BI Desktop (optimized for Power BI Report Server) - $version"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        DisplayName     = "Microsoft Power BI Desktop (optimized for Power BI Report Server)"
        InstallerFile   = $installerName
        InstallerType   = "EXE"
        InstallArgs     = "-quiet -norestart ACCEPT_EULA=1 DISABLE_UPDATE_NOTIFICATION=1"
        UninstallArgs   = "-uninstall -quiet -norestart"
        RunningProcess  = @("PBIDesktop")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
            Is64Bit             = $false
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}

function Invoke-PackagePowerBIReportServerDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Power BI Desktop (optimized for Power BI Report Server) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }

    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()
    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifest = Read-StageManifest -Path (Join-Path $localContentPath "stage-manifest.json")

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
}

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $latest = Get-PowerBIReportServerDesktopVersion -Quiet
        if (-not $latest) { exit 1 }
        Write-Output $latest
        exit 0
    }
    catch {
        exit 1
    }
}

try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Power BI Desktop (optimized for Power BI Report Server) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePowerBIReportServerDesktop
    }
    elseif ($PackageOnly) {
        Invoke-PackagePowerBIReportServerDesktop
    }
    else {
        Invoke-StagePowerBIReportServerDesktop
        Invoke-PackagePowerBIReportServerDesktop
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-powerbi-desktop-report-server'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
