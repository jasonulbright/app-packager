<#
Vendor: Microsoft
App: Windows Admin Center
CMName: Windows Admin Center
VendorUrl: https://learn.microsoft.com/windows-server/manage/windows-admin-center/overview
CPE: cpe:2.3:a:microsoft:windows_admin_center:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/windows-server/manage/windows-admin-center/support/release-history
DownloadPageUrl: https://www.microsoft.com/evalcenter/download-windows-admin-center
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Windows Admin Center (v2, modernized gateway) for MECM.

.DESCRIPTION
    Downloads the current Windows Admin Center v2 installer from the vendor's
    permanent download link, stages content to a versioned local folder, and
    creates an MECM Application with file-version-based detection.

    The version is read from the downloaded setup binary rather than from the
    file name: the link resolves to a release-stamped name (for example
    WindowsAdminCenter2606.exe) whose digits are a release label, not the
    product version the installed files carry.

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\Windows Admin Center\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER HttpsPortNumber
    HTTPS port the gateway listens on. Ports below 1024 are not supported by
    the product. Default: 443

.PARAMETER CertificateThumbprint
    Thumbprint of a Server Authentication certificate already present in the
    LocalMachine\My store. When omitted the installer generates a self-signed
    certificate that expires after 60 days.

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Windows Admin Center version string and exits.

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
    [ValidateRange(1024, 65535)]
    [int]$HttpsPortNumber = 443,
    [string]$CertificateThumbprint = "",
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
$DownloadUrl      = "https://aka.ms/WACdownload"
$InstallerFileName = "WindowsAdminCenter.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "Windows Admin Center"

$BaseDownloadRoot = Join-Path $DownloadRoot "WindowsAdminCenter"

# Setup installs 64-bit to a fixed path and names its uninstaller by install
# order, so the first uninstaller in that folder is the product's own.
$InstallDir = "{0}\WindowsAdminCenter" -f $env:ProgramFiles

# The gateway's component binaries carry the product version, which is what
# makes an upgrade detectable; the folder itself survives an uninstall, so
# existence alone would report a removed product as installed.
$DetectionFileName = "WindowsAdminCenterAccountManagement.exe"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The download redirect can land on an interstitial that answers 200
        with HTML, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-SetupFileVersion {
    param([Parameter(Mandatory)][string]$Path)

    $version = (Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Setup binary carries no file version; cannot derive the product version."
    }
    return $version.Trim()
}


function Get-StagedInstaller {
    <#
    .SYNOPSIS
        Downloads the setup binary once and returns its path and version.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot

    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $DownloadUrl" -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe -Quiet:$Quiet
    }
    else {
        Write-Log "Local installer exists. Skipping download." -Quiet:$Quiet
    }

    Assert-PayloadIsExecutable -Path $localExe

    $version = Get-SetupFileVersion -Path $localExe
    Write-Log "Latest Windows Admin Center  : $version" -Quiet:$Quiet

    return [pscustomobject]@{ Path = $localExe; Version = $version }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWindowsAdminCenter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows Admin Center (v2) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $installer = Get-StagedInstaller
    $version   = $installer.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $InstallerFileName"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $installer.Path -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # /VERYSILENT suppresses the wizard entirely; port and certificate are
    # the only unattended settings the product exposes on the command line.
    $installArgList = @("'/VERYSILENT'", ("'/HTTPSPortNumber={0}'" -f $HttpsPortNumber))
    if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $installArgList += ("'/CertificateThumbprint={0}'" -f ($CertificateThumbprint -replace "'", "''"))
    }
    $installArgs = $installArgList -join ', '

    Write-Log "Install arguments            : $($installArgList -join ' ')"

    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand 'unused'

    # Setup names its uninstaller unins###.exe by install order, so the ARP
    # UninstallString is the only value that names the right one; the fixed
    # install path is the fallback for an entry that was already removed.
    $uninstallContent = @'
$installDir = Join-Path $env:ProgramFiles 'WindowsAdminCenter'
$exe = $null
$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $entry = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.DisplayName -like 'Windows Admin Center*' -and $_.UninstallString } |
        Select-Object -First 1
    if ($entry) {
        $cmd = $entry.UninstallString
        if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
        break
    }
}
if (-not $exe) {
    $exe = Get-ChildItem -LiteralPath $installDir -Filter 'unins*.exe' -ErrorAction SilentlyContinue |
        Sort-Object Name | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $exe -or -not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : $DetectionFileName"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName          = "Windows Admin Center (v2)"
        DisplayName      = "Windows Admin Center"
        Publisher        = "Microsoft Corporation"
        SoftwareVersion  = $version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = ("/VERYSILENT /HTTPSPortNumber={0}" -f $HttpsPortNumber)
        UninstallCommand = (Join-Path $InstallDir 'unins000.exe')
        UninstallArgs    = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = $DetectionFileName
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
        }
    }
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageWindowsAdminCenter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows Admin Center (v2) - PACKAGE phase"
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
    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

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
        $installer = Get-StagedInstaller -Quiet
        if (-not $installer) { exit 1 }
        Write-Output $installer.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Windows Admin Center GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows Admin Center (v2) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadUrl                  : $DownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageWindowsAdminCenter
    }
    elseif ($PackageOnly) {
        Invoke-PackageWindowsAdminCenter
    }
    else {
        Invoke-StageWindowsAdminCenter
        Invoke-PackageWindowsAdminCenter
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-windowsadmincenter'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
