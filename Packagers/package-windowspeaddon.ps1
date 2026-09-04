<#
Vendor: Microsoft
App: Windows PE add-on for the Windows ADK
CMName: Windows PE add-on for the Windows ADK
VendorUrl: https://learn.microsoft.com/windows-hardware/get-started/adk-install
CPE: cpe:2.3:a:microsoft:windows_assessment_and_deployment_kit:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/windows-hardware/get-started/what-s-new-in-kits-and-tools
DownloadPageUrl: https://learn.microsoft.com/windows-hardware/get-started/adk-install
IconSource: External
UpdateCadenceDays: 180

.SYNOPSIS
    Packages the Windows PE add-on for the Windows ADK for MECM as an
    offline layout.

.DESCRIPTION
    Downloads adkwinpesetup.exe, then runs it with /quiet /layout so the
    staged content is a complete offline install set rather than a
    bootstrapper that would need internet access from every client. The
    layout carries adkwinpesetup.exe at its root, so the install wrapper runs
    the copied bootstrapper and it consumes the local Installers folder.

    A layout is not feature-filterable: the bootstrapper downloads the whole
    add-on (about 1.9 GB, most of it the boot images) regardless of the
    install selection.

    The add-on must be installed on top of a matching Windows ADK release and
    into the same kit root. Pair this application with the ADK application as
    a dependency; nothing in the add-on installer enforces the ordering.

    Detection is derived at stage time from the layout's own Windows PE boot
    image MSI: its ProductVersion equals the kit version and its ProductCode
    is the ARP key Windows writes. The feature MSIs are 32-bit packages, so
    the entry lands in the WOW6432 uninstall hive.

    Supports two-phase operation:
      -StageOnly    Download bootstrapper, build offline layout, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Windows PE Add-on\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER InstallPath
    Kit install root passed to the setup /installpath switch. Must match the
    root the Windows ADK was installed to.
    Default: C:\Program Files (x86)\Windows Kits\10

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase: download bootstrapper, build the offline
    layout, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the current add-on version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - About 2.5 GB free under DownloadRoot for the layout
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [ValidateSet('Nested','Flat')]
    [string]$ContentLayout = "Nested",
    [string]$DownloadRoot = "C:\temp\ap",
    [string]$InstallPath = "C:\Program Files (x86)\Windows Kits\10",
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
# Permanent vendor link for the Windows PE add-on that pairs with the ADK
# release supporting Windows 11 25H2, 24H2 and earlier supported builds.
$DownloadUrl       = "https://go.microsoft.com/fwlink/?linkid=2289981"
$InstallerFileName = "adkwinpesetup.exe"

# The add-on installs exactly one selectable feature.
$Features          = @('OptionId.WindowsPreinstallationEnvironment')

$DetectionMsiName  = "Windows PE wims (DesktopEditions)-x86_en-us.msi"

$VendorFolder = "Microsoft"
$AppFolder    = "Windows PE Add-on"

$BaseDownloadRoot = Join-Path $DownloadRoot "WindowsPEAddon"

# The bundle refuses a layout target that sits under the folder it is running
# from (exit 2002), so the bootstrapper is kept out of the staging root.
$BootstrapperFolder = Join-Path $BaseDownloadRoot "bootstrapper"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The fwlink redirect can land on an error document that answers 200
        with HTML, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-StagedBootstrapper {
    <#
    .SYNOPSIS
        Downloads adkwinpesetup.exe once and returns its path and version.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BootstrapperFolder

    $localExe = Join-Path $BootstrapperFolder $InstallerFileName
    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $DownloadUrl" -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe -Quiet:$Quiet
    }
    else {
        Write-Log "Local bootstrapper exists. Skipping download." -Quiet:$Quiet
    }

    Assert-PayloadIsExecutable -Path $localExe

    $version = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "$InstallerFileName carries no file version; cannot derive the kit version."
    }
    $version = $version.Trim()

    Write-Log "Windows PE add-on version    : $version" -Quiet:$Quiet

    return [pscustomobject]@{ Path = $localExe; Version = $version }
}


function New-WinPeOfflineLayout {
    <#
    .SYNOPSIS
        Runs the add-on bootstrapper with /quiet /layout into the target folder.
    .DESCRIPTION
        A layout that already carries the root bootstrapper and an Installers
        folder is left alone; re-running re-verifies every payload.
    #>
    param(
        [Parameter(Mandatory)][string]$BootstrapperPath,
        [Parameter(Mandatory)][string]$LayoutPath
    )

    $layoutSetup = Join-Path $LayoutPath $InstallerFileName
    $layoutInstallers = Join-Path $LayoutPath "Installers"
    if ((Test-Path -LiteralPath $layoutSetup) -and (Test-Path -LiteralPath $layoutInstallers)) {
        Write-Log "Offline layout exists. Skipping /layout."
        return
    }

    # The bundle validates the layout root against the longest payload name it
    # will write and refuses with exit 2002 ("Invalid download path") when the
    # result would approach MAX_PATH. Fail here instead, with the cause named.
    if ($LayoutPath.Length -gt 120) {
        throw "Layout path is too long for the bundle ($($LayoutPath.Length) chars): $LayoutPath. Use a shorter -DownloadRoot."
    }

    Write-Log "Building offline layout (this downloads the full add-on)..."
    # The bootstrapper takes the Windows Installer mutex even for a layout and
    # answers 1618 (ERROR_INSTALL_ALREADY_RUNNING) while another setup holds it.
    $attempt = 0
    do {
        $attempt++
        $proc = Start-Process -FilePath $BootstrapperPath `
            -ArgumentList @('/quiet', '/layout', $LayoutPath) -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 1618 -and $attempt -lt 6) {
            Write-Log "Windows Installer is busy (exit 1618); retrying the layout in 60 seconds (attempt $attempt of 6)." -Level WARN
            Start-Sleep -Seconds 60
        }
    } while ($proc.ExitCode -eq 1618 -and $attempt -lt 6)
    if ($proc.ExitCode -ne 0) {
        throw "$InstallerFileName /layout failed with exit code $($proc.ExitCode)."
    }

    if (-not (Test-Path -LiteralPath $layoutSetup)) {
        throw "Layout completed but $InstallerFileName is missing from $LayoutPath."
    }
    if (-not (Test-Path -LiteralPath $layoutInstallers)) {
        throw "Layout completed but the Installers folder is missing from $LayoutPath."
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageWinPeAddon {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows PE add-on for the Windows ADK - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $bootstrapper = Get-StagedBootstrapper
    $version = $bootstrapper.Version

    Write-Log "Version                      : $version"
    Write-Log ("Features                     : {0}" -f ($Features -join ', '))
    Write-Log "Install path                 : $InstallPath"
    Write-Log ""

    # --- Versioned local content folder holds the layout itself ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    New-WinPeOfflineLayout -BootstrapperPath $bootstrapper.Path -LayoutPath $localContentPath

    $layoutSize = (Get-ChildItem -LiteralPath $localContentPath -File -Recurse -ErrorAction Stop |
        Measure-Object -Property Length -Sum).Sum
    Write-Log ("Layout size                  : {0:N0} MB" -f ($layoutSize / 1MB))

    # --- Derive ARP detection from the layout's feature MSI ---
    $detectionMsiPath = Join-Path (Join-Path $localContentPath "Installers") $DetectionMsiName
    if (-not (Test-Path -LiteralPath $detectionMsiPath)) {
        throw "Detection anchor missing from the layout: $detectionMsiPath"
    }

    $props = Get-MsiPropertyMap -MsiPath $detectionMsiPath
    $productVersion = $props["ProductVersion"]
    $productCode    = $props["ProductCode"]
    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "Detection MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "Detection MSI ProductCode missing." }

    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log ""
    Write-Log "Detection MSI                : $DetectionMsiName"
    Write-Log "Detection ProductVersion     : $productVersion"
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log ""

    # --- Generate content wrappers ---
    $installArgList = @("'/quiet'", "'/installpath'", ("'{0}'" -f ($InstallPath -replace "'", "''")), "'/features'")
    foreach ($feature in $Features) {
        $installArgList += ("'{0}'" -f $feature)
    }
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs ($installArgList -join ', ') `
        -UninstallCommand 'unused'

    # The bundle uninstalls itself; the copy in the content folder is the same
    # build, so it removes what it installed without needing the ARP entry.
    $uninstallContent = (
        ('$setupPath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        'if (-not (Test-Path -LiteralPath $setupPath)) { exit 0 }',
        '$proc = Start-Process -FilePath $setupPath -ArgumentList @(''/quiet'', ''/uninstall'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName          = "Windows PE add-on for the Windows ADK $version"
        DisplayName      = "Windows PE add-on for the Windows ADK"
        Publisher        = "Microsoft Corporation"
        SoftwareVersion  = $version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = ("/quiet /installpath ""{0}"" /features {1}" -f $InstallPath, ($Features -join ' '))
        UninstallCommand = $InstallerFileName
        UninstallArgs    = "/quiet /uninstall"
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersion
            Is64Bit             = $false
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

function Invoke-PackageWinPeAddon {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows PE add-on for the Windows ADK - PACKAGE phase"
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

    # --- Copy staged content to network (the layout is a tree) ---
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
        $bootstrapper = Get-StagedBootstrapper -Quiet
        if (-not $bootstrapper) { exit 1 }
        Write-Output $bootstrapper.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Windows PE add-on GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows PE add-on for the Windows ADK Auto-Packager starting"
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
        Invoke-StageWinPeAddon
    }
    elseif ($PackageOnly) {
        Invoke-PackageWinPeAddon
    }
    else {
        Invoke-StageWinPeAddon
        Invoke-PackageWinPeAddon
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-windowspeaddon'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
