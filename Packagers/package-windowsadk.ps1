<#
Vendor: Microsoft
App: Windows ADK for Windows 11
CMName: Windows ADK for Windows 11
VendorUrl: https://learn.microsoft.com/windows-hardware/get-started/adk-install
CPE: cpe:2.3:a:microsoft:windows_assessment_and_deployment_kit:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/windows-hardware/get-started/what-s-new-in-kits-and-tools
DownloadPageUrl: https://learn.microsoft.com/windows-hardware/get-started/adk-install
IconSource: External
UpdateCadenceDays: 180

.SYNOPSIS
    Packages the Windows ADK for Windows 11 for MECM as an offline layout.

.DESCRIPTION
    Downloads adksetup.exe, then runs it with /quiet /layout so the staged
    content is a complete offline install set rather than a bootstrapper that
    would need internet access from every client. The layout carries
    adksetup.exe at its root, so the install wrapper runs the copied
    bootstrapper and it consumes the local Installers folder.

    A layout is not feature-filterable: adksetup downloads the whole kit
    (about 1.5 GB) regardless of which features the install selects. That is
    the price of an offline-capable package and it is paid once per release.

    Detection is derived at stage time from the layout's own feature MSI for
    Deployment Tools: its ProductVersion equals the kit version and its
    ProductCode is the ARP key Windows writes, so no temp install is needed.
    The feature MSIs are 32-bit packages, so the entry lands in the WOW6432
    uninstall hive.

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\Windows ADK\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER Features
    ADK feature option IDs to install. OptionId.DeploymentTools must stay in
    the set: detection is anchored to it.
    Default: OptionId.DeploymentTools, OptionId.UserStateMigrationTool

.PARAMETER InstallPath
    Kit install root passed to adksetup /installpath. The Windows PE add-on
    must be installed to the same root.
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
    Outputs only the current ADK version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - About 2 GB free under DownloadRoot for the layout
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [ValidateSet('Nested','Flat')]
    [string]$ContentLayout = "Nested",
    [string]$DownloadRoot = "C:\temp\ap",
    [string[]]$Features = @('OptionId.DeploymentTools', 'OptionId.UserStateMigrationTool'),
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
# Permanent vendor link for the ADK release that supports Windows 11 25H2,
# 24H2 and every earlier supported Windows 10/11 build.
$DownloadUrl       = "https://go.microsoft.com/fwlink/?linkid=2289980"
$InstallerFileName = "adksetup.exe"

# Deployment Tools ships in every supported feature selection this packager
# offers, and its feature MSI carries the kit version.
$DetectionMsiName  = "Windows Deployment Tools-x86_en-us.msi"
$DetectionFeature  = "OptionId.DeploymentTools"

$VendorFolder = "Microsoft"
$AppFolder    = "Windows ADK"

$BaseDownloadRoot = Join-Path $DownloadRoot "WindowsADK"

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
        Downloads adksetup.exe once and returns its path and kit version.
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
        throw "adksetup.exe carries no file version; cannot derive the kit version."
    }
    $version = $version.Trim()

    Write-Log "Windows ADK version          : $version" -Quiet:$Quiet

    return [pscustomobject]@{ Path = $localExe; Version = $version }
}


function New-AdkOfflineLayout {
    <#
    .SYNOPSIS
        Runs adksetup /quiet /layout into the target folder.
    .DESCRIPTION
        The bootstrapper is idempotent over an existing layout but re-verifies
        every payload, so a layout that already carries the root bootstrapper
        and an Installers folder is left alone.
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

    Write-Log "Building offline layout (this downloads the full kit)..."
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
        throw "adksetup /layout failed with exit code $($proc.ExitCode)."
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

function Invoke-StageWindowsAdk {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows ADK for Windows 11 - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if ($Features -notcontains $DetectionFeature) {
        throw "$DetectionFeature must stay in -Features: detection is anchored to its feature MSI."
    }

    $bootstrapper = Get-StagedBootstrapper
    $version = $bootstrapper.Version

    Write-Log "Version                      : $version"
    Write-Log ("Features                     : {0}" -f ($Features -join ', '))
    Write-Log "Install path                 : $InstallPath"
    Write-Log ""

    # --- Versioned local content folder holds the layout itself ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    New-AdkOfflineLayout -BootstrapperPath $bootstrapper.Path -LayoutPath $localContentPath

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

    # The feature MSIs are 32-bit packages, so Windows writes their ARP entry
    # to the WOW6432 uninstall hive.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log ""
    Write-Log "Detection MSI                : $DetectionMsiName"
    Write-Log "Detection ProductVersion     : $productVersion"
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log ""

    # --- Generate content wrappers ---
    $installArgList = @("'/quiet'", ("'/installpath'"), ("'{0}'" -f ($InstallPath -replace "'", "''")), "'/features'")
    foreach ($feature in $Features) {
        $installArgList += ("'{0}'" -f ($feature -replace "'", "''"))
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
        AppName          = "Windows ADK for Windows 11 $version"
        DisplayName      = "Windows ADK for Windows 11"
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

function Invoke-PackageWindowsAdk {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows ADK for Windows 11 - PACKAGE phase"
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
        [Console]::Error.WriteLine("Windows ADK GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Windows ADK for Windows 11 Auto-Packager starting"
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
        Invoke-StageWindowsAdk
    }
    elseif ($PackageOnly) {
        Invoke-PackageWindowsAdk
    }
    else {
        Invoke-StageWindowsAdk
        Invoke-PackageWindowsAdk
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-windowsadk'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
