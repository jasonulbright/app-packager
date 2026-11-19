<#
Vendor: Microsoft
App: PowerShell 7 LTS (MSI)
CMName: PowerShell 7
VendorUrl: https://github.com/PowerShell/PowerShell
CPE: cpe:2.3:a:microsoft:powershell:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/PowerShell/PowerShell/releases
DownloadPageUrl: https://github.com/PowerShell/PowerShell/releases
SecurityAdvisoryUrl: https://github.com/PowerShell/PowerShell/security/advisories
IconSource: Installer
WsusSupport: Yes
ArpName: PowerShell 7-x64

.SYNOPSIS
    Packages PowerShell 7 LTS (x64) MSI for ConfigMgr.

.DESCRIPTION
    This packager follows the Long Term Servicing (LTS) line of PowerShell 7. Microsoft
    Learn ("PowerShell Support Lifecycle") names PowerShell 7.6 as the current LTS
    release, supported until 14-Nov-2028. The packager takes the newest 7.6.x release
    that carries the x64 MSI and never a release of another line: a newer Stable
    release (PowerShell 7.7.0 publishes no MSI) is not packaged here. The line is
    the value of $LtsLine below; change it deliberately when the lifecycle page names
    a new LTS line.

    The three PowerShell 7 packagers of this tool, by release channel:
      - PowerShell 7 LTS (MSI), this packager: the LTS line, Microsoft MSI for all users,
        with remoting and the all-users scope. The choice for servers and for managed devices.
      - PowerShell 7 Stable (MSIX, per user), package-powershell7msix.ps1: the newest Stable release
        (7.7), the Microsoft MSIX package. It is single-user and has no PowerShell remoting.
      - PowerShell 7 (Devolutions, third party), package-devolutionspowershell.ps1: the
        MSI that Devolutions repackages from the Microsoft payload. Devolutions is not
        Microsoft; Microsoft supports only its own packages.

    Downloads the newest PowerShell 7 LTS x64 MSI from GitHub releases, stages
    content to a versioned local folder with ARP detection metadata, and
    creates a ConfigMgr Application with registry-based detection. The detection
    accepts the packaged version or a newer one, and the MSI opts the device into
    updates through Microsoft Update, WSUS and Configuration Manager (USE_MU=1), so
    a device that Microsoft Update moved to a newer version stays compliant.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    The install wrapper passes additional MSI properties to enable PATH
    registration, Explorer context menu, and PS Remoting.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\PowerShell 7 (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (for example, <DownloadRoot>\PowerShell7).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Queries the GitHub releases API for the newest PowerShell 7 LTS version, outputs
    the version string, and exits. No download or ConfigMgr changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (ConfigurationManager PowerShell module available)
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
$GitHubApiUrl = "https://api.github.com/repos/PowerShell/PowerShell/releases?per_page=100"

# The LTS line this packager follows. Microsoft Learn, "PowerShell Support Lifecycle", names PowerShell 7.6 as the LTS release.
$LtsLine = "7.6"

$VendorFolder = "Microsoft"
$AppFolder    = "PowerShell 7 (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "PowerShell7"

# --- Functions ---


function Select-PowerShell7LtsRelease {
    # The newest release of the LTS line that carries the x64 MSI. A release of another line is never a candidate, however new
    # it is or whether it carries an MSI. The asset list decides within the line; the file name is never built from the version.
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][AllowNull()][object[]]$Releases,
        [Parameter(Mandatory)][string]$Line,
        [Parameter(Mandatory)][string]$AssetPattern
    )
    $flat = @(); foreach ($entry in @($Releases)) { $flat += $entry }
    $linePattern = '^v' + [regex]::Escape($Line) + '\.\d+$'
    $inLine = @($flat | Where-Object { $_ -and ([string]$_.tag_name) -match $linePattern })
    return (Select-GitHubReleaseWithAsset -Releases $inLine -AssetPattern $AssetPattern)
}

function Get-PowerShell7Release {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet
    Write-Log "LTS line                     : $LtsLine" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $selected = Select-PowerShell7LtsRelease -Releases @(ConvertFrom-Json $json) -Line $LtsLine -AssetPattern '^PowerShell-.+-win-x64\.msi$'
        if (-not $selected) { throw "No PowerShell $LtsLine release carries an x64 MSI." }
        if ([string]::IsNullOrWhiteSpace($selected.Version)) { throw "Could not parse version from GitHub release." }

        if ($selected.SkippedNewer.Count -gt 0) {
            Write-Log ("PowerShell {0} of the {1} line publishes no MSI; packaging {2}." -f ($selected.SkippedNewer -join ', '), $LtsLine, $selected.Version) -Level WARN -Quiet:$Quiet
        }
        Write-Log "Latest PowerShell 7 LTS      : $($selected.Version)" -Quiet:$Quiet
        return $selected
    }
    catch {
        Write-Log "Failed to get PowerShell 7 LTS version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-LatestPowerShell7Version {
    param([switch]$Quiet)
    $selected = Get-PowerShell7Release -Quiet:$Quiet
    if (-not $selected) { return $null }
    return $selected.Version
}

# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePowerShell7 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 LTS (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-PowerShell7Release
    if (-not $release) { throw "Could not resolve PowerShell 7 version." }
    $version = $release.Version

    # The name and the URL come from the asset of the release, not from the version.
    $MsiFileName = $release.AssetName
    $downloadUrl = $release.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading PowerShell 7 MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    # The release lists no digest, so the signature is the proof that the file is the one Microsoft built.
    $signer = Assert-InstallerSignature -Path $localMsi -ExpectedSigner 'Microsoft Corporation' -RemoveOnFailure
    Write-Log "Authenticode signer          : $signer (valid)"

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."
    Write-Log ""
    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers (custom install args for PS7) ---
    $installPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''/quiet'', ''/norestart'', ''ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1'', ''ENABLE_PSREMOTING=1'', ''REGISTER_MANIFEST=1'', ''ADD_PATH=1'', ''USE_MU=1'', ''ENABLE_MU=1'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $publisher = "Microsoft Corporation"
    $appName = $productName

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/quiet /norestart ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 ADD_PATH=1 USE_MU=1 ENABLE_MU=1"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("pwsh")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackagePowerShell7 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 LTS (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $msiFiles = Get-ChildItem -Path $BaseDownloadRoot -Filter "PowerShell-*-win-x64.msi" -File
    if (-not $msiFiles -or $msiFiles.Count -eq 0) {
        throw "No staged PowerShell 7 MSI found - run Stage phase first."
    }
    $localMsi = $msiFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    $props = Get-MsiPropertyMap -MsiPath $localMsi.FullName
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $productVersion   = $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
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
        $v = Get-LatestPowerShell7Version -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
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
    Write-Log "PowerShell 7 (x64) Auto-Packager starting"
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
        Invoke-StagePowerShell7
    }
    elseif ($PackageOnly) {
        Invoke-PackagePowerShell7
    }
    else {
        Invoke-StagePowerShell7
        Invoke-PackagePowerShell7
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-powershell7'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
