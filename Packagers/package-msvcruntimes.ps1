<#
Vendor: Microsoft
App: Microsoft Visual C++ v14 Redistributable (x86+x64)
CMName: Microsoft Visual C++ v14 Redistributable
VendorUrl: https://learn.microsoft.com/cpp/windows/latest-supported-vc-redist
CPE: cpe:2.3:a:microsoft:visual_c%2b%2b_redistributable:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist
DownloadPageUrl: https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist

.SYNOPSIS
    Packages Microsoft Visual C++ v14 Redistributable (x86+x64) for MECM.

.DESCRIPTION
    Downloads the latest vc_redist.x86.exe and vc_redist.x64.exe from Microsoft's
    permalink URLs, reads the version from VersionInfo on the x64 installer, stages
    both to a versioned local folder with dual-registry detection metadata, and
    creates an MECM Application with compound registry detection (AND logic).
    Detection uses HKLM registry Version string under both X86 and X64
    VC\Runtimes keys must equal the expected value (vMAJOR.MINOR.BUILD.00).

    NOTE: Microsoft renamed this product from "2015-2022" to "v14 Redistributable"
    starting with version 14.50.x. The ARP DisplayName is now
    "Microsoft Visual C++ v14 Redistributable (x64/x86) - VERSION".

    NOTE: The aka.ms permalink URLs always serve the current release. Installers
    are always re-downloaded to ensure the latest version is packaged.

    GetLatestVersionOnly downloads only the x64 installer to a local staging
    folder, reads the version from VersionInfo, outputs the short version string,
    and exits.

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\VC++ 2015-2022 Redistributable\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\MsvcRedist).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installers, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with compound registry detection.

.PARAMETER GetLatestVersionOnly
    Downloads the x64 installer, reads the version from VersionInfo, outputs the
    short version string, and exits. No MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$UrlX86 = "https://aka.ms/vc14/vc_redist.x86.exe"
$UrlX64 = "https://aka.ms/vc14/vc_redist.x64.exe"

$FileNameX86 = "vc_redist.x86.exe"
$FileNameX64 = "vc_redist.x64.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "VC++ v14 Redistributable"

$BaseDownloadRoot = Join-Path $DownloadRoot "MsvcRedist"

# Registry detection paths (stable across 2015-2022+ releases)
$RegKeyX86    = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86"
$RegKeyX64    = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64"
$RegValueName = "Version"

# --- Functions ---


function Get-ExeFileVersion {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Quiet
    )

    $vi = (Get-Item -LiteralPath $Path).VersionInfo
    if (-not $vi) { throw "Could not read VersionInfo from: $Path" }

    $fv = $vi.FileVersion
    $pv = $vi.ProductVersion

    if (-not [string]::IsNullOrWhiteSpace($fv)) { $fv = $fv.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($pv)) { $pv = $pv.Trim() }

    Write-Log "EXE FileVersion              : $fv" -Quiet:$Quiet
    Write-Log "EXE ProductVersion           : $pv" -Quiet:$Quiet

    if ($fv -match '^\d+\.\d+\.\d+\.\d+$') { return $fv }
    if ($pv -match '^\d+\.\d+\.\d+\.\d+$') { return $pv }

    throw "Could not determine quad version from VersionInfo for: $Path"
}

function Get-ShortVersionFromQuad {
    param([Parameter(Mandatory)][string]$QuadVersion)

    $parts = $QuadVersion -split '\.'
    if ($parts.Count -lt 3) { throw "Unexpected version format: $QuadVersion" }
    return ("{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2])
}

function Convert-ShortVersionToRegExpected {
    param([Parameter(Mandatory)][string]$ShortVersion)

    return ("v{0}.00" -f $ShortVersion)
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageMsvcRedist {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC 2015-2022 Redistributable (x86+x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # Always re-download (permalink URLs serve latest release)
    $localX86 = Join-Path $BaseDownloadRoot $FileNameX86
    $localX64 = Join-Path $BaseDownloadRoot $FileNameX64

    Write-Log "Downloading x86 installer..."
    Invoke-DownloadWithRetry -Url $UrlX86 -OutFile $localX86

    Write-Log "Downloading x64 installer..."
    Invoke-DownloadWithRetry -Url $UrlX64 -OutFile $localX64

    # Version from x64 installer VersionInfo
    $quadVersion  = Get-ExeFileVersion -Path $localX64
    $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion
    $regExpected  = Convert-ShortVersionToRegExpected -ShortVersion $shortVersion

    Write-Log ""
    Write-Log "Version (short)              : $shortVersion"
    Write-Log "Version (quad)               : $quadVersion"
    Write-Log "Registry expected            : $regExpected"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $shortVersion
    Initialize-Folder -Path $localContentPath

    $stagedX86 = Join-Path $localContentPath $FileNameX86
    if (-not (Test-Path -LiteralPath $stagedX86)) {
        Copy-Item -LiteralPath $localX86 -Destination $stagedX86 -Force -ErrorAction Stop
        Write-Log "Copied x86 EXE to staged     : $stagedX86"
    }
    else {
        Write-Log "Staged x86 EXE exists. Skipping copy."
    }

    $stagedX64 = Join-Path $localContentPath $FileNameX64
    if (-not (Test-Path -LiteralPath $stagedX64)) {
        Copy-Item -LiteralPath $localX64 -Destination $stagedX64 -Force -ErrorAction Stop
        Write-Log "Copied x64 EXE to staged     : $stagedX64"
    }
    else {
        Write-Log "Staged x64 EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$x86Path = Join-Path $PSScriptRoot ''{0}''' -f $FileNameX86),
        ('$x64Path = Join-Path $PSScriptRoot ''{0}''' -f $FileNameX64),
        '$proc1 = Start-Process -FilePath $x86Path -ArgumentList @(''/install'', ''/quiet'', ''/norestart'', ''/log'', (Join-Path $PSScriptRoot ''x86.install.log'')) -Wait -PassThru -NoNewWindow',
        'if ($proc1.ExitCode -ne 0) { exit $proc1.ExitCode }',
        '$proc2 = Start-Process -FilePath $x64Path -ArgumentList @(''/install'', ''/quiet'', ''/norestart'', ''/log'', (Join-Path $PSScriptRoot ''x64.install.log'')) -Wait -PassThru -NoNewWindow',
        'exit $proc2.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$x86Path = Join-Path $PSScriptRoot ''{0}''' -f $FileNameX86),
        ('$x64Path = Join-Path $PSScriptRoot ''{0}''' -f $FileNameX64),
        '$proc1 = Start-Process -FilePath $x86Path -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'', ''/log'', (Join-Path $PSScriptRoot ''x86.uninstall.log'')) -Wait -PassThru -NoNewWindow',
        'if ($proc1.ExitCode -ne 0) { exit $proc1.ExitCode }',
        '$proc2 = Start-Process -FilePath $x64Path -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'', ''/log'', (Join-Path $PSScriptRoot ''x64.uninstall.log'')) -Wait -PassThru -NoNewWindow',
        'exit $proc2.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent `
        -InstallBatExitCode '3010' `
        -UninstallBatExitCode '3010'

    # --- Write stage manifest ---
    $appName   = "Microsoft Visual C++ v14 Redistributable (x86+x64) - $shortVersion"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection (AND)              : HKLM\$RegKeyX86\$RegValueName == $regExpected"
    Write-Log "Detection (AND)              : HKLM\$RegKeyX64\$RegValueName == $regExpected"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName              = $appName
        Publisher            = $publisher
        SoftwareVersion      = $shortVersion
        InstallerFiles       = @($FileNameX86, $FileNameX64)
        InstallerType        = "EXE"
        InstallArgs          = "/install /quiet /norestart"
        UninstallArgs        = "/uninstall /quiet /norestart"
        RunningProcess       = @()
        PostExecutionBehavior = "ForceReboot"
        Detection            = @{
            Type      = "Compound"
            Connector = "And"
            Clauses   = @(
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $RegKeyX86
                    ValueName           = $RegValueName
                    PropertyType        = "String"
                    ExpectedValue       = $regExpected
                    Operator            = "IsEquals"
                },
                @{
                    Type                = "RegistryKeyValue"
                    RegistryKeyRelative = $RegKeyX64
                    ValueName           = $RegValueName
                    PropertyType        = "String"
                    ExpectedValue       = $regExpected
                    Operator            = "IsEquals"
                }
            )
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $shortVersion -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageMsvcRedist {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC 2015-2022 Redistributable (x86+x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

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
    Write-Log "PostExecutionBehavior        : $($manifest.PostExecutionBehavior)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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
        Initialize-Folder -Path $BaseDownloadRoot

        $localX64 = Join-Path $BaseDownloadRoot $FileNameX64

        Invoke-DownloadWithRetry -Url $UrlX64 -OutFile $localX64 -Quiet

        $quadVersion = Get-ExeFileVersion -Path $localX64 -Quiet
        $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion

        Write-Output $shortVersion
        exit 0
    }
    catch {
        Write-Log $_.Exception.Message -Level ERROR
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "MSVC 2015-2022 Redistributable (x86+x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "UrlX86                       : $UrlX86"
    Write-Log "UrlX64                       : $UrlX64"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageMsvcRedist
    }
    elseif ($PackageOnly) {
        Invoke-PackageMsvcRedist
    }
    else {
        Invoke-StageMsvcRedist
        Invoke-PackageMsvcRedist
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
