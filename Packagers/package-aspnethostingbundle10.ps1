<#
Vendor: Microsoft
App: ASP.NET 10 Server Hosting Bundle (x64)
CMName: Microsoft .NET 10
VendorUrl: https://dotnet.microsoft.com/download/dotnet/10.0
CPE: cpe:2.3:a:microsoft:asp.net_core:10.*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/dotnet/core/tree/main/release-notes/10.0
DownloadPageUrl: https://dotnet.microsoft.com/en-us/download/dotnet/10.0
IconSource: None

.SYNOPSIS
    Packages ASP.NET 10 Server Hosting Bundle for MECM.

.DESCRIPTION
    Downloads the latest .NET 10 ASP.NET Core Windows Server Hosting Bundle
    installer from the official Microsoft CDN, stages content to a versioned
    local folder, and creates an MECM Application with file existence
    detection.
    Detection uses the versioned shared framework file that the runtime
    installs, because .NET 10 no longer writes a per-version ASP.NET Core
    Shared Framework registry key.

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\.NET Core\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\ASPNETHostingBundle10).
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
    create MECM application with file existence detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available .NET 10 runtime version string and exits.

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
$ReleasesIndexUrl  = "https://builds.dotnet.microsoft.com/dotnet/release-metadata/releases-index.json"
$DownloadUrlBase   = "https://builds.dotnet.microsoft.com/dotnet/aspnetcore/Runtime"

$VendorFolder = "Microsoft"
$AppFolder    = "ASP.NET Core Hosting Bundle"

$InstallerFileNamePattern = "dotnet-hosting-{0}-win.exe"

$BaseDownloadRoot = Join-Path $DownloadRoot "ASPNETHostingBundle10"

# --- Functions ---


function Get-LatestDotNet10Version {
    param([switch]$Quiet)

    Write-Log "Releases index URL           : $ReleasesIndexUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleasesIndexUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET release info: $ReleasesIndexUrl" }

        $releases = ConvertFrom-Json $json
        $dotnet10Channel = $releases.'releases-index' |
            Where-Object { $_.'channel-version' -eq '10.0' -and $_.'release-type' -eq 'lts' } |
            Select-Object -First 1

        if (-not $dotnet10Channel -or -not $dotnet10Channel.'latest-runtime') {
            throw "Could not find .NET 10.0 LTS release channel or latest runtime."
        }

        $version = $dotnet10Channel.'latest-runtime'

        Write-Log "Latest .NET 10 runtime version: $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get .NET 10 version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageASPNETHostingBundle10 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "ASP.NET 10 Hosting Bundle - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestDotNet10Version
    if (-not $version) { throw "Could not resolve .NET 10 runtime version." }

    $installerFileName = $InstallerFileNamePattern -f $version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        $downloadUrl = "${DownloadUrlBase}/${version}/${installerFileName}"
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/install'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # .NET 10 does not write the per-version "ASP.NET Core\Shared Framework"
    # key that 8.0 registers, and the dotnet InstalledVersions values outlive
    # the runtime they name; the versioned shared-framework folder does not.
    $sharedFrameworkPath = "C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App\${version}"
    $sharedFrameworkFile = "Microsoft.AspNetCore.dll"

    $appName   = "Microsoft .NET ${version} - Windows Server Hosting"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection file               : $sharedFrameworkPath\$sharedFrameworkFile"

    # A client that already took the next patch still counts as installed:
    # the runtime folder is replaced, not kept beside the old one.
    $detection = @{
        Type         = "File"
        FilePath     = $sharedFrameworkPath
        FileName     = $sharedFrameworkFile
        PropertyType = "Existence"
        Is64Bit      = $true
    }
    $nextVersion = Get-NextPatchVersion -Version $version
    if ($nextVersion) {
        Write-Log "Detection also accepts       : $nextVersion (successor patch)"
        $detection = @{
            Type       = "Compound"
            Connector  = "Or"
            GroupSizes = @(1, 1)
            Clauses    = @(
                @{
                    Type         = "File"
                    FilePath     = $sharedFrameworkPath
                    FileName     = $sharedFrameworkFile
                    PropertyType = "Existence"
                    Is64Bit      = $true
                },
                @{
                    Type         = "File"
                    FilePath     = "C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App\${nextVersion}"
                    FileName     = $sharedFrameworkFile
                    PropertyType = "Existence"
                    Is64Bit      = $true
                }
            )
        }
    }
    else {
        Write-Log "Non-numeric patch component in '$version'; single-version detection only." -Level WARN
    }
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/install /quiet /norestart"
        UninstallArgs   = "/uninstall /quiet /norestart"
        RunningProcess  = @()
        Detection       = $detection
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

function Invoke-PackageASPNETHostingBundle10 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "ASP.NET 10 Hosting Bundle - PACKAGE phase"
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
        $v = Get-LatestDotNet10Version -Quiet
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
    Write-Log "ASP.NET 10 Hosting Bundle Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleasesIndexUrl             : $ReleasesIndexUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageASPNETHostingBundle10
    }
    elseif ($PackageOnly) {
        Invoke-PackageASPNETHostingBundle10
    }
    else {
        Invoke-StageASPNETHostingBundle10
        Invoke-PackageASPNETHostingBundle10
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-aspnethostingbundle10'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
