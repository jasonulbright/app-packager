<#
Vendor: Microsoft
App: .NET 8 Desktop Runtime (x86+x64)
CMName: Microsoft Windows Desktop Runtime - 8
VendorUrl: https://dotnet.microsoft.com/download/dotnet/8.0
CPE: cpe:2.3:a:microsoft:.net:8.*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/dotnet/core/tree/main/release-notes/8.0
DownloadPageUrl: https://dotnet.microsoft.com/en-us/download/dotnet/8.0

.SYNOPSIS
    Packages .NET 8 Windows Desktop Runtime (x86 and x64) for MECM.

.DESCRIPTION
    Downloads the latest .NET 8 Windows Desktop Runtime installers for both
    x86 and x64 from the official Microsoft CDN, stages content to a versioned
    local folder with compound file-based detection metadata, and creates an
    MECM Application with file-based detection.
    Detection uses hostfxr.dll existence in the version-specific fxr path for
    both architectures (x86 AND x64).

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
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\DotNet8DesktopRuntime).
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
    create MECM application with compound file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available .NET 8 runtime version string and exits.

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
$ReleasesIndexUrl  = "https://builds.dotnet.microsoft.com/dotnet/release-metadata/releases-index.json"
$DownloadUrlBase   = "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop"

$VendorFolder = "Microsoft"
$AppFolder    = ".NET Core"

$X64FileNamePattern = "windowsdesktop-runtime-{0}-win-x64.exe"
$X86FileNamePattern = "windowsdesktop-runtime-{0}-win-x86.exe"

$BaseDownloadRoot = Join-Path $DownloadRoot "DotNet8DesktopRuntime"

# --- Functions ---


function Get-LatestDotNet8Version {
    param([switch]$Quiet)

    Write-Log "Releases index URL           : $ReleasesIndexUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleasesIndexUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET release info: $ReleasesIndexUrl" }

        $releases = ConvertFrom-Json $json
        $dotnet8Channel = $releases.'releases-index' |
            Where-Object { $_.'channel-version' -eq '8.0' -and $_.'release-type' -eq 'lts' } |
            Select-Object -First 1

        if (-not $dotnet8Channel -or -not $dotnet8Channel.'latest-runtime') {
            throw "Could not find .NET 8.0 LTS release channel or latest runtime."
        }

        $version = $dotnet8Channel.'latest-runtime'

        Write-Log "Latest .NET 8 runtime version: $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get .NET 8 version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDotNet8 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log ".NET 8 Desktop Runtime (x86+x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestDotNet8Version
    if (-not $version) { throw "Could not resolve .NET 8 runtime version." }

    $x64FileName = $X64FileNamePattern -f $version
    $x86FileName = $X86FileNamePattern -f $version

    Write-Log "Version                      : $version"
    Write-Log "x64 installer filename       : $x64FileName"
    Write-Log "x86 installer filename       : $x86FileName"
    Write-Log ""

    # --- Download x64 ---
    $localX64 = Join-Path $BaseDownloadRoot $x64FileName
    Write-Log "Local x64 installer path     : $localX64"

    if (-not (Test-Path -LiteralPath $localX64)) {
        $downloadUrl = "${DownloadUrlBase}/${version}/${x64FileName}"
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log "Downloading x64 installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localX64
    }
    else {
        Write-Log "Local x64 installer exists. Skipping download."
    }

    # --- Download x86 ---
    $localX86 = Join-Path $BaseDownloadRoot $x86FileName
    Write-Log "Local x86 installer path     : $localX86"

    if (-not (Test-Path -LiteralPath $localX86)) {
        $downloadUrl = "${DownloadUrlBase}/${version}/${x86FileName}"
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log "Downloading x86 installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localX86
    }
    else {
        Write-Log "Local x86 installer exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedX64 = Join-Path $localContentPath $x64FileName
    if (-not (Test-Path -LiteralPath $stagedX64)) {
        Copy-Item -LiteralPath $localX64 -Destination $stagedX64 -Force -ErrorAction Stop
        Write-Log "Copied x64 EXE to staged     : $stagedX64"
    }
    else {
        Write-Log "Staged x64 EXE exists. Skipping copy."
    }

    $stagedX86 = Join-Path $localContentPath $x86FileName
    if (-not (Test-Path -LiteralPath $stagedX86)) {
        Copy-Item -LiteralPath $localX86 -Destination $stagedX86 -Force -ErrorAction Stop
        Write-Log "Copied x86 EXE to staged     : $stagedX86"
    }
    else {
        Write-Log "Staged x86 EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$x64Path = Join-Path $PSScriptRoot ''{0}''' -f $x64FileName),
        ('$x86Path = Join-Path $PSScriptRoot ''{0}''' -f $x86FileName),
        '$proc1 = Start-Process -FilePath $x64Path -ArgumentList @(''/install'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'if ($proc1.ExitCode -ne 0) { exit $proc1.ExitCode }',
        '$proc2 = Start-Process -FilePath $x86Path -ArgumentList @(''/install'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc2.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$x64Path = Join-Path $PSScriptRoot ''{0}''' -f $x64FileName),
        ('$x86Path = Join-Path $PSScriptRoot ''{0}''' -f $x86FileName),
        '$proc1 = Start-Process -FilePath $x64Path -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'if ($proc1.ExitCode -ne 0) { exit $proc1.ExitCode }',
        '$proc2 = Start-Process -FilePath $x86Path -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc2.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $x86DetectionPath = "{0} (x86)\dotnet\host\fxr\{1}" -f $env:ProgramFiles, $version
    $x64DetectionPath = "{0}\dotnet\host\fxr\{1}" -f $env:ProgramFiles, $version

    $appName   = "Microsoft Windows Desktop Runtime - ${version}"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection x86 path           : $x86DetectionPath"
    Write-Log "Detection x64 path           : $x64DetectionPath"
    Write-Log "Detection file               : hostfxr.dll"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFiles  = @($x64FileName, $x86FileName)
        InstallerType   = "EXE"
        InstallArgs     = "/install /quiet /norestart"
        UninstallArgs   = "/uninstall /quiet /norestart"
        RunningProcess  = @()
        Detection       = @{
            Type      = "Compound"
            Connector = "And"
            Clauses   = @(
                @{
                    Type         = "File"
                    FilePath     = $x86DetectionPath
                    FileName     = "hostfxr.dll"
                    PropertyType = "Existence"
                    Is64Bit      = $false
                },
                @{
                    Type         = "File"
                    FilePath     = $x64DetectionPath
                    FileName     = "hostfxr.dll"
                    PropertyType = "Existence"
                    Is64Bit      = $true
                }
            )
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

function Invoke-PackageDotNet8 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log ".NET 8 Desktop Runtime (x86+x64) - PACKAGE phase"
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
        $v = Get-LatestDotNet8Version -Quiet
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
    Write-Log ".NET 8 Desktop Runtime (x86+x64) Auto-Packager starting"
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
        Invoke-StageDotNet8
    }
    elseif ($PackageOnly) {
        Invoke-PackageDotNet8
    }
    else {
        Invoke-StageDotNet8
        Invoke-PackageDotNet8
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
