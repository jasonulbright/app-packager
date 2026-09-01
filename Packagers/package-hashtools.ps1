<#
Vendor: Binary Fortress Software
App: HashTools
CMName: HashTools
VendorUrl: https://www.binaryfortress.com/HashTools/
CPE: cpe:2.3:a:binaryfortress:hashtools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.binaryfortress.com/HashTools/ChangeLog/
DownloadPageUrl: https://www.binaryfortress.com/HashTools/Download/
UpdateCadenceDays: 90

.SYNOPSIS
    Packages HashTools (x64) Inno Setup installer for MECM.

.DESCRIPTION
    Resolves the current HashTools build from the vendor download redirect,
    downloads the Setup executable, stages content to a versioned local folder,
    and creates an MECM Application with file-version detection.

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
    Content is staged under: <FileServerPath>\Applications\Binary Fortress Software\HashTools\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available HashTools version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$VendorDownloadUrl = "https://www.binaryfortress.com/Data/Download/?Package=hashtools"

$VendorFolder = "Binary Fortress Software"
$AppFolder    = "HashTools"

$BaseDownloadRoot = Join-Path $DownloadRoot "HashTools"

# The Inno installer defaults to a per-machine Program Files directory but the
# path is pinned with /DIR so file detection cannot drift from the install.
$InstallPath = "{0}\HashTools" -f $env:ProgramFiles

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        A CDN that answers 200 with an HTML error body would otherwise stage as
        a valid-looking installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestHashToolsRelease {
    <#
    .SYNOPSIS
        Returns the current HashTools version and its installer URL.
    .DESCRIPTION
        The vendor exposes no version endpoint. The download entry point answers
        302 to a versioned file on the download host, so the redirect chain is
        followed without transferring the body and the version is taken from the
        resolved file name. The resolved URL is returned rather than the entry
        point so the staged payload and the reported version cannot come from
        two different builds if a release lands mid-run.
    #>
    param([switch]$Quiet)

    Write-Log "Vendor download entry point  : $VendorDownloadUrl" -Quiet:$Quiet

    try {
        $resolved = (curl.exe --silent --head --location --fail --output NUL -w "%{url_effective}" $VendorDownloadUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve the vendor download redirect." }

        $fileName = [System.IO.Path]::GetFileName(([uri]$resolved).AbsolutePath)
        $m = [regex]::Match($fileName, '(?i)^HashToolsSetup-(?<ver>\d+(\.\d+)+)\.exe$')
        if (-not $m.Success) {
            throw "Resolved download does not look like a HashTools installer: $fileName"
        }

        Write-Log "Latest HashTools version     : $($m.Groups['ver'].Value)" -Quiet:$Quiet
        Write-Log "Installer file               : $fileName" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $m.Groups['ver'].Value
            FileName    = $fileName
            DownloadUrl = $resolved
        }
    }
    catch {
        Write-Log "Failed to get HashTools version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageHashTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HashTools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestHashToolsRelease
    if (-not $releaseInfo) { throw "Could not resolve HashTools version." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExePayload -Path $localExe

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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs ("'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/SP-', '/DIR={0}'" -f $InstallPath) `
        -UninstallCommand ("{0}\unins000.exe" -f $InstallPath) `
        -UninstallArgs "'/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall `
        -InstallBatExitCode '3010' `
        -UninstallBatExitCode '3010'

    # --- Write stage manifest ---
    Write-Log ""
    Write-Log "Detection path               : $InstallPath"
    Write-Log "Detection file               : HashTools.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "HashTools"
        Publisher       = "Binary Fortress Software"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP- /DIR=$InstallPath"
        UninstallArgs   = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES"
        RunningProcess  = @("HashTools")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallPath
            FileName      = "HashTools.exe"
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

function Invoke-PackageHashTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HashTools (x64) - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

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
        $info = Get-LatestHashToolsRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("HashTools GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "HashTools (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VendorDownloadUrl            : $VendorDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageHashTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageHashTools
    }
    else {
        Invoke-StageHashTools
        Invoke-PackageHashTools
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-hashtools'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
