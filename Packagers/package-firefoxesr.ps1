<#
Vendor: Mozilla
App: Mozilla Firefox ESR
CMName: Mozilla Firefox ESR
VendorUrl: https://www.mozilla.org/firefox/enterprise/
CPE: cpe:2.3:a:mozilla:firefox_esr:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.mozilla.org/en-US/firefox/organizations/notes/
DownloadPageUrl: https://www.mozilla.org/en-US/firefox/enterprise/
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Mozilla Firefox ESR (x64) MSI for MECM.

.DESCRIPTION
    Resolves the current ESR line from the Mozilla product-details feed,
    downloads the matching x64 en-US MSI from the official release servers,
    stages content to a versioned local folder, and creates an MECM
    Application with file-version-based detection.

    The ESR line is whatever FIREFOX_ESR names in the feed, not a pinned
    major version; FIREFOX_ESR115 and FIREFOX_ESR_NEXT are deliberately
    ignored so the packager follows the supported line as it rolls forward.

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
    Content is staged under: <FileServerPath>\Applications\Mozilla\Firefox ESR\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download MSI, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Firefox ESR version string and exits.

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
$VersionsJsonUrl = "https://product-details.mozilla.org/1.0/firefox_versions.json"
$DownloadBase    = "https://releases.mozilla.org/pub/firefox/releases"

$VendorFolder = "Mozilla"
$AppFolder    = "Firefox ESR"

$BaseDownloadRoot = Join-Path $DownloadRoot "FirefoxESR"

# --- Functions ---


function Assert-PayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-file
        signature that every Windows Installer package carries.
    .DESCRIPTION
        The release CDN answers 200 with an HTML error body for a path that
        no longer exists, which would otherwise stage as a valid-looking MSI
        and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $sig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
    if ($bytes.Count -lt 8) { throw "Downloaded payload is too small to be an MSI: $Path" }
    for ($i = 0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $sig[$i]) { throw "Downloaded payload is not an MSI (no OLE header): $Path" }
    }
}


function Get-LatestFirefoxEsrVersion {
    param([switch]$Quiet)

    Write-Log "Versions JSON URL            : $VersionsJsonUrl" -Quiet:$Quiet

    try {
        $jsonText = (curl.exe -L --fail --silent --show-error $VersionsJsonUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Firefox version info: $VersionsJsonUrl" }

        $json = ConvertFrom-Json $jsonText
        $version = $json.FIREFOX_ESR
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "FIREFOX_ESR field was empty."
        }

        Write-Log "Latest Firefox ESR version   : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Firefox ESR version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-EsrFileVersion {
    <#
    .SYNOPSIS
        Strips the 'esr' suffix so the value compares as a file version.
    .DESCRIPTION
        firefox.exe carries 140.14.0 while the release train is named
        140.14.0esr; a detection clause comparing versions rejects the suffix.
    #>
    param([Parameter(Mandatory)][string]$Version)
    return ($Version -replace 'esr$', '')
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageFirefoxEsr {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Mozilla Firefox ESR (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestFirefoxEsrVersion
    if (-not $version) { throw "Could not resolve Firefox ESR version." }

    $fileVersion = Get-EsrFileVersion -Version $version

    $msiFileName = "Firefox Setup $version.msi"
    $downloadUrl = "$DownloadBase/$version/win64/en-US/" + ($msiFileName -replace ' ', '%20')

    Write-Log "Version                      : $version"
    Write-Log "Detection file version       : $fileVersion"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-PayloadIsMsi -Path $localMsi

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrappers = New-MsiWrapperContent -MsiFileName $msiFileName

    # The Firefox MSI is a thin wrapper around the EXE installer. msiexec /x
    # returns 1605 because the product isn't registered as an MSI install.
    # Use the native uninstaller (helper.exe /S) instead.
    $customUninstall = (
        '$helperPath = Join-Path $env:ProgramFiles ''Mozilla Firefox\uninstall\helper.exe''',
        'if (-not (Test-Path -LiteralPath $helperPath)) { exit 0 }',
        '$proc = Start-Process -FilePath $helperPath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $customUninstall

    # --- Write stage manifest ---
    # The ESR MSI installs to the same Program Files path as the rapid
    # channel, so the two cannot be detected apart by path; the version
    # floor is what separates them.
    $detectionPath = "{0}\Mozilla Firefox" -f $env:ProgramFiles

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : firefox.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName         = "Mozilla Firefox ESR (x64 en-US)"
        DisplayName     = "Mozilla Firefox ESR"
        Publisher       = "Mozilla"
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        RunningProcess  = @("firefox")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "firefox.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $fileVersion
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

function Invoke-PackageFirefoxEsr {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Mozilla Firefox ESR (x64) - PACKAGE phase"
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
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
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
        $v = Get-LatestFirefoxEsrVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Firefox ESR GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Mozilla Firefox ESR (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionsJsonUrl              : $VersionsJsonUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageFirefoxEsr
    }
    elseif ($PackageOnly) {
        Invoke-PackageFirefoxEsr
    }
    else {
        Invoke-StageFirefoxEsr
        Invoke-PackageFirefoxEsr
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-firefoxesr'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
