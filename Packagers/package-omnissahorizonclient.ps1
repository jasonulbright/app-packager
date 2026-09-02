<#
Vendor: Omnissa
App: Omnissa Horizon Client
CMName: Omnissa Horizon Client
VendorUrl: https://www.omnissa.com/
CPE: cpe:2.3:a:omnissa:horizon_client:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.omnissa.com/category/HorizonClientforWindowsReleaseNotes
DownloadPageUrl: https://customerconnect.omnissa.com/downloads/info/slug/virtual_desktop_and_apps/omnissa_horizon_clients/8
UpdateCadenceDays: 90

.SYNOPSIS
    Packages the Omnissa Horizon Client (x64, per-machine) for MECM.

.DESCRIPTION
    Resolves the current Windows client release from the public Customer
    Connect download API, downloads the vendor installer, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection on the installed client binary.

    The installer is a WiX burn bundle. Its bundle GUID changes with every
    release and the inner MSIs register as system components, so detection uses
    the installed horizon-client.exe rather than an ARP key.

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
    Content is staged under: <FileServerPath>\Applications\Omnissa\Omnissa Horizon Client\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Omnissa Horizon Client).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 20

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 45

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Horizon Client version string and exits.

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
    [int]$EstimatedRuntimeMins = 20,
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

# --- Configuration ---
$DlgListUrl    = "https://customerconnect.omnissa.com/channel/public/api/v1.0/products/getRelatedDLGList?locale=en_US&category=virtual_desktop_and_apps&product=omnissa_horizon_clients&version=8&dlgType=PRODUCT_BINARY"
$DlgDetailsUrl = "https://customerconnect.omnissa.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup={0}&productId={1}"
$WindowsDlgName = "Omnissa Horizon Client for Windows"
$ProductId      = "1616"

$VendorFolder = "Omnissa"
$AppFolder    = "Omnissa Horizon Client"

$BaseDownloadRoot  = Join-Path $DownloadRoot "Omnissa Horizon Client"
$InstallerFileName = "omnissa-horizon-client.exe"

# The x64 client MSI installs under the native Program Files tree.
$InstallDir = "$env:ProgramFiles\Omnissa\Omnissa Horizon Client"
$ClientExe  = "horizon-client.exe"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE/MZ signature.
    .DESCRIPTION
        The download host answers 200 with an HTML body for an expired or
        withdrawn build, which would otherwise stage as a valid-looking
        installer and fail only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2) { throw "Downloaded payload is too small to be an executable: $Path" }
    if ($bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not an executable (no MZ header): $Path"
    }
}


function Get-LatestHorizonClientRelease {
    <#
    .SYNOPSIS
        Returns the current Windows Horizon Client version and its download URL.
    .DESCRIPTION
        The public download API keys each release by a per-quarter download
        group code, so the group is resolved from the product listing before
        the file details are read. The details response carries the anonymous
        CDN URL that the download page hands out.
    #>
    param([switch]$Quiet)

    Write-Log "Omnissa download group API   : $DlgListUrl" -Quiet:$Quiet

    try {
        $listJson = (curl.exe -L --fail --silent --show-error $DlgListUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query the Omnissa download group list." }

        $list = ConvertFrom-Json $listJson
        $group = $list.dlgEditionsLists |
            Where-Object { $_.name -eq $WindowsDlgName } |
            ForEach-Object { $_.dlgList } |
            Where-Object { $_.name -eq $WindowsDlgName } |
            Select-Object -First 1

        if (-not $group -or [string]::IsNullOrWhiteSpace($group.code)) {
            throw "Could not locate the '$WindowsDlgName' download group."
        }

        Write-Log "Download group code          : $($group.code)" -Quiet:$Quiet

        $detailsUrl = $DlgDetailsUrl -f $group.code, $ProductId
        $detailsJson = (curl.exe -L --fail --silent --show-error $detailsUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query download details for $($group.code)." }

        $details = ConvertFrom-Json $detailsJson
        $file = $details.downloadFiles |
            Where-Object { $_.fileType -eq 'exe' -and $_.fileName -match '\.exe$' } |
            Select-Object -First 1

        if (-not $file) { throw "Download group $($group.code) carries no Windows exe." }
        if ([string]::IsNullOrWhiteSpace($file.thirdPartyDownloadUrl)) {
            throw "Download group $($group.code) exposes no anonymous download URL."
        }

        # fileName carries the marketing release and the product version:
        # Omnissa-Horizon-Client-2606-8.19.0-32215845441.exe
        $m = [regex]::Match($file.fileName, '-(?<ver>\d+\.\d+\.\d+)-\d+\.exe$')
        if (-not $m.Success) {
            throw "Could not parse a product version from file name: $($file.fileName)"
        }
        $version = $m.Groups['ver'].Value

        Write-Log "Latest Horizon Client        : $version" -Quiet:$Quiet
        Write-Log "Resolved download URL        : $($file.thirdPartyDownloadUrl)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $file.fileName
            DownloadUrl = $file.thirdPartyDownloadUrl
            Sha256      = $file.sha256checksum
        }
    }
    catch {
        Write-Log "Failed to get Horizon Client version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageHorizonClient {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Omnissa Horizon Client (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestHorizonClientRelease
    if (-not $releaseInfo) { throw "Could not resolve Omnissa Horizon Client version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Vendor file name             : $($releaseInfo.FileName)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
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

    # --- Verify the vendor checksum ---
    if (-not [string]::IsNullOrWhiteSpace($releaseInfo.Sha256)) {
        $actual = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $releaseInfo.Sha256.ToUpperInvariant()) {
            throw "Downloaded installer SHA256 ($actual) does not match the vendor value ($($releaseInfo.Sha256))."
        }
        Write-Log "SHA256 matches vendor value  : $actual"
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedExe"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrappers = New-ExeWrapperContent -InstallerFileName $InstallerFileName `
        -InstallArgs "'/silent', '/norestart'" `
        -UninstallCommand 'unused'

    # The bundle caches itself under ProgramData and publishes the removal
    # command line on its own ARP entry; the staged installer is not on the
    # client at uninstall time. A 32-bit burn bundle registers in the
    # WOW6432Node view, so both views are searched.
    $uninstallContent = @'
$roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
$entry = $null
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $entry = Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.DisplayName -eq 'Omnissa Horizon Client' } |
        Select-Object -First 1
    if ($entry) { break }
}
if (-not $entry) { exit 0 }

$command = $entry.QuietUninstallString
if ([string]::IsNullOrWhiteSpace($command)) { $command = $entry.UninstallString }
if ([string]::IsNullOrWhiteSpace($command)) { exit 1 }

if ($command -match '^\s*"(?<exe>[^"]+)"\s*(?<rest>.*)$') {
    $exe = $Matches['exe']
    $rest = $Matches['rest']
}
else {
    $exe = $command.Trim()
    $rest = ''
}

$argList = @('/uninstall', '/silent', '/norestart')
foreach ($token in ($rest -split '\s+')) {
    if ($token -and ($argList -notcontains $token)) { $argList += $token }
}

$proc = Start-Process -FilePath $exe -ArgumentList $argList -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    Write-Log ""
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : $ClientExe"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Omnissa Horizon Client"
        Publisher       = "Omnissa, LLC"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/silent /norestart"
        UninstallArgs   = "/uninstall /silent /norestart"
        RunningProcess  = @("horizon-client")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = $ClientExe
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageHorizonClient {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Omnissa Horizon Client (x64) - PACKAGE phase"
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
        $info = Get-LatestHorizonClientRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Omnissa Horizon Client GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Omnissa Horizon Client (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DlgListUrl                   : $DlgListUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageHorizonClient
    }
    elseif ($PackageOnly) {
        Invoke-PackageHorizonClient
    }
    else {
        Invoke-StageHorizonClient
        Invoke-PackageHorizonClient
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-omnissahorizonclient'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
