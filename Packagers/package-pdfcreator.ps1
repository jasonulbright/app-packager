<#
Vendor: pdfforge GmbH
App: PDFCreator
CMName: PDFCreator
VendorUrl: https://www.pdfforge.org/pdfcreator
CPE: cpe:2.3:a:pdfforge:pdfcreator:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.pdfforge.org/pdfcreator/changelog
DownloadPageUrl: https://www.pdfforge.org/pdfcreator/download
UpdateCadenceDays: 60

.SYNOPSIS
    Packages PDFCreator Free (x64) for MECM.

.DESCRIPTION
    Resolves the current stable build from the vendor's download redirect,
    downloads the setup executable, stages content to a versioned local folder,
    and creates an MECM Application with file-version detection.

    The free edition ships only as a setup executable; the vendor's MSI is
    restricted to the Professional and Terminal Server editions. The setup
    selects every optional component when /COMPONENTS is omitted, which pulls
    in the PDF Architect promotion; the install wrapper passes
    /COMPONENTS=none so neither PDF Architect nor Images2PDF is installed.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available PDFCreator version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$StableRedirectUrl = "https://download.pdfforge.org/download/pdfcreator/PDFCreator-stable?download"

$VendorFolder = "pdfforge"
$AppFolder    = "PDFCreator"

$BaseDownloadRoot = Join-Path $DownloadRoot "PDFCreator"
$InstallDir       = "C:\Program Files\PDFCreator"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        The download host answers 200 with an interstitial HTML page, which
        would otherwise stage as a valid-looking installer and fail only at
        install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestPdfCreatorRelease {
    <#
    .SYNOPSIS
        Returns the current stable PDFCreator version and its setup URL.
    .DESCRIPTION
        The vendor publishes no version index; the stable channel is a redirect
        that lands on a versioned CDN path. A one-byte range request resolves
        that path without pulling the ~150 MB setup, and the version comes from
        the resolved URL rather than the filename's underscore form.
    #>
    param([switch]$Quiet)

    Write-Log "PDFCreator stable redirect   : $StableRedirectUrl" -Quiet:$Quiet

    $probePath = Join-Path $BaseDownloadRoot "redirect-probe.bin"

    try {
        Initialize-Folder -Path $BaseDownloadRoot

        $resolved = (curl.exe -L --fail --silent --show-error -r '0-0' -o $probePath -w '%{url_effective}' $StableRedirectUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve the PDFCreator stable download redirect." }

        $m = [regex]::Match($resolved, '/pdfcreator/(?<ver>\d+\.\d+\.\d+)/(?<file>PDFCreator-[0-9_]+-Setup\.exe)$')
        if (-not $m.Success) {
            throw "Stable redirect did not land on a versioned setup URL: $resolved"
        }

        Write-Log "Latest PDFCreator version    : $($m.Groups['ver'].Value)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $m.Groups['ver'].Value
            FileName    = $m.Groups['file'].Value
            DownloadUrl = $resolved
        }
    }
    catch {
        Write-Log "Failed to get PDFCreator version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
    finally {
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePdfCreator {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFCreator (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestPdfCreatorRelease
    if (-not $releaseInfo) { throw "Could not resolve PDFCreator version." }

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
        Write-Log "Downloading PDFCreator..."
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
    # /COMPONENTS=none is load-bearing: omitting it selects every optional
    # component, which installs PDF Architect alongside PDFCreator.
    $installArgs = "'/VerySilent', '/COMPONENTS=none', '/NoIcons'"

    # The setup installs an embedded MSI whose ProductCode changes with every
    # release, so removal reads the code back from the ARP entry the install
    # wrote. An absent entry means the product is not present, so removal
    # exits clean.
    $uninstallScript = (
        '$roots = @(',
        '    ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'',',
        '    ''HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall''',
        ')',
        '$entry = Get-ChildItem -Path $roots -ErrorAction SilentlyContinue |',
        '    Where-Object { $_.PSChildName -match ''^\{[0-9A-Fa-f-]{36}\}$'' } |',
        '    Where-Object { (Get-ItemProperty -Path $_.PSPath -Name DisplayName -ErrorAction SilentlyContinue).DisplayName -eq ''PDFCreator'' } |',
        '    Select-Object -First 1',
        'if (-not $entry) { exit 0 }',
        'Get-Process PDFCreator -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', $entry.PSChildName, ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs $installArgs `
        -UninstallCommand 'msiexec.exe'

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallScript

    # --- Write stage manifest ---
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : PDFCreator.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "PDFCreator"
        Publisher       = "pdfforge GmbH"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VerySilent /COMPONENTS=none /NoIcons"
        RunningProcess  = @("PDFCreator")
        Detection       = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "PDFCreator.exe"
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

function Invoke-PackagePdfCreator {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFCreator (x64) - PACKAGE phase"
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
        $info = Get-LatestPdfCreatorRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("PDFCreator GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PDFCreator (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "StableRedirectUrl            : $StableRedirectUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePdfCreator
    }
    elseif ($PackageOnly) {
        Invoke-PackagePdfCreator
    }
    else {
        Invoke-StagePdfCreator
        Invoke-PackagePdfCreator
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-pdfcreator'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
