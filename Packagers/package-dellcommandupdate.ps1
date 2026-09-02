<#
Vendor: Dell Inc.
App: Dell Command Update
CMName: Dell Command Update
VendorUrl: https://www.dell.com/
CPE: cpe:2.3:a:dell:command_update:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update
DownloadPageUrl: https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update
UpdateCadenceDays: 120

.SYNOPSIS
    Packages the Dell Command | Update Windows Universal Application (x64) for MECM.

.DESCRIPTION
    Resolves the current release from the Dell client catalog
    (downloads.dell.com/catalog/CatalogPC.cab) rather than the support page:
    the FOLDER path segment of the download URL changes with every release and
    the support page is not machine-readable.

    Stages content to a versioned local folder and creates an MECM Application
    with file-version detection on dcu-cli.exe.

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
    Outputs only the latest available Dell Command | Update version string and exits.

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force -ErrorAction Stop
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$CatalogUrl      = "https://downloads.dell.com/catalog/CatalogPC.cab"
$CatalogBaseUrl  = "https://downloads.dell.com"
$InstallDir      = "C:\Program Files\Dell\CommandUpdate"

$VendorFolder = "Dell"
$AppFolder    = "Dell Command Update"

$BaseDownloadRoot = Join-Path $DownloadRoot "Dell Command Update"

# --- Functions ---


function Assert-ExePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the MZ signature.
    .DESCRIPTION
        downloads.dell.com answers 200 with an HTML body for withdrawn FOLDER
        paths, which would otherwise stage as a valid-looking installer and fail
        only at install time.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Expand-DellCatalog {
    <#
    .SYNOPSIS
        Downloads the Dell client catalog and returns the path to the extracted XML.
    .DESCRIPTION
        The published cab nests a second cab that carries the UTF-16 XML, so two
        expand passes are required. Each pass extracts into its own directory
        because expand.exe writes a file, not a directory, when the destination
        does not exist.
    #>
    param([Parameter(Mandatory)][string]$WorkRoot)

    $cabPath = Join-Path $WorkRoot "CatalogPC.cab"
    Invoke-DownloadWithRetry -Url $CatalogUrl -OutFile $cabPath -Quiet

    $expand = Join-Path $env:SystemRoot "System32\expand.exe"
    $stage1 = Join-Path $WorkRoot "catalog-stage1"
    $stage2 = Join-Path $WorkRoot "catalog-stage2"
    foreach ($d in @($stage1, $stage2)) {
        if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction Stop }
        New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
    }

    & $expand $cabPath "-F:*" $stage1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Failed to expand Dell catalog cab: $cabPath" }

    $inner = Get-ChildItem -LiteralPath $stage1 -File | Sort-Object Length -Descending | Select-Object -First 1
    if (-not $inner) { throw "Dell catalog cab expanded to nothing." }

    # The outer cab may already yield the XML; only expand again when the
    # extracted file is itself a cab (MSCF signature).
    $sig = Get-Content -LiteralPath $inner.FullName -Encoding Byte -TotalCount 4 -ErrorAction Stop
    if ($sig.Count -ge 4 -and $sig[0] -eq 0x4D -and $sig[1] -eq 0x53 -and $sig[2] -eq 0x43 -and $sig[3] -eq 0x46) {
        & $expand $inner.FullName "-F:*" $stage2 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Failed to expand inner Dell catalog cab." }
        $inner = Get-ChildItem -LiteralPath $stage2 -File | Sort-Object Length -Descending | Select-Object -First 1
        if (-not $inner) { throw "Inner Dell catalog cab expanded to nothing." }
    }

    return $inner.FullName
}


function Get-LatestDellCommandUpdateRelease {
    <#
    .SYNOPSIS
        Returns the newest Dell Command | Update Windows Universal Application
        (WIN64) release and its download URL.
    .DESCRIPTION
        Dell publishes several components at the same version that differ only in
        which models they list, so ties on vendorVersion are broken by release
        date and then by supported-model count: the broadest package is the one
        that installs across a mixed fleet.
    #>
    param([switch]$Quiet)

    Write-Log "Dell client catalog          : $CatalogUrl" -Quiet:$Quiet

    $workRoot = Join-Path $BaseDownloadRoot "catalog"
    Initialize-Folder -Path $workRoot

    try {
        $xmlPath = Expand-DellCatalog -WorkRoot $workRoot
        $text = Get-Content -LiteralPath $xmlPath -Raw -Encoding Unicode -ErrorAction Stop

        $rx = [regex]'(?s)<SoftwareComponent[^>]*path="[^"]*Dell-Command-Update-Windows-Universal-Application[^"]*WIN64[^"]*".*?</SoftwareComponent>'
        $candidates = foreach ($m in $rx.Matches($text)) {
            $node = ([xml]$m.Value).SoftwareComponent
            [pscustomobject]@{
                Version     = [string]$node.vendorVersion
                Path        = [string]$node.path
                ReleaseDate = [datetime]::Parse([string]$node.releaseDate, [Globalization.CultureInfo]::InvariantCulture)
                ModelCount  = @($node.SupportedSystems.Brand.Model).Count
            }
        }

        if (-not $candidates) { throw "No Dell Command | Update WIN64 universal component found in the catalog." }

        $best = $candidates |
            Sort-Object @{ Expression = { [version]$_.Version } }, ReleaseDate, ModelCount -Descending |
            Select-Object -First 1

        Write-Log "Latest DCU version           : $($best.Version)" -Quiet:$Quiet
        Write-Log "Catalog path                 : $($best.Path)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $best.Version
            FileName    = Split-Path -Leaf $best.Path
            DownloadUrl = "$CatalogBaseUrl/$($best.Path)"
        }
    }
    catch {
        Write-Log "Failed to get Dell Command Update version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDellCommandUpdate {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Dell Command Update (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestDellCommandUpdateRelease
    if (-not $releaseInfo) { throw "Could not resolve Dell Command Update version." }

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
        Write-Log "Downloading Dell Command Update..."
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
    # The Dell installer is its own uninstaller: /s installs silently, /x removes.
    # Uninstall therefore runs the copy in the content folder, not an installed path.
    $installPs1 = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/s'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/x'', ''/s'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    Write-Log "Detection path               : $InstallDir"
    Write-Log "Detection file               : dcu-cli.exe >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Dell Command Update"
        Publisher        = "Dell Inc."
        SoftwareVersion  = $version
        InstallerFile    = $installerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/s"
        UninstallCommand = ".\$installerFileName"
        UninstallArgs    = "/x /s"
        RunningProcess   = @("DellCommandUpdate", "dcu-cli")
        Detection        = @{
            Type          = "File"
            FilePath      = $InstallDir
            FileName      = "dcu-cli.exe"
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

function Invoke-PackageDellCommandUpdate {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Dell Command Update (x64) - PACKAGE phase"
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
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

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
        $info = Get-LatestDellCommandUpdateRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Dell Command Update GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Dell Command Update (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CatalogUrl                   : $CatalogUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageDellCommandUpdate
    }
    elseif ($PackageOnly) {
        Invoke-PackageDellCommandUpdate
    }
    else {
        Invoke-StageDellCommandUpdate
        Invoke-PackageDellCommandUpdate
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-dellcommandupdate'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
