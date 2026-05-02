<#
Vendor: Adobe Inc.
App: Adobe Acrobat Reader
CMName: Adobe Acrobat Reader
VendorUrl: https://www.adobe.com/acrobat/pdf-reader.html
CPE: cpe:2.3:a:adobe:acrobat_reader_dc:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html
DownloadPageUrl: https://www.adobe.com/acrobat/pdf-reader.html
RequiresTools: 7-Zip

.SYNOPSIS
    Packages the latest Adobe Acrobat Reader for MECM.

.DESCRIPTION
    Parses Adobe's official release notes page to determine the current Acrobat
    version, constructs the enterprise installer URL, downloads the x86 en_US EXE
    from the enterprise distribution CDN, stages content to a versioned local folder
    with file-based detection metadata, and creates an MECM Application with file
    version-based detection.
    Detection uses AcroRd32.exe file version >= packaged version in the Program
    Files (x86) install path.

    NOTE: Adobe renamed this product in the 26.x release. We use the static name
    "Adobe Acrobat Reader" regardless of Adobe's branding changes.

    Adobe Acrobat version notation:
      Release notes use format NN.NNN.NNNNN (e.g., 25.001.21223)
      Download URL uses the same parts concatenated (e.g., 2500121223)

    Supports two-phase operation:
      -StageOnly    Download, read FileVersion, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Adobe\Acrobat Reader\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AdobeReader).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, read FileVersion, generate
    content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Parses Adobe's release notes page for the current version, outputs the version
    string, and exits. No download or MECM changes are made.

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
$AdobeReleaseNotesUrl = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"
$AdobeDownloadBase    = "https://ardownload3.adobe.com/pub/adobe/reader/win/AcrobatDC"

$VendorFolder = "Adobe"
$AppFolder    = "Acrobat Reader"

$BaseDownloadRoot = Join-Path $DownloadRoot "AdobeReader"

# --- Functions ---


function Get-AdobeAcrobatVersion {
    param([switch]$Quiet)

    Write-Log "Release notes URL            : $AdobeReleaseNotesUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }

        $verMatch = [regex]::Match($html, '\b(\d{2}\.\d{3}\.\d{5})\b')
        if (-not $verMatch.Success) { throw "Could not parse Acrobat DC version from release notes page." }

        $version = $verMatch.Groups[1].Value

        Write-Log "Latest Acrobat DC version    : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Acrobat DC version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function ConvertTo-AdobeUrlVersion {
    param([Parameter(Mandatory)][string]$Version)

    $parts      = $Version -split '\.'
    return "$($parts[0])$($parts[1])$($parts[2])"
}

function Get-AdobeInstallerInfo {
    param([Parameter(Mandatory)][string]$Version)

    $urlVersion = ConvertTo-AdobeUrlVersion -Version $Version
    $fileName   = "AcroRdrDC${urlVersion}_en_US.exe"
    $url        = "$AdobeDownloadBase/$urlVersion/$fileName"

    return [PSCustomObject]@{
        UrlVersion  = $urlVersion
        FileName    = $fileName
        DownloadUrl = $url
    }
}

function Get-AdobePatchInfo {
    param([Parameter(Mandatory)][string]$Version)

    $urlVersion = ConvertTo-AdobeUrlVersion -Version $Version
    $fileName   = "AcroRdrDCUpd${urlVersion}.msp"
    $url        = "$AdobeDownloadBase/$urlVersion/$fileName"

    return [PSCustomObject]@{
        UrlVersion  = $urlVersion
        FileName    = $fileName
        DownloadUrl = $url
    }
}

function Test-AdobeDownloadUrl {
    param([Parameter(Mandatory)][string]$Url)

    & curl.exe -I -L --fail --silent --show-error $Url 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Get-AdobeReleaseVersions {
    param([Parameter(Mandatory)][string]$Html)

    $seen = @{}
    $versions = New-Object System.Collections.Generic.List[string]

    foreach ($match in [regex]::Matches($Html, '\b\d{2}\.\d{3}\.\d{5}\b')) {
        $version = $match.Value
        if (-not $seen.ContainsKey($version)) {
            $seen[$version] = $true
            [void]$versions.Add($version)
        }
    }

    return $versions.ToArray()
}

function Resolve-AdobeInstallerPlan {
    param(
        [Parameter(Mandatory)][string]$Version,
        [switch]$Quiet
    )

    $fullInfo = Get-AdobeInstallerInfo -Version $Version
    if (Test-AdobeDownloadUrl -Url $fullInfo.DownloadUrl) {
        Write-Log "Full installer available     : $($fullInfo.FileName)" -Quiet:$Quiet
        return [PSCustomObject]@{
            Mode                 = 'FullExe'
            PackageVersion       = $Version
            FullInstallerVersion = $Version
            FullInstaller        = $fullInfo
            Patch                = $null
        }
    }

    Write-Log "Full installer unavailable   : $($fullInfo.DownloadUrl)" -Level WARN -Quiet:$Quiet

    $patchInfo = Get-AdobePatchInfo -Version $Version
    if (-not (Test-AdobeDownloadUrl -Url $patchInfo.DownloadUrl)) {
        throw "Neither full installer nor update MSP is available for Adobe Reader $Version."
    }
    Write-Log "Update MSP available         : $($patchInfo.FileName)" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }

    foreach ($candidateVersion in (Get-AdobeReleaseVersions -Html $html)) {
        if ($candidateVersion -eq $Version) { continue }

        $candidateInfo = Get-AdobeInstallerInfo -Version $candidateVersion
        if (Test-AdobeDownloadUrl -Url $candidateInfo.DownloadUrl) {
            Write-Log "Using base full installer    : $($candidateInfo.FileName)" -Quiet:$Quiet
            return [PSCustomObject]@{
                Mode                 = 'FullExePlusPatch'
                PackageVersion       = $Version
                FullInstallerVersion = $candidateVersion
                FullInstaller        = $candidateInfo
                Patch                = $patchInfo
            }
        }
    }

    throw "Could not find an available Adobe Reader full installer to pair with update $Version."
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAdobeReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Adobe Acrobat Reader - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-AdobeAcrobatVersion
    if (-not $version) { throw "Could not resolve Acrobat DC version." }

    $installPlan       = Resolve-AdobeInstallerPlan -Version $version
    $dlInfo            = $installPlan.FullInstaller
    $installerFileName = $dlInfo.FileName
    $downloadUrl       = $dlInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Package URL version          : $(ConvertTo-AdobeUrlVersion -Version $version)"
    if ($installPlan.Mode -eq 'FullExePlusPatch') {
        Write-Log "Full installer version       : $($installPlan.FullInstallerVersion)"
        Write-Log "Full installer URL version   : $($dlInfo.UrlVersion)"
        Write-Log "Full installer filename      : $installerFileName"
        Write-Log "Update MSP filename          : $($installPlan.Patch.FileName)"
    }
    else {
        Write-Log "Installer filename           : $installerFileName"
    }
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
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

    # --- Extract EXE to get setup.exe + MSI + setup.ini ---
    # Adobe enterprise EXE is a self-extracting 7z archive containing:
    #   setup.exe (bootstrapper), setup.ini (config), AcroRead.msi, abcpy.ini,
    #   Data1.cab, and an .msp patch file.
    # We ship the entire extracted contents; setup.exe handles prereqs via setup.ini.
    # Use 7-Zip for fast, silent extraction (the EXE's own -sfx switches pop a GUI).
    # APP_PACKAGER_SEVENZIP is set by start-apppackager.ps1 when a non-default
    # 7-Zip install was detected via the pre-flight scan. Fall back to the
    # Program Files default for CLI / un-hosted invocations.
    Write-Log "Extracting installer package..."
    $sevenZip = $env:APP_PACKAGER_SEVENZIP
    if ([string]::IsNullOrWhiteSpace($sevenZip) -or -not (Test-Path -LiteralPath $sevenZip)) {
        $sevenZip = Join-Path $env:ProgramFiles "7-Zip\7z.exe"
    }
    if (-not (Test-Path -LiteralPath $sevenZip)) {
        throw "7-Zip not found at $sevenZip - required to extract Adobe enterprise installer. Install 7-Zip or verify detection in MECM Preferences."
    }
    $extractProc = $null
    try {
        $extractProc = Start-Process -FilePath $sevenZip -ArgumentList @('x', "-o$localContentPath", '-y', $localExe) -Wait -PassThru -NoNewWindow
        if ($extractProc.ExitCode -ne 0) {
            throw "7-Zip extraction failed with exit code $($extractProc.ExitCode)"
        }
    }
    finally {
        if ($extractProc) { try { $extractProc.Dispose() } catch { } }
    }

    if ($installPlan.Mode -eq 'FullExePlusPatch') {
        $patchFileName = $installPlan.Patch.FileName
        $localMsp = Join-Path $BaseDownloadRoot $patchFileName

        if (-not (Test-Path -LiteralPath $localMsp)) {
            Write-Log "Update MSP URL               : $($installPlan.Patch.DownloadUrl)"
            Write-Log "Downloading update MSP..."
            Invoke-DownloadWithRetry -Url $($installPlan.Patch.DownloadUrl) -OutFile $localMsp
        }
        else {
            Write-Log "Local update MSP exists. Skipping download."
        }

        Get-ChildItem -Path $localContentPath -Filter "*.msp" -File -ErrorAction SilentlyContinue |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop }

        Copy-Item -LiteralPath $localMsp -Destination (Join-Path $localContentPath $patchFileName) -Force -ErrorAction Stop

        $setupIniPath = Join-Path $localContentPath "setup.ini"
        $setupIni = @(
            "[Startup]",
            "RequireMSI=3.0",
            "",
            "[Product]",
            "PATCH=$patchFileName",
            "msi=AcroRead.msi"
        ) -join "`r`n"
        Set-Content -LiteralPath $setupIniPath -Value $setupIni -Encoding ASCII -ErrorAction Stop
        Write-Log "Updated setup.ini patch      : $patchFileName"
    }

    $extractedFiles = Get-ChildItem -Path $localContentPath -File
    Write-Log "Extracted files              : $($extractedFiles.Count)"
    foreach ($f in $extractedFiles) { Write-Log "  $($f.Name)" }

    # --- Verify setup.exe and MSI exist ---
    $setupExe = Join-Path $localContentPath "setup.exe"
    if (-not (Test-Path -LiteralPath $setupExe)) {
        throw "setup.exe not found in extracted content"
    }

    $msiFile = Get-ChildItem -Path $localContentPath -Filter "*.msi" | Select-Object -First 1
    if (-not $msiFile) {
        throw "No MSI found in extracted content"
    }

    $msiFileName = $msiFile.Name
    Write-Log "MSI file                     : $msiFileName"

    # --- Read MSI properties for detection and uninstall ---
    $msiProps = Get-MsiPropertyMap -MsiPath $msiFile.FullName
    $productCode = $msiProps.ProductCode
    $msiVersion  = $msiProps.ProductVersion
    if (-not $productCode) {
        throw "Could not read ProductCode from $msiFileName"
    }

    Write-Log "MSI ProductCode              : $productCode"
    Write-Log "MSI ProductVersion           : $msiVersion"

    # Base MSI version (15.x) is outdated; the .msp patch brings it to the release version.
    # Always use the release notes version for detection.
    $detectionVersion = $version

    # --- Generate content wrappers ---
    # Install via setup.exe bootstrapper (reads setup.ini, handles prereqs)
    $installContent = (
        '$setupPath = Join-Path $PSScriptRoot ''setup.exe''',
        '$proc = Start-Process -FilePath $setupPath -ArgumentList @(''/sAll'', ''/rs'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall via msiexec with hardcoded ProductCode
    $uninstallContent = (
        ('$productCode = ''{0}''' -f $productCode),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', $productCode, ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $detectionPath = "{0}\Adobe\Acrobat Reader DC\Reader" -f ${env:ProgramFiles(x86)}

    $appName   = "Adobe Acrobat Reader $version"
    $publisher = "Adobe Inc."

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : AcroRd32.exe"
    Write-Log "Detection version            : $detectionVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = "setup.exe"
        InstallerType   = "EXE"
        InstallArgs     = "/sAll /rs"
        UninstallArgs   = "/x $productCode /qn /norestart"
        RunningProcess  = @("AcroRd32")
        ProductCode     = $productCode
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "AcroRd32.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $detectionVersion
            Is64Bit       = $false
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

function Invoke-PackageAdobeReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Adobe Acrobat Reader - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log "Detection Version            : $($manifest.Detection.ExpectedValue)"
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
        $v = Get-AdobeAcrobatVersion -Quiet
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
    Write-Log "Adobe Acrobat Reader Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "AdobeReleaseNotesUrl         : $AdobeReleaseNotesUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAdobeReader
    }
    elseif ($PackageOnly) {
        Invoke-PackageAdobeReader
    }
    else {
        Invoke-StageAdobeReader
        Invoke-PackageAdobeReader
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
