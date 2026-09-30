<#
Vendor: Adobe Inc.
App: Adobe Acrobat Reader
CMName: Adobe Acrobat Reader
VendorUrl: https://www.adobe.com/acrobat/pdf-reader.html
CPE: cpe:2.3:a:adobe:acrobat_reader_dc:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html
DownloadPageUrl: https://www.adobe.com/acrobat/pdf-reader.html
IconSource: Installer
RequiresTools: 7-Zip

.SYNOPSIS
    Packages the latest Adobe Acrobat Reader for ConfigMgr.

.DESCRIPTION
    Parses Adobe's official release notes page to determine the current Acrobat
    version, constructs the enterprise installer URL, downloads the x86 English or
    multilingual (MUI) EXE from the enterprise distribution CDN, stages content to a versioned local folder
    with file-based detection metadata, and creates a ConfigMgr Application with file
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
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER Edition
    English installs the en_US installer. MUI installs the multilingual installer with
    the languages in -Languages (English is always included) and lets the UI follow
    the language of the operating system. Defaults to Packagers\packager-preferences.json
    under AdobeReaderInstallOptions, which the Packager Preferences UI writes; English
    when nothing is stored.

.PARAMETER Languages
    MUI only. Comma-separated Adobe locale codes (de_DE,fr_FR) or All. Defaults to the
    stored preference.

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
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, read FileVersion, generate
    content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Parses Adobe's release notes page for the current version, outputs the version
    string, and exits. No download or ConfigMgr changes are made.

.PARAMETER VerboseLog
    Enables DEBUG-level diagnostic logging (CM module/drive state, manifest
    fields, deployment type parameters, provider site-code validation).
    Equivalent to setting APP_PACKAGER_VERBOSE=1. On failure the log always
    includes the exception chain, failing file:line, and script stack trace.

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
    [ValidateSet('', 'English', 'MUI')][string]$Edition = '',
    [string]$Languages = '',
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
$AdobeReleaseNotesUrl = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"
$AdobeDownloadBase    = "https://ardownload3.adobe.com/pub/adobe/reader/win/AcrobatDC"

$VendorFolder = "Adobe"
$AppFolder    = "Acrobat Reader"

$BaseDownloadRoot = Join-Path $DownloadRoot "AdobeReader"

# The languages the MUI installer lists in its setup.ini (27 LCIDs). English
# is always installed.
$AdobeMuiLocales = @('ca_ES','cs_CZ','da_DK','de_DE','en_US','es_ES','eu_ES','fi_FI','fr_FR','hr_HR','hu_HU','it_IT','ja_JP','ko_KR','nb_NO','nl_NL','pl_PL','pt_BR','ro_RO','ru_RU','sk_SK','sl_SI','sv_SE','tr_TR','uk_UA','zh_CN','zh_TW')

# --- Functions ---

function Get-AdobeReaderInstallOptions {
    # Explicit parameters win over the stored preference; English is the
    # default when nothing is stored.
    param([string]$Edition = '', [string]$Languages = '')
    $options = [pscustomobject]@{ Edition = 'English'; Languages = @() }
    $prefs = Get-PackagerPreferences
    if ($prefs -and $prefs.PSObject.Properties['AdobeReaderInstallOptions'] -and $prefs.AdobeReaderInstallOptions) {
        $cfg = $prefs.AdobeReaderInstallOptions
        if ([string]$cfg.Edition -in @('English', 'MUI')) { $options.Edition = [string]$cfg.Edition }
        if ($cfg.PSObject.Properties['Languages'] -and $null -ne $cfg.Languages) { $options.Languages = @($cfg.Languages | ForEach-Object { [string]$_ } | Where-Object { $_ }) }
    }
    if ($Edition) { $options.Edition = $Edition }
    if ($Languages) { $options.Languages = @($Languages -split '[,;\s]+' | Where-Object { $_ }) }
    if ($options.Edition -eq 'MUI') {
        $options.Languages = @(ConvertTo-AdobeLanguageList -Languages $options.Languages)
    }
    else { $options.Languages = @() }
    return $options
}

function ConvertTo-AdobeLanguageList {
    # Normalizes a language selection to the LANG_LIST value: All, or the
    # sorted locale codes with en_US present. An unknown code stops the run.
    param([string[]]$Languages)
    $list = @($Languages | ForEach-Object { [string]$_ } | Where-Object { $_ })
    if ($list -contains 'All' -or $list -contains 'all') { return @('All') }
    $unknown = @($list | Where-Object { $_ -notin $AdobeMuiLocales })
    if ($unknown.Count) { throw ("Unknown Adobe locale code(s): {0}. Known codes: {1}." -f ($unknown -join ', '), ($AdobeMuiLocales -join ', ')) }
    $set = @($list + 'en_US' | Sort-Object -Unique)
    return $set
}

function Get-AdobeInstallerSuffix {
    param([Parameter(Mandatory)][string]$Edition)
    if ($Edition -eq 'MUI') { return '_MUI' }
    return '_en_US'
}

function Set-AdobeSetupIniCommandLine {
    # Writes the MSI properties setup.exe hands to msiexec: the CmdLine key
    # of the [Product] section. An empty value removes the key, so a stage
    # folder reused for the English edition carries no MUI properties. Other
    # keys of the file stay as extracted.
    param([Parameter(Mandatory)][string]$Path, [AllowEmptyString()][string]$CommandLine = '')
    $lines = @()
    if (Test-Path -LiteralPath $Path) { $lines = @(Get-Content -LiteralPath $Path -ErrorAction Stop) }
    $out = New-Object System.Collections.Generic.List[string]
    $inProduct = $false
    $written = $false
    if (-not $CommandLine) { $written = $true }
    foreach ($line in $lines) {
        if ($line -match '^\s*\[(.+)\]\s*$') {
            if ($inProduct -and -not $written) { $out.Add('CmdLine=' + $CommandLine); $written = $true }
            $inProduct = ($Matches[1] -eq 'Product')
            $out.Add($line)
            continue
        }
        if ($inProduct -and $line -match '^\s*CmdLine\s*=') {
            if (-not $written) { $out.Add('CmdLine=' + $CommandLine); $written = $true }
            continue
        }
        $out.Add($line)
    }
    if (-not $written) {
        if (-not $inProduct) { $out.Add('[Product]') }
        $out.Add('CmdLine=' + $CommandLine)
    }
    Set-Content -LiteralPath $Path -Value ($out -join "`r`n") -Encoding ASCII -ErrorAction Stop
}


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
    param([Parameter(Mandatory)][string]$Version, [string]$Edition = 'English')

    $urlVersion = ConvertTo-AdobeUrlVersion -Version $Version
    $fileName   = "AcroRdrDC${urlVersion}$(Get-AdobeInstallerSuffix -Edition $Edition).exe"
    $url        = "$AdobeDownloadBase/$urlVersion/$fileName"

    return [PSCustomObject]@{
        UrlVersion  = $urlVersion
        FileName    = $fileName
        DownloadUrl = $url
    }
}

function Get-AdobePatchInfo {
    # A MUI base takes the MUI patch; the English patch does not apply to it.
    param([Parameter(Mandatory)][string]$Version, [string]$Edition = 'English')

    $urlVersion = ConvertTo-AdobeUrlVersion -Version $Version
    $fileName   = "AcroRdrDCUpd${urlVersion}$(if ($Edition -eq 'MUI') { '_MUI' } else { '' }).msp"
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
        [string]$Edition = 'English',
        [switch]$Quiet
    )

    $fullInfo = Get-AdobeInstallerInfo -Version $Version -Edition $Edition
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

    $patchInfo = Get-AdobePatchInfo -Version $Version -Edition $Edition
    if (-not (Test-AdobeDownloadUrl -Url $patchInfo.DownloadUrl)) {
        throw "Neither full installer nor update MSP is available for Adobe Reader $Version."
    }
    Write-Log "Update MSP available         : $($patchInfo.FileName)" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }

    foreach ($candidateVersion in (Get-AdobeReleaseVersions -Html $html)) {
        if ($candidateVersion -eq $Version) { continue }

        $candidateInfo = Get-AdobeInstallerInfo -Version $candidateVersion -Edition $Edition
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

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-AdobeAcrobatVersion
    if (-not $version) { throw "Could not resolve Acrobat DC version." }

    $installOptions    = Get-AdobeReaderInstallOptions -Edition $Edition -Languages $Languages
    $installPlan       = Resolve-AdobeInstallerPlan -Version $version -Edition $installOptions.Edition
    $dlInfo            = $installPlan.FullInstaller
    $installerFileName = $dlInfo.FileName
    $downloadUrl       = $dlInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Edition                      : $($installOptions.Edition)"
    if ($installOptions.Edition -eq 'MUI') { Write-Log "Languages                    : $($installOptions.Languages -join ',')" }
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
        throw "7-Zip not found at $sevenZip - required to extract Adobe enterprise installer. Install 7-Zip or verify detection in ConfigMgr Preferences."
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

    # MUI: the languages to install and the UI language following the OS.
    # setup.exe hands the CmdLine value of setup.ini to msiexec.
    $muiCommandLine = ''
    if ($installOptions.Edition -eq 'MUI') {
        $muiCommandLine = ('LANG_LIST="{0}" SUPPRESSLANGSELECTION=YES' -f ($installOptions.Languages -join ','))
        Write-Log "setup.ini CmdLine            : $muiCommandLine"
    }
    Set-AdobeSetupIniCommandLine -Path (Join-Path $localContentPath "setup.ini") -CommandLine $muiCommandLine

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
        Edition         = $installOptions.Edition
        Languages       = @($installOptions.Languages)
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
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-adobereader'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
