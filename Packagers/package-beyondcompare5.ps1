<#
Vendor: Scooter Software
App: Beyond Compare 5
CMName: Beyond Compare 5
VendorUrl: https://www.scootersoftware.com/
CPE: cpe:2.3:a:scootersoftware:beyond_compare:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.scootersoftware.com/kb/dl5_winalternate
DownloadPageUrl: https://www.scootersoftware.com/kb/dl5_winalternate
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Beyond Compare 5 (x64) for ConfigMgr.

.DESCRIPTION
    Resolves the newest build from the vendor's alternate Windows download
    page, downloads the setup zip, extracts the Inno Setup installer, and
    stages it with the organization's BC5Key.txt beside the installer, where
    setup reads it to register the product.

    Beyond Compare 5 requires a license key file. Stage and Package both stop
    when no key file is available. The key file is placed as-is; its contents
    are not read or validated.

    The installer runs per machine with /ALLUSERS. The uninstall key is taken
    from the installer header at Stage time, and uninstall resolves unins000.exe
    from that ARP entry.

    Supports two-phase operation:
      -StageOnly    Download, extract, place the key file, generate wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Scooter Software\Beyond Compare 5\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\BeyondCompare5).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER KeyFile
    Path to the BC5Key.txt license key file. Overrides the key file chosen in
    Options > Packager Preferences.

.PARAMETER StageOnly
    Runs only the Stage phase: download and extract the installer, place the
    key file, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Beyond Compare 5 version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - A Beyond Compare 5 license key file (BC5Key.txt)
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
    [string]$KeyFile = "",
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
$DownloadPageUrl = "https://www.scootersoftware.com/kb/dl5_winalternate"
$DownloadHost    = "https://www.scootersoftware.com"

$VendorFolder = "Scooter Software"
$AppFolder    = "Beyond Compare 5"

$BaseDownloadRoot = Join-Path $DownloadRoot "BeyondCompare5"

# Setup looks for this exact file name in its own folder.
$KeyFileName = "BC5Key.txt"

$InstallSwitches   = @('/VERYSILENT', '/NORESTART', '/ALLUSERS', '/DISABLEUPDATES', '/SUPPRESSMSGBOXES')
$UninstallSwitches = @('/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES')

# --- Functions ---

function Get-BeyondCompare5ReleaseFromPage {
    <#
    .SYNOPSIS
        Extracts the version and zip path from the alternate download page HTML.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html)

    $m = [regex]::Match($Html, '(?<path>/files/(?<file>BCompareSetup-(?<ver>\d+(?:\.\d+){3})\.zip))')
    if (-not $m.Success) { return $null }

    return [pscustomobject]@{
        Version     = $m.Groups['ver'].Value
        FileName    = $m.Groups['file'].Value
        DownloadUrl = $DownloadHost + $m.Groups['path'].Value
    }
}


function Get-LatestBeyondCompare5Release {
    <#
    .SYNOPSIS
        Returns the newest Beyond Compare 5 version and its setup zip URL.
    #>
    param([switch]$Quiet)

    Write-Log "BC5 download page            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error -A "Mozilla/5.0" $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Beyond Compare 5 download page: $DownloadPageUrl" }

        $release = Get-BeyondCompare5ReleaseFromPage -Html $html
        if (-not $release) { throw "No BCompareSetup-<version>.zip link found on $DownloadPageUrl" }

        Write-Log "Latest BC5 version           : $($release.Version)" -Quiet:$Quiet
        return $release
    }
    catch {
        Write-Log "Failed to get Beyond Compare 5 version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Resolve-BeyondCompare5KeyFile {
    <#
    .SYNOPSIS
        Returns the license key file path from -KeyFile or packager preferences.
    .DESCRIPTION
        Throws when no key file is configured or the file is missing. The file
        contents are never read.
    #>
    param([string]$Override)

    $path = $Override
    if ([string]::IsNullOrWhiteSpace($path)) {
        $prefs = Get-PackagerPreferences
        if ($prefs -and $prefs.PSObject.Properties['BeyondCompareKeyFile']) {
            $path = [string]$prefs.BeyondCompareKeyFile
        }
    }

    if ([string]::IsNullOrWhiteSpace($path)) {
        throw "Beyond Compare 5 requires a license key file. Choose BC5Key.txt in Options > Packager Preferences > Beyond Compare 5."
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Beyond Compare 5 license key file not found: $path"
    }
    return $path
}


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the extracted file starts with the PE 'MZ' signature.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Extracted installer is not a Windows executable (no MZ header): $Path"
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageBeyondCompare5 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Beyond Compare 5 (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- License key file (checked before any download) ---
    $keySource = Resolve-BeyondCompare5KeyFile -Override $KeyFile
    Write-Log "License key file             : $keySource"

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestBeyondCompare5Release
    if (-not $releaseInfo) { throw "Could not resolve Beyond Compare 5 version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log ""

    # --- Download ZIP ---
    $localZip = Join-Path $BaseDownloadRoot $releaseInfo.FileName
    Write-Log "Local ZIP path               : $localZip"

    if (-not (Test-Path -LiteralPath $localZip)) {
        Write-Log "Downloading Beyond Compare 5 ZIP..."
        Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localZip
    }
    else {
        Write-Log "Local ZIP exists. Skipping download."
    }

    # --- Extract installer from ZIP ---
    $extractDir = Join-Path $BaseDownloadRoot "_extracted"
    if (Test-Path -LiteralPath $extractDir) {
        Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive -LiteralPath $localZip -DestinationPath $extractDir -Force -ErrorAction Stop

    $exeFiles = @(Get-ChildItem -LiteralPath $extractDir -Filter "*.exe" -Recurse -File)
    if ($exeFiles.Count -ne 1) {
        throw "Expected one installer EXE inside $($releaseInfo.FileName), found $($exeFiles.Count)."
    }
    $extractedExe = $exeFiles[0]
    Assert-PayloadIsExecutable -Path $extractedExe.FullName
    $installerFileName = $extractedExe.Name

    Write-Log "Extracted installer          : $($extractedExe.FullName)"

    # --- Uninstall key from the installer header, all-users branch ---
    $analysis = Get-InstallerAnalysis -Path $extractedExe.FullName
    if (@($analysis.InstallModes) -contains 'AllUsers') {
        $analysis = Set-InstallerAnalysisMode -Analysis $analysis -Mode AllUsers
    }
    $decodedKey = [string]$analysis.UninstallRegistryKey
    if ([string]::IsNullOrWhiteSpace($decodedKey)) {
        throw "The Beyond Compare 5 installer header names no uninstall registry key; detection cannot be built."
    }
    $arpKey = ($decodedKey -replace '^HK(LM|CU):\\', '') -replace '(?i)^SOFTWARE\\WOW6432Node\\', 'SOFTWARE\'
    $is64Bit = ([string]$analysis.RegistryView -ne '32')

    Write-Log "ARP key                      : $arpKey ($(if ($is64Bit) { '64' } else { '32' })-bit view)"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    Copy-Item -LiteralPath $extractedExe.FullName -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $stagedExe"

    $stagedKey = Join-Path $localContentPath $KeyFileName
    Copy-Item -LiteralPath $keySource -Destination $stagedKey -Force -ErrorAction Stop
    Write-Log "Placed key beside installer  : $stagedKey"

    # --- Generate content wrappers ---
    $installArgList = ($InstallSwitches | ForEach-Object { "'$_'" }) -join ', '
    $installScript = @"
`$exePath = Join-Path `$PSScriptRoot '$installerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList @($installArgList) -WorkingDirectory `$PSScriptRoot -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    # The install folder is not known at stage time, so uninstall resolves
    # unins000.exe from the ARP entry the installer writes.
    $uninstallArgList = ($UninstallSwitches | ForEach-Object { "'$_'" }) -join ', '
    $registryView = if ($is64Bit) { 'Registry64' } else { 'Registry32' }
    $uninstallScript = @"
`$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::$registryView)
`$entry = `$base.OpenSubKey('$arpKey')
if (-not `$entry) {
    Write-Error 'Beyond Compare 5 uninstall entry not found.'
    exit 1
}
`$uninstallString = [string]`$entry.GetValue('UninstallString')
`$entry.Close()
`$exe = (`$uninstallString -replace '^"([^"]+)".*`$', '`$1').Trim('"')
if (-not `$exe -or -not (Test-Path -LiteralPath `$exe)) {
    Write-Error "Beyond Compare 5 uninstaller not found: `$uninstallString"
    exit 1
}
`$proc = Start-Process -FilePath `$exe -ArgumentList @($uninstallArgList) -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    $expectedVersion = $version
    $declared = [string]$analysis.DisplayVersion
    $parsed = $null
    if ($declared -and [version]::TryParse($declared, [ref]$parsed)) { $expectedVersion = $declared }

    Write-Log ""
    Write-Log "Detection                    : $arpKey DisplayVersion >= $expectedVersion"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Beyond Compare 5"
        Publisher       = "Scooter Software"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = ($InstallSwitches -join ' ')
        UninstallArgs   = ($UninstallSwitches -join ' ')
        RunningProcess  = @("BCompare")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $expectedVersion
            Is64Bit             = $is64Bit
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

function Invoke-PackageBeyondCompare5 {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Beyond Compare 5 (x64) - PACKAGE phase"
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

    if (-not (Test-Path -LiteralPath (Join-Path $localContentPath $KeyFileName) -PathType Leaf)) {
        throw "Staged content has no $KeyFileName. Choose the license key file in Options > Packager Preferences > Beyond Compare 5 and stage again."
    }

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
        $info = Get-LatestBeyondCompare5Release -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Beyond Compare 5 GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Beyond Compare 5 (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageBeyondCompare5
    }
    elseif ($PackageOnly) {
        Invoke-PackageBeyondCompare5
    }
    else {
        Invoke-StageBeyondCompare5
        Invoke-PackageBeyondCompare5
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-beyondcompare5'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
