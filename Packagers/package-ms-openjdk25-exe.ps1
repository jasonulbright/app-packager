<#
Vendor: Microsoft
App: Microsoft Build of OpenJDK 25 (x64, EXE)
CMName: Microsoft OpenJDK 25 (EXE)
VendorUrl: https://learn.microsoft.com/en-us/java/openjdk/
CPE: cpe:2.3:a:microsoft:build_of_openjdk:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/java/openjdk/release-notes
DownloadPageUrl: https://learn.microsoft.com/en-us/java/openjdk/download
IconSource: Installer

.SYNOPSIS
    Packages Microsoft Build of OpenJDK 25 (x64) EXE, per machine, for ConfigMgr.

.DESCRIPTION
    Resolves the latest Microsoft Build of OpenJDK 25 release from the
    major-version download alias, downloads the x64 Inno Setup EXE, verifies
    it against the published .sha256sum.txt, stages content to a versioned
    local folder, and creates a ConfigMgr Application with HKLM ARP
    DisplayVersion detection on the Inno Setup ARP entry.

    The install runs /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /ALLUSERS
    /TASKS="FeatureEnvironment,FeatureJarFileRunWith" into the default
    per-machine folder.

    Install only one installation method per major version. MSI and EXE
    installs of Microsoft Build of OpenJDK 25 must not exist on the same
    device. Uninstall the other method before you deploy this application.

    Supports two-phase operation:
      -StageOnly    Download, verify SHA-256, derive ARP detection, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder.

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs the latest available version string and exits.

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
$FeatureVersion   = 25
$InstallerType    = "EXE"
$DownloadBaseUrl  = "https://aka.ms/download-jdk/"
# Only the MSI has a major-only alias; the EXE alias answers with a redirect
# to a search page. Both installers ship under the same full version.
$LatestAliasUrl   = "https://aka.ms/download-jdk/microsoft-jdk-$FeatureVersion-windows-x64.msi"

$VendorFolder     = "Microsoft"
$AppFolder        = "OpenJDK 25 (EXE)"
$Publisher        = "Microsoft"
$AppName          = "Microsoft OpenJDK 25 (EXE)"
$BaseDownloadRoot = Join-Path $DownloadRoot "MsOpenJDK25-EXE"

# Inno Setup derives the ARP key name from the compiled AppId, which stays the
# same across JDK 25 releases; setup runs in 64-bit mode.
$ArpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9DA4BE12-C58B-3DE4-C047-AF46365065FC}_is1"


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A redirect that lands on an error page answers 200 with HTML, which
        would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}

# --- Functions ---


function ConvertFrom-MsOpenJdkDownloadUrl {
    <#
    .SYNOPSIS
        Returns the full version carried by a microsoft-jdk-<version>-windows-x64
        installer URL or file name for the given major and extension, or
        nothing for any other name.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Url,
        [Parameter(Mandatory)][int]$Major,
        [Parameter(Mandatory)][ValidateSet('msi', 'exe')][string]$Extension
    )

    $name = ($Url -split '[?#]')[0]
    $name = ($name -split '/')[-1]
    $pattern = '^microsoft-jdk-(?<ver>{0}(?:\.\d+){{1,4}})-windows-x64\.{1}$' -f $Major, $Extension
    $m = [regex]::Match($name, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return $null }
    return $m.Groups['ver'].Value
}


function ConvertFrom-MsOpenJdkSha256File {
    <#
    .SYNOPSIS
        Returns the lower-case SHA-256 hash that a .sha256sum.txt file lists
        for FileName, or nothing when no line names that file.
    .DESCRIPTION
        A line that carries a hash and no file name counts as a match.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content,
        [Parameter(Mandatory)][string]$FileName
    )

    foreach ($line in ($Content -split "`r?`n")) {
        $m = [regex]::Match($line, '^\s*(?<hash>[0-9A-Fa-f]{64})(?:\s+\*?(?<name>\S.*?))?\s*$')
        if (-not $m.Success) { continue }
        $name = $m.Groups['name'].Value
        if ([string]::IsNullOrEmpty($name) -or $name -ieq $FileName) {
            return $m.Groups['hash'].Value.ToLowerInvariant()
        }
    }
    return $null
}


function Get-LatestMsOpenJdkRelease {
    param([switch]$Quiet)

    Write-Log "Latest alias URL             : $LatestAliasUrl" -Quiet:$Quiet

    try {
        $target = (curl.exe -sIL --fail -o NUL -w '%{url_effective}' $LatestAliasUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve the latest-version redirect: $LatestAliasUrl" }

        $version = ConvertFrom-MsOpenJdkDownloadUrl -Url $target -Major $FeatureVersion -Extension 'msi'
        if (-not $version) { throw "Redirect target does not name a JDK $FeatureVersion x64 MSI: $target" }

        $ext = $InstallerType.ToLowerInvariant()
        $fileName = "microsoft-jdk-$version-windows-x64.$ext"

        Write-Log "Latest OpenJDK $FeatureVersion version    : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = $DownloadBaseUrl + $fileName
            Sha256Url   = $DownloadBaseUrl + $fileName + '.sha256sum.txt'
        }
    }
    catch {
        Write-Log "Failed to get Microsoft Build of OpenJDK $FeatureVersion version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Assert-MsOpenJdkFileHash {
    <#
    .SYNOPSIS
        Throws unless the file matches the hash in the vendor .sha256sum.txt.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Sha256Url
    )

    $content = (curl.exe -L --fail --silent --show-error $Sha256Url) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to download checksum file: $Sha256Url" }

    $fileName = Split-Path -Path $Path -Leaf
    $expected = ConvertFrom-MsOpenJdkSha256File -Content $content -FileName $fileName
    if (-not $expected) { throw "Checksum file names no hash for $fileName : $Sha256Url" }

    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    Write-Log "SHA-256 expected             : $expected"
    Write-Log "SHA-256 actual               : $actual"
    if ($actual -ne $expected) { throw "SHA-256 mismatch for $fileName." }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageMsOpenJdk25Exe {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$AppName - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest release info ---
    $release = Get-LatestMsOpenJdkRelease
    if (-not $release) { throw "Could not get release info." }

    $installerFileName = $release.FileName
    $version           = $release.Version

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log ""

    # --- Download and verify ---
    $localInstaller = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localInstaller"

    if (Test-Path -LiteralPath $localInstaller) {
        try {
            Assert-MsOpenJdkFileHash -Path $localInstaller -Sha256Url $release.Sha256Url
            Write-Log "Local installer exists and matches. Skipping download."
        }
        catch {
            Write-Log "Local installer does not verify; downloading again." -Level WARN
            Remove-Item -LiteralPath $localInstaller -Force -ErrorAction Stop
        }
    }
    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localInstaller
        Assert-MsOpenJdkFileHash -Path $localInstaller -Sha256Url $release.Sha256Url
    }

    Assert-PayloadIsExecutable -Path $localInstaller
    Assert-ArpDetectionKey -InstallerPath $localInstaller -ExpectedKey $ArpRegistryKey -Is64BitView $true

    # --- Read the ARP DisplayVersion from the Inno Setup header ---
    $analysis = Get-InstallerAnalysis -Path $localInstaller
    $displayVersion = [string]$analysis.SoftwareVersion
    if ($displayVersion -notmatch ('^{0}\.\d+(\.\d+)*$' -f $FeatureVersion)) {
        throw "Installer header carries no JDK $FeatureVersion DisplayVersion: '$displayVersion'"
    }
    Write-Log "Installer engine             : $($analysis.InstallerType)"
    Write-Log "ARP DisplayVersion           : $displayVersion"

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localInstaller -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrappers = New-ExeWrapperContent -InstallerFileName $installerFileName `
        -InstallArgs "'/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-', '/ALLUSERS', '/TASKS=`"FeatureEnvironment,FeatureJarFileRunWith`"'" `
        -UninstallCommand 'unused'

    # Inno Setup names the uninstaller unins###.exe by install order, so the
    # ARP UninstallString is the only value that names the right one.
    $uninstallContent = @'
$key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9DA4BE12-C58B-3DE4-C047-AF46365065FC}_is1'
if (-not (Test-Path -LiteralPath $key)) { exit 0 }
$cmd = (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue).UninstallString
if (-not $cmd) { exit 0 }
if ($cmd -match '^"([^"]+)"') { $exe = $matches[1] } else { $exe = $cmd.Trim() }
if (-not (Test-Path -LiteralPath $exe)) { exit 0 }
$proc = Start-Process -FilePath $exe -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $uninstallContent

    Write-Log ""
    Write-Log "ARP RegistryKey              : HKLM\$ArpRegistryKey"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /ALLUSERS /TASKS="FeatureEnvironment,FeatureJarFileRunWith"'
        UninstallArgs   = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        RunningProcess  = @("java", "javaw")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $ArpRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $displayVersion
            Is64Bit             = $true
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

function Invoke-PackageMsOpenJdk25Exe {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$AppName - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $rel = Get-LatestMsOpenJdkRelease -Quiet
        if (-not $rel) { exit 1 }
        Write-Output $rel.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("$AppName GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$AppName Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageMsOpenJdk25Exe
    }
    elseif ($PackageOnly) {
        Invoke-PackageMsOpenJdk25Exe
    }
    else {
        Invoke-StageMsOpenJdk25Exe
        Invoke-PackageMsOpenJdk25Exe
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-ms-openjdk25-exe'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
