<#
Vendor: Salesforce (Tableau)
App: Tableau Desktop (x64)
CMName: Tableau Desktop
VendorUrl: https://www.tableau.com/products/desktop
CPE: cpe:2.3:a:tableau:tableau:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.tableau.com/support/releases
DownloadPageUrl: https://www.tableau.com/products/desktop/download
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages Tableau Desktop (x64) for MECM.

.DESCRIPTION
    Reads the latest release from the Tableau release page and downloads the
    installer from the Tableau CDN. Both refuse requests that do not carry
    browser request headers, so every request sends them. -SourceFolder stages
    the newest TableauDesktop-64bit-<yyyy>-<m>-<p>.exe from a local folder
    instead.

    The installer is a WiX Burn bundle. Detection searches the bundle's
    32-bit uninstall entries by display name and compares DisplayVersion with
    the installer's ProductVersion, because the entry key changes with every
    release. Uninstall runs the executable named in that entry's
    QuietUninstallString with /uninstall /quiet /norestart, never the 64-bit
    MSI entry.

    Supports two-phase operation:
      -StageOnly    Download the installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Tableau\Tableau Desktop (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TableauDesktop).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 60

.PARAMETER SourceFolder
    Folder that holds the Tableau Desktop installer. When set, Stage uses the
    newest matching installer there instead of downloading.

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs the latest release version and exits.

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
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 60,
    [string]$SourceFolder = "",
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
$PackagerName       = "package-tableaudesktop"
$ProductLabel       = "Tableau Desktop"
$InstallerFilter    = "TableauDesktop-64bit-*.exe"
$DisplayNamePattern = "Tableau 20*"
$InstallSwitches    = @('/install', '/quiet', '/norestart', 'ACCEPTEULA=1', 'REMOVEINSTALLEDAPP=1', 'SENDTELEMETRY=0')
$UninstallSwitches  = @('/uninstall', '/quiet', '/norestart')
$RunningProcesses   = @('tableau')
$ReleasePageUrl     = "https://www.tableau.com/support/releases"
# {0} is the hyphenated release (2026-2-2), {1} the dotted one (2026.2.2).
$DownloadUrlTemplate = "https://downloads.tableau.com/esdalt/{1}/TableauDesktop-64bit-{0}.exe"

$VendorFolder = "Tableau"
$AppFolder    = "Tableau Desktop (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "TableauDesktop"

# --- Functions ---

function Get-TableauVersionFromFileName {
    <#
    .SYNOPSIS
        Returns the release version from an installer name such as
        TableauDesktop-64bit-2026-2-2.exe (2026.2.2).
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$FileName)

    $m = [regex]::Match($FileName, '-(?<y>20\d{2})-(?<m>\d+)-(?<p>\d+)\.exe$')
    if (-not $m.Success) { return $null }
    return ('{0}.{1}.{2}' -f $m.Groups['y'].Value, $m.Groups['m'].Value, $m.Groups['p'].Value)
}


function Get-TableauBrowserCurlArgs {
    <#
    .SYNOPSIS
        Returns the curl.exe arguments for the request headers a browser sends.
    .DESCRIPTION
        The Tableau release pages and CDN answer 403 to requests without them.
    #>
    return @(
        '--compressed',
        '-A', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0',
        '-H', 'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        '-H', 'Accept-Language: en-US,en;q=0.9',
        '-H', 'Referer: https://www.tableau.com/',
        '-H', 'Sec-Fetch-Dest: document',
        '-H', 'Sec-Fetch-Mode: navigate',
        '-H', 'Sec-Fetch-Site: same-site',
        '-H', 'Sec-Fetch-User: ?1',
        '-H', 'Upgrade-Insecure-Requests: 1'
    )
}


function Get-TableauVersionFromPage {
    <#
    .SYNOPSIS
        Returns the highest yyyy.m.p release found in release page HTML.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html)

    $versions = foreach ($m in [regex]::Matches($Html, '\b(20\d{2})\.(\d+)(?:\.(\d+))?\b')) {
        $patch = if ($m.Groups[3].Success) { $m.Groups[3].Value } else { '0' }
        [version]('{0}.{1}.{2}' -f $m.Groups[1].Value, $m.Groups[2].Value, $patch)
    }
    $latest = @($versions | Sort-Object -Descending | Select-Object -First 1)
    if ($latest.Count -eq 0) { return $null }
    return $latest[0].ToString()
}


function Get-LatestTableauRelease {
    param([switch]$Quiet)

    Write-Log "Tableau release page         : $ReleasePageUrl" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error @(Get-TableauBrowserCurlArgs) $ReleasePageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to read the Tableau release page: $ReleasePageUrl" }

    $version = Get-TableauVersionFromPage -Html $html
    if (-not $version) { throw "No yyyy.m.p release found on $ReleasePageUrl" }

    $url = $DownloadUrlTemplate -f ($version -replace '\.', '-'), $version
    Write-Log "Latest release               : $version" -Quiet:$Quiet
    return [pscustomobject]@{ Version = $version; DownloadUrl = $url; FileName = [System.IO.Path]::GetFileName($url) }
}


function Get-TableauLocalInstaller {
    <#
    .SYNOPSIS
        Resolves the newest local installer and its release version.
    #>
    param([switch]$Quiet)

    $path = Resolve-LocalSourceInstaller -PackagerName $PackagerName -Filter $InstallerFilter -Override $SourceFolder
    $version = Get-TableauVersionFromFileName -FileName (Split-Path -Leaf $path)
    if (-not $version) {
        throw "Installer name does not carry a <yyyy>-<m>-<p> release version: $path"
    }
    Write-Log "Local installer              : $path" -Quiet:$Quiet
    Write-Log "Release version              : $version" -Quiet:$Quiet
    return [pscustomobject]@{ Path = $path; Version = $version }
}


function New-TableauBundleSearchText {
    <#
    .SYNOPSIS
        Returns script text that enumerates the bundle's 32-bit uninstall
        entries matching the display name pattern.
    .DESCRIPTION
        The bundle writes its entry to the 32-bit view with a
        QuietUninstallString that runs the cached bundle executable; the MSI
        it installs writes a hidden 64-bit entry that must not be used. The
        view is opened explicitly so the script behaves the same in a 32-bit
        or 64-bit host.
    #>
    param([Parameter(Mandatory)][string]$Pattern)

    return (
        '$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)',
        '$un = $base.OpenSubKey(''SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'')',
        '$bundles = @()',
        'if ($un) {',
        '    foreach ($name in $un.GetSubKeyNames()) {',
        '        $k = $un.OpenSubKey($name)',
        '        if (-not $k) { continue }',
        '        $entry = [pscustomobject]@{ DisplayName = [string]$k.GetValue(''DisplayName''); DisplayVersion = [string]$k.GetValue(''DisplayVersion''); QuietUninstall = [string]$k.GetValue(''QuietUninstallString'') }',
        '        $k.Close()',
        ('        if ($entry.DisplayName -like ''{0}'' -and $entry.QuietUninstall) {{ $bundles += $entry }}' -f $Pattern),
        '    }',
        '    $un.Close()',
        '}'
    ) -join "`r`n"
}


function New-TableauDetectionScript {
    param(
        [Parameter(Mandatory)][string]$Pattern,
        [Parameter(Mandatory)][string]$MinimumVersion
    )

    return (
        (New-TableauBundleSearchText -Pattern $Pattern),
        ('$minimum = [version]''{0}''' -f $MinimumVersion),
        'foreach ($b in $bundles) {',
        '    $v = $null',
        '    if ([version]::TryParse($b.DisplayVersion, [ref]$v) -and $v -ge $minimum) { Write-Output ''Installed''; exit 0 }',
        '}',
        'exit 0'
    ) -join "`r`n"
}


function New-TableauUninstallScript {
    param([Parameter(Mandatory)][string]$Pattern)

    $switchList = ($UninstallSwitches | ForEach-Object { "'$_'" }) -join ', '
    return (
        (New-TableauBundleSearchText -Pattern $Pattern),
        'if ($bundles.Count -eq 0) { exit 0 }',
        '$exitCode = 0',
        'foreach ($b in $bundles) {',
        '    $exe = ($b.QuietUninstall -replace ''^\s*"([^"]+)".*$'', ''$1'').Trim()',
        '    if (-not $exe.EndsWith(''.exe'', [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $exe)) { Write-Error ("Bundle uninstaller not found: " + $b.QuietUninstall); $exitCode = 1; continue }',
        ('    $proc = Start-Process -FilePath $exe -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $switchList),
        '    if ($proc.ExitCode -ne 0) { $exitCode = $proc.ExitCode }',
        '}',
        'exit $exitCode'
    ) -join "`r`n"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTableau {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$ProductLabel (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    if ($SourceFolder) {
        $installer = Get-TableauLocalInstaller
    }
    else {
        $release = Get-LatestTableauRelease
        $localExe = Join-Path $BaseDownloadRoot $release.FileName
        Write-Log "Download URL                 : $($release.DownloadUrl)"
        if (-not (Test-Path -LiteralPath $localExe)) {
            Write-Log "Downloading $ProductLabel..."
            Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localExe -ExtraCurlArgs (Get-TableauBrowserCurlArgs)
            if (-not (Test-Path -LiteralPath $localExe)) { throw "Download produced no file: $($release.DownloadUrl)" }
        }
        else {
            Write-Log "Local installer exists. Skipping download."
        }
        $installer = [pscustomobject]@{ Path = $localExe; Version = $release.Version }
    }
    $version = $installer.Version
    $installerFileName = Split-Path -Leaf $installer.Path

    $signature = Get-AuthenticodeSignature -LiteralPath $installer.Path
    if ($signature.Status -ne 'Valid') {
        throw "Installer signature is $($signature.Status), not Valid: $($installer.Path)"
    }
    Write-Log "Installer signer             : $($signature.SignerCertificate.Subject)"

    $productVersion = ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($installer.Path).ProductVersion)
    $parsed = $null
    if (-not $productVersion -or -not [version]::TryParse($productVersion.Trim(), [ref]$parsed)) {
        throw "Installer carries no numeric ProductVersion to compare with the installed DisplayVersion: $($installer.Path)"
    }
    $productVersion = $productVersion.Trim()
    Write-Log "Installer ProductVersion     : $productVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    Copy-Item -LiteralPath $installer.Path -Destination $stagedExe -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $stagedExe"

    # --- Generate content wrappers ---
    $installArgList = ($InstallSwitches | ForEach-Object { "'$_'" }) -join ', '
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $installArgList),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content (New-TableauUninstallScript -Pattern $DisplayNamePattern)

    Write-Log ""
    Write-Log "Detection                    : 32-bit bundle entry '$DisplayNamePattern', DisplayVersion >= $productVersion"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $ProductLabel
        Publisher       = "Salesforce"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = ($InstallSwitches -join ' ')
        UninstallArgs   = ($UninstallSwitches -join ' ')
        RunningProcess  = $RunningProcesses
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = (New-TableauDetectionScript -Pattern $DisplayNamePattern -MinimumVersion $productVersion)
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

function Invoke-PackageTableau {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$ProductLabel (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifest = Read-StageManifest -Path (Join-Path $localContentPath "stage-manifest.json")

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
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
        $info = if ($SourceFolder) { Get-TableauLocalInstaller -Quiet } else { Get-LatestTableauRelease -Quiet }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("$ProductLabel GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "$ProductLabel (x64) Auto-Packager starting"
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
        Invoke-StageTableau
    }
    elseif ($PackageOnly) {
        Invoke-PackageTableau
    }
    else {
        Invoke-StageTableau
        Invoke-PackageTableau
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context $PackagerName
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
