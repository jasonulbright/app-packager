<#
Vendor: Salesforce (Tableau)
App: Tableau Prep Builder (x64)
CMName: Tableau Prep Builder
VendorUrl: https://www.tableau.com/products/prep
CPE: cpe:2.3:a:tableau:prep_builder:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.tableau.com/support/releases
DownloadPageUrl: https://www.tableau.com/products/prep/download
IconSource: Installer
UpdateCadenceDays: 90
WsusSupport: No (DetectionNotMappable)
ArpName: Tableau Prep*

.SYNOPSIS
    Packages Tableau Prep Builder (x64) for ConfigMgr.

.DESCRIPTION
    Reads the latest release from the Tableau release pages and downloads the
    installer from the Tableau CDN. Both refuse requests that do not carry
    browser request headers, so every request sends them.

    The Prep page names the release train only (2026.3). The point release comes
    from the page of Desktop and Reader for the same train (2026.3.1) and from a
    HEAD probe of the download address: a point release that does not exist is
    stepped down, one that no page lists yet is found by stepping up. The host
    answers a request without a sign-in with 403, which says nothing about the
    file: a 403 never counts as "does not exist", the version stays at what the
    pages say, and the log says that it was not confirmed. The download then
    decides and stops with the status if the file is not there.

    The installer is 1.1 GB. The download goes to a .part file, continues where
    it stopped after a lost connection (also in a later run), logs its size and
    speed every few seconds, and takes its name only when the size matches the
    size that the server announced. -SourceFolder stages
    the newest TableauPrep-<yyyy>-<m>-<p>.exe from a local folder
    instead.

    The installer is a WiX Burn bundle. Detection searches the bundle's
    64-bit uninstall entries by display name and compares DisplayVersion with
    the installer's ProductVersion, because the entry key changes with every
    release. Uninstall runs that entry's QuietUninstallString as written,
    never the hidden MSI entry.

    Supports two-phase operation:
      -StageOnly    Download the installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Tableau\Tableau Prep Builder (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging.
    Each packager creates a subfolder under this path (for example, <DownloadRoot>\TableauPrep).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 60

.PARAMETER SourceFolder
    Folder that holds the Tableau Prep Builder installer. When set, Stage uses the
    newest matching installer there instead of downloading.

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs the latest release version and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (ConfigurationManager PowerShell module available)
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
$PackagerName       = "package-tableauprep"
$ProductLabel       = "Tableau Prep Builder"
$InstallerFilter    = "TableauPrep-*.exe"
$DisplayNamePattern = "Tableau Prep*"
$InstallSwitches    = @('/install', '/quiet', '/norestart', 'ACCEPTEULA=1', 'SENDTELEMETRY=0')
$RunningProcesses   = @('tabprep')
$ReleasePageUrl     = "https://www.tableau.com/support/releases/prep"
# The Prep page names the release train (2026.3) and no point release. The page of Tableau Desktop and Reader lists the
# point release of the same train (2026.3.1), and a probe of the download address settles what the pages do not say.
$TrainPageUrl       = "https://www.tableau.com/support/releases"
# {0} is the hyphenated release (2026-3-1), {1} the dotted one (2026.3.1). The path is the one that the Tableau download
# page itself uses for a signed-in user (the Desktop packager uses the same esdalt host). An earlier template, tssoftware/,
# was never shown to work: no stage of it ever ran against the vendor host.
$DownloadUrlTemplate = "https://downloads.tableau.com/esdalt/tableau_prep/{1}/TableauPrep-{0}.exe"

$VendorFolder = "Tableau"
$AppFolder    = "Tableau Prep Builder (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "TableauPrep"

# --- Functions ---

function Get-TableauVersionFromFileName {
    <#
    .SYNOPSIS
        Returns the release version from an installer name such as
        TableauPrep-2026-2-2.exe (2026.2.2).
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


function Get-TableauTrainPatches {
    <#
    .SYNOPSIS
        The point releases that release page HTML lists for one release train (year.month).
    .DESCRIPTION
        Returns the patch numbers (0 for the bare train name, as in 2026.3) found for the train. A page that names only the
        train gives 0; a page that names 2026.3.1 gives 1 as well.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory)][string]$Train
    )
    $parts = $Train -split '\.'
    $patches = New-Object System.Collections.Generic.List[int]
    foreach ($m in [regex]::Matches($Html, '\b(20\d{2})\.(\d+)(?:\.(\d+))?\b')) {
        if ($m.Groups[1].Value -ne $parts[0] -or $m.Groups[2].Value -ne $parts[1]) { continue }
        $patches.Add($(if ($m.Groups[3].Success) { [int]$m.Groups[3].Value } else { 0 }))
    }
    return @($patches | Sort-Object -Unique)
}


function ConvertTo-TableauProbeResult {
    <#
    .SYNOPSIS
        Reads the answer of a HEAD request (status|content type|final address) as Exists, Missing or Unknown.
    #>
    param([AllowEmptyString()][string]$Answer)
    $f = @($Answer -split '\|', 3)
    if ($f.Count -lt 3) { return 'Unknown' }
    $code = $f[0]; $type = $f[1]; $effective = $f[2]
    if ($code -in @('404', '410')) { return 'Missing' }
    if ($code -in @('200', '206') -and $type -notmatch 'text/html' -and $effective -match '^https://downloads\.tableau\.com/') { return 'Exists' }
    return 'Unknown'
}


function Test-TableauInstallerUrl {
    <#
    .SYNOPSIS
        Asks the download host whether an installer address exists: Exists, Missing or Unknown.
    .DESCRIPTION
        A HEAD request with the browser headers. Exists: status 200 or 206 for a file (not a web page, and not a
        redirect to a sign-in page). Missing: status 404 or 410. Every other answer, above all 403, is Unknown: the host
        refuses a request that carries no sign-in, so a 403 says nothing about the file and is never read as "missing".
    #>
    param([Parameter(Mandatory)][string]$Url)
    $out = ''
    try {
        $out = ([string]((& curl.exe -I -L --silent --max-time 30 --max-redirs 5 -o NUL -w '%{http_code}|%{content_type}|%{url_effective}' @(Get-TableauBrowserCurlArgs) $Url 2>$null) -join '')).Trim()
    }
    catch { return 'Unknown' }
    return (ConvertTo-TableauProbeResult -Answer $out)
}


function Resolve-TableauPrepVersion {
    <#
    .SYNOPSIS
        The Prep release to package: the newest train of the Prep page with its newest point release.
    .DESCRIPTION
        The Prep page names the train only (2026.3). The point release comes from the highest patch that the Prep page or
        the Desktop and Reader page lists for the same train. The probe (Exists, Missing or Unknown for a version) then
        settles it: a listed patch that the probe says is Missing is stepped down until one exists; a patch that Exists
        is stepped up while the next one exists too, so that a point release that no page lists yet is found. A probe that
        cannot tell (Unknown, the usual answer of the Tableau host to a request without a sign-in) never changes the
        version: it stays at what the pages say, and the result says that it was not confirmed. Returns Version, State
        (Exists, Unknown or Missing) and Notes (lines for the log).
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$PrepHtml,
        [AllowEmptyString()][string]$TrainHtml = '',
        [Parameter(Mandatory)][scriptblock]$Probe
    )
    $notes = New-Object System.Collections.Generic.List[string]
    $seen = foreach ($m in [regex]::Matches($PrepHtml, '\b(20\d{2})\.(\d+)(?:\.(\d+))?\b')) { [version]('{0}.{1}' -f $m.Groups[1].Value, $m.Groups[2].Value) }
    $newest = @($seen | Sort-Object -Descending | Select-Object -First 1)
    if ($newest.Count -eq 0) { return $null }
    $train = '{0}.{1}' -f $newest[0].Major, $newest[0].Minor

    $patches = @(Get-TableauTrainPatches -Html $PrepHtml -Train $train) + @(Get-TableauTrainPatches -Html $TrainHtml -Train $train)
    $patch = [int](@($patches | Measure-Object -Maximum).Maximum)
    $from = if ($patch -gt 0) { 'a page lists it' } else { 'no page lists a point release' }
    $notes.Add("Release train                : $train; start at $train.$patch ($from)")

    $state = & $Probe "$train.$patch"
    while ($state -eq 'Missing' -and $patch -gt 0) {
        $notes.Add("Prep $train.$patch does not exist on the download host; trying $train.$($patch - 1)")
        $patch--
        $state = & $Probe "$train.$patch"
    }
    if ($state -eq 'Exists') {
        for ($i = 0; $i -lt 20; $i++) {
            $next = & $Probe ('{0}.{1}' -f $train, ($patch + 1))
            if ($next -ne 'Exists') {
                if ($next -eq 'Unknown') { $notes.Add("The download host cannot tell whether $train.$($patch + 1) exists; staying at $train.$patch") }
                break
            }
            $patch++
            $notes.Add("A point release that no page lists exists: $train.$patch")
        }
    }
    elseif ($state -eq 'Unknown') {
        $notes.Add("The download host cannot tell whether Prep $train.$patch exists (it answers a request without a sign-in with 403). The version stays at what the pages say. The download decides: it stops with the status if the file is not there.")
    }
    else {
        $notes.Add("No Prep installer of the train $train exists on the download host.")
    }
    return [pscustomobject]@{ Version = ('{0}.{1}' -f $train, $patch); State = $state; Notes = @($notes) }
}


function Get-LatestTableauRelease {
    param([switch]$Quiet)

    Write-Log "Tableau release page         : $ReleasePageUrl" -Quiet:$Quiet

    $html = (curl.exe -L --fail --silent --show-error @(Get-TableauBrowserCurlArgs) $ReleasePageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to read the Tableau release page: $ReleasePageUrl" }

    # The page of Desktop and Reader lists the point release of the train. When it cannot be read, the probe and the Prep page decide.
    $trainHtml = ''
    $trainText = (curl.exe -L --fail --silent --show-error @(Get-TableauBrowserCurlArgs) $TrainPageUrl 2>$null) -join "`n"
    if ($LASTEXITCODE -eq 0) { $trainHtml = $trainText }
    else { Write-Log "The release page of Desktop and Reader could not be read ($TrainPageUrl); the Prep page and the probe decide." -Level WARN -Quiet:$Quiet }

    $probe = { param($v) Test-TableauInstallerUrl -Url ($DownloadUrlTemplate -f ($v -replace '\.', '-'), $v) }
    $resolved = Resolve-TableauPrepVersion -PrepHtml $html -TrainHtml $trainHtml -Probe $probe
    if (-not $resolved) { throw "No yyyy.m release found on $ReleasePageUrl" }
    foreach ($line in $resolved.Notes) { Write-Log $line -Level $(if ($resolved.State -eq 'Exists') { 'INFO' } else { 'WARN' }) -Quiet:$Quiet }
    if ($resolved.State -eq 'Missing') { throw "No Tableau Prep installer exists for the release train of $ReleasePageUrl." }
    $version = $resolved.Version

    $url = $DownloadUrlTemplate -f ($version -replace '\.', '-'), $version
    Write-Log ("Latest release               : {0} ({1})" -f $version, $(if ($resolved.State -eq 'Exists') { 'the download host confirms the file' } else { 'not confirmed by the download host' })) -Quiet:$Quiet
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
        Returns script text that enumerates the bundle's 64-bit uninstall
        entries matching the display name pattern.
    .DESCRIPTION
        The bundle writes its entry to the 64-bit view with a
        QuietUninstallString that runs the cached bundle executable; the MSI
        it installs writes a hidden entry with no QuietUninstallString that
        must not be used. The view is opened explicitly so the script behaves
        the same in a 32-bit or 64-bit host.
    #>
    param([Parameter(Mandatory)][string]$Pattern)

    return (
        '$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)',
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

    return (
        (New-TableauBundleSearchText -Pattern $Pattern),
        'if ($bundles.Count -eq 0) { exit 0 }',
        '$exitCode = 0',
        'foreach ($b in $bundles) {',
        '    $m = [regex]::Match($b.QuietUninstall, ''^\s*"(?<exe>[^"]+\.exe)"\s+(?<args>.+)$'', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)',
        '    if (-not $m.Success -or -not (Test-Path -LiteralPath $m.Groups[''exe''].Value)) { Write-Error ("Bundle uninstaller not found: " + $b.QuietUninstall); $exitCode = 1; continue }',
        '    $proc = Start-Process -FilePath $m.Groups[''exe''].Value -ArgumentList $m.Groups[''args''].Value -Wait -PassThru -NoNewWindow',
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
            Invoke-ResumableDownload -Url $release.DownloadUrl -OutFile $localExe -ExtraCurlArgs (Get-TableauBrowserCurlArgs)
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
    Write-Log "Detection                    : 64-bit bundle entry '$DisplayNamePattern', DisplayVersion >= $productVersion"
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
        UninstallArgs   = ''
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
