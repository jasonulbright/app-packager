<#
Vendor: WIBU-SYSTEMS
App: CodeMeter Runtime Kit
CMName: CodeMeter Runtime Kit
VendorUrl: https://www.wibu.com/
CPE: cpe:2.3:a:wibu:codemeter_runtime:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.wibu.com/support/user/user-software.html
DownloadPageUrl: https://www.wibu.com/us/support/user/downloads-user-software.html
IconSource: Installer
UpdateCadenceDays: 90

.SYNOPSIS
    Packages the CodeMeter Runtime Kit (x64) for MECM.

.DESCRIPTION
    Scrapes the vendor user-software download listing for the newest 64-bit
    Windows runtime, resolves its direct-download link, downloads the
    bootstrapper EXE, stages content to a versioned local folder, and creates
    an MECM Application with file-version detection.

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
    Content is staged under: <FileServerPath>\Applications\WIBU-SYSTEMS\CodeMeter Runtime Kit\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\CodeMeterRuntime).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available CodeMeter Runtime Kit version string and exits.

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
$DownloadListUrl = "https://www.wibu.com/us/support/user/user-software.html"
$FileDetailUrl   = "https://www.wibu.com/us/support/user/downloads-user-software/file/download/{0}.html"

$VendorFolder = "WIBU-SYSTEMS"
$AppFolder    = "CodeMeter Runtime Kit"

$BaseDownloadRoot  = Join-Path $DownloadRoot "CodeMeterRuntime"
$InstallerFileName = "CodeMeterRuntime.exe"

# The vendor site rejects the default curl user agent on the HTML pages.
$BrowserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The download handler answers 200 with an HTML page when the link token
        no longer matches, which would otherwise stage as a valid-looking EXE.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-LatestCodeMeterRuntime {
    <#
    .SYNOPSIS
        Returns the newest 64-bit CodeMeter Runtime version and its direct URL.
    .DESCRIPTION
        The listing renders each runtime download as an <option> carrying the
        file id and the version, grouped by architecture; the 64-bit group is
        the only one that still receives current releases. The direct-download
        link carries a server-issued token, so it is read from the file detail
        page rather than composed.
    #>
    param([switch]$Quiet)

    Write-Log "Download listing URL         : $DownloadListUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error -A $BrowserAgent $DownloadListUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the user software listing: $DownloadListUrl" }

        # Narrow to the Windows runtime block: the same page also lists Linux
        # and macOS runtimes whose option markup is identical.
        $blockRx = [regex]'(?s)CodeMeter User Runtime for Windows.*?</select>'
        $block = $blockRx.Match($html)
        if (-not $block.Success) {
            throw "Could not locate the Windows runtime download block on the listing page."
        }

        $groupRx = [regex]'(?s)<optgroup[^>]*label="Windows 64-bit"[^>]*>(?<body>.*?)</optgroup>'
        $group = $groupRx.Match($block.Value)
        if (-not $group.Success) {
            throw "Could not locate a 'Windows 64-bit' download group."
        }

        $optionRx = [regex]'<option\s+value="(?<id>\d+)"[^>]*>\s*Version\s+(?<ver>\d+\.\d+[a-z]?)'
        $options = $optionRx.Matches($group.Groups['body'].Value)
        if (-not $options -or $options.Count -lt 1) {
            throw "The 'Windows 64-bit' group carries no version entries."
        }

        # Maintenance releases of older lines are published alongside the
        # current one, so the highest numeric version wins rather than the
        # first entry.
        $candidates = foreach ($o in $options) {
            $raw = $o.Groups['ver'].Value
            [pscustomobject]@{
                Id         = $o.Groups['id'].Value
                RawVersion = $raw
                SortKey    = [version]([regex]::Replace($raw, '[a-z]$', ''))
            }
        }
        $best = $candidates | Sort-Object SortKey -Descending | Select-Object -First 1

        $detailUrl = $FileDetailUrl -f $best.Id
        $detail = (curl.exe -L --fail --silent --show-error -A $BrowserAgent $detailUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the file detail page: $detailUrl" }

        $linkRx = [regex]'href="(?<href>[^"]*directDownload[^"]*)"'
        $link = $linkRx.Match($detail)
        if (-not $link.Success) {
            throw "Could not locate a direct-download link on the file detail page: $detailUrl"
        }

        $href = [System.Net.WebUtility]::HtmlDecode($link.Groups['href'].Value)
        $url  = ([uri]::new([uri]"https://www.wibu.com/", $href)).AbsoluteUri

        Write-Log "Listed runtime version       : $($best.RawVersion)" -Quiet:$Quiet

        return [pscustomobject]@{
            ListedVersion = $best.RawVersion
            DownloadUrl   = $url
        }
    }
    catch {
        Write-Log "Failed to get CodeMeter Runtime version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-CodeMeterInstaller {
    <#
    .SYNOPSIS
        Downloads the runtime bootstrapper and returns its path and version.
    .DESCRIPTION
        The listing shows a two-segment version (9.10) while the product and
        its ARP entry carry the full four-segment build, so the version comes
        from the downloaded binary.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestCodeMeterRuntime -Quiet:$Quiet
    if (-not $releaseInfo) { throw "Could not resolve the CodeMeter Runtime download." }

    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localExe -ExtraCurlArgs @('-A', $BrowserAgent) -Quiet:$Quiet

    Assert-PayloadIsExecutable -Path $localExe

    $fileVersion = (Get-Item -LiteralPath $localExe -ErrorAction Stop).VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($fileVersion)) {
        throw "The runtime bootstrapper carries no file version; cannot derive the product version."
    }
    $fileVersion = $fileVersion.Trim()

    return [pscustomobject]@{
        Path          = $localExe
        Version       = $fileVersion
        ListedVersion = $releaseInfo.ListedVersion
        DownloadUrl   = $releaseInfo.DownloadUrl
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCodeMeter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CodeMeter Runtime Kit (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    Write-Log "Downloading installer..."
    $installer = Get-CodeMeterInstaller

    $version  = $installer.Version
    $localExe = $installer.Path

    Write-Log "Download URL                 : $($installer.DownloadUrl)"
    Write-Log "Listed version               : $($installer.ListedVersion)"
    Write-Log "Binary file version          : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # The bootstrapper forwards /ComponentArgs to the MSI it wraps, and its
    # value carries embedded quotes; a single ArgumentList string reaches the
    # process command line unaltered, which an array of elements would not.
    $installContent = @"
`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'
`$proc = Start-Process -FilePath `$exePath -ArgumentList '/q /nosplash /ComponentArgs "*":"/quiet /norestart"' -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    # The wrapped MSI gets a new ProductCode every release, so the uninstall
    # command is read from the ARP entry the install wrote.
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
        Where-Object { $_.DisplayName -like 'CodeMeter Runtime Kit*' } |
        Select-Object -First 1
    if ($entry) { break }
}
if (-not $entry) { exit 0 }
$code = $entry.PSChildName
if ($code -notmatch '^\{[0-9A-Fa-f-]{36}\}$') { exit 0 }
$proc = Start-Process msiexec.exe -ArgumentList @('/x', $code, '/qn', '/norestart') -Wait -PassThru -NoNewWindow
exit $proc.ExitCode
'@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # The runtime has shipped under both Program Files paths across releases
    # and the wrapped MSI's ProductCode is not readable before install, so
    # detection matches the service binary in either location.
    $detection = @{
        Type      = "Compound"
        Connector = "Or"
        Clauses   = @(
            @{
                Type          = "File"
                FilePath      = "{0}\CodeMeter\Runtime\bin" -f $env:ProgramFiles
                FileName      = "CodeMeter.exe"
                PropertyType  = "Version"
                Operator      = "GreaterEquals"
                ExpectedValue = $version
                Is64Bit       = $true
            },
            @{
                Type          = "File"
                FilePath      = "{0}\CodeMeter\Runtime\bin" -f ${env:ProgramFiles(x86)}
                FileName      = "CodeMeter.exe"
                PropertyType  = "Version"
                Operator      = "GreaterEquals"
                ExpectedValue = $version
                Is64Bit       = $true
            }
        )
    }

    Write-Log ""
    Write-Log "Detection file version       : $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "CodeMeter Runtime Kit"
        Publisher       = "WIBU-SYSTEMS"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "EXE"
        InstallArgs     = '/q /nosplash /ComponentArgs "*":"/quiet /norestart"'
        UninstallArgs   = "/qn /norestart"
        RunningProcess  = @("CodeMeter", "CodeMeterCC")
        Detection       = $detection
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

function Invoke-PackageCodeMeter {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CodeMeter Runtime Kit (x64) - PACKAGE phase"
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
        $installer = Get-CodeMeterInstaller -Quiet
        Write-Output $installer.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("CodeMeter Runtime GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CodeMeter Runtime Kit (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadListUrl              : $DownloadListUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCodeMeter
    }
    elseif ($PackageOnly) {
        Invoke-PackageCodeMeter
    }
    else {
        Invoke-StageCodeMeter
        Invoke-PackageCodeMeter
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-codemeterruntime'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
