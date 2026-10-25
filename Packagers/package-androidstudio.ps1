<#
Vendor: Google
App: Android Studio
CMName: Android Studio
VendorUrl: https://developer.android.com/studio
CPE: cpe:2.3:a:google:android_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://developer.android.com/studio/releases
DownloadPageUrl: https://developer.android.com/studio
UpdateCadenceDays: 45
WsusSupport: No (DetectionNotMappable)
ArpName: Android Studio

.SYNOPSIS
    Packages the latest stable Android Studio (x64) installer for ConfigMgr.

.DESCRIPTION
    Reads the vendor release feed, selects the highest version on the Release
    channel, downloads the matching Windows installer, stages content to a
    versioned local folder, and creates a ConfigMgr Application with script
    detection on the build number in the installed product-info.json.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    The payload is roughly 1.5 GB, so the Stage phase re-uses an existing local
    installer instead of downloading it again.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Google\Android Studio\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AndroidStudio).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest stable Android Studio version string and exits.

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
$ReleaseFeedUrl = "https://jb.gg/android-studio-releases-list.json"

$VendorFolder = "Google"
$AppFolder    = "Android Studio"

$BaseDownloadRoot = Join-Path $DownloadRoot "AndroidStudio"

$InstallDir = "{0}\Android\Android Studio" -f $env:ProgramFiles

# --- Functions ---


function Assert-ExecutablePayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        A CDN error page or consent interstitial answers 200 with HTML, which
        would otherwise stage as a valid-looking installer.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = New-Object byte[] 2
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($bytes, 0, 2) } finally { $stream.Dispose() }
    if ($read -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Resolve-AndroidStudioRelease {
    <#
    .SYNOPSIS
        Returns the highest Release-channel version and its Windows installer URL.
    .DESCRIPTION
        The feed carries every channel and every historical build in one list,
        so entries are filtered to the Release channel and ordered by version
        rather than trusting document order. Each release publishes both an
        installer (.exe) and a portable archive (.zip); only the installer path
        is taken.
    #>
    param([switch]$Quiet)

    Write-Log "Release feed URL             : $ReleaseFeedUrl" -Quiet:$Quiet

    try {
        $jsonText = (curl.exe -L --fail --silent --show-error $ReleaseFeedUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Android Studio release feed: $ReleaseFeedUrl" }

        $feed = ConvertFrom-Json $jsonText
        $items = @($feed.content.item | Where-Object { $_.channel -eq 'Release' -and $_.version -match '^\d+(\.\d+)+$' })
        if ($items.Count -lt 1) { throw "The release feed contained no Release-channel entries." }

        $latest = $items | Sort-Object { [version]$_.version } | Select-Object -Last 1

        $windowsInstaller = @($latest.download | Where-Object { $_.link -match '/install/[^/]+/[^/]+-windows\.exe$' }) | Select-Object -First 1
        if (-not $windowsInstaller) {
            throw "Release $($latest.version) publishes no Windows installer link."
        }

        Write-Log "Latest Android Studio version: $($latest.version)" -Quiet:$Quiet
        Write-Log "Resolved installer URL       : $($windowsInstaller.link)" -Quiet:$Quiet

        return [pscustomobject]@{
            Version  = $latest.version
            Url      = $windowsInstaller.link
            FileName = ($windowsInstaller.link -split '/')[-1]
            Build    = ([string]$latest.build) -replace '^[A-Za-z]+-', ''
            Checksum = $windowsInstaller.checksum
        }
    }
    catch {
        Write-Log "Failed to resolve the Android Studio release: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAndroidStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Android Studio (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Resolve release ---
    $release = Resolve-AndroidStudioRelease
    if (-not $release) { throw "Could not resolve the Android Studio release." }

    Write-Log "Version                      : $($release.Version)"
    Write-Log "Installer filename           : $($release.FileName)"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $release.FileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading installer (about 1.5 GB)..."
        Invoke-DownloadWithRetry -Url $release.Url -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    Assert-ExecutablePayload -Path $localExe

    # --- Verify the vendor-published checksum ---
    # The CDN has answered 200 with a short body; a cached file that fails the
    # checksum is replaced by one fresh download before the run gives up, so a
    # torn transfer cannot pin every later run to the same bad file.
    if ($release.Checksum) {
        $expected = $release.Checksum.ToUpperInvariant()
        $actual = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($actual -ne $expected) {
            Write-Log "Installer SHA256 ($actual) does not match the feed checksum ($expected); downloading again." -Level WARN
            Remove-Item -LiteralPath $localExe -Force -ErrorAction Stop
            Invoke-DownloadWithRetry -Url $release.Url -OutFile $localExe
            Assert-ExecutablePayload -Path $localExe
            $actual = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256 -ErrorAction Stop).Hash
            if ($actual -ne $expected) {
                Remove-Item -LiteralPath $localExe -Force -ErrorAction SilentlyContinue
                throw "Installer SHA256 ($actual) does not match the feed checksum ($expected)."
            }
        }
        Write-Log "SHA256 matches the release feed."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $release.Version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $release.FileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # The installer is NSIS: /S installs silently to the default location and
    # the uninstaller sits at the root of that install directory.
    $uninstallExe = Join-Path $InstallDir "uninstall.exe"
    $wrappers = New-ExeWrapperContent -InstallerFileName $release.FileName `
        -InstallArgs "'/S'" `
        -UninstallCommand $uninstallExe `
        -UninstallArgs "'/S'"

    # A complete silent install ends with exit code 1223 (operation canceled);
    # the product files are present, so 1223 counts as success only when
    # studio64.exe exists.
    $installContent = $wrappers.Install -replace '(?m)^exit \$proc\.ExitCode\s*$', (
        "`$code = `$proc.ExitCode`r`n" +
        "if (`$code -eq 1223 -and (Test-Path -LiteralPath (Join-Path `$env:ProgramFiles 'Android\Android Studio\bin\studio64.exe'))) { `$code = 0 }`r`n" +
        "exit `$code")
    if ($installContent -eq $wrappers.Install) { throw "The install wrapper no longer ends with 'exit `$proc.ExitCode'." }

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $wrappers.Uninstall

    # --- Write stage manifest ---
    # studio64.exe carries the platform build number without the release
    # suffix and no usable file version, and the ARP DisplayVersion holds
    # only year and release train (2026.2). product-info.json carries the
    # full build number the release feed publishes, so the detection script
    # compares that.
    if ([string]::IsNullOrWhiteSpace($release.Build) -or $release.Build -notmatch '^\d+(\.\d+)+$') {
        throw "The release feed gave no usable build number for Android Studio $($release.Version)."
    }
    $detectionScript = @(
        '$root = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }'
        '$file = Join-Path $root ''Android\Android Studio\product-info.json'''
        'if (-not (Test-Path -LiteralPath $file)) { exit 0 }'
        'try {'
        '    $info = Get-Content -LiteralPath $file -Raw -ErrorAction Stop | ConvertFrom-Json'
        '    $have = @(([string]$info.buildNumber).Split(''.'') | ForEach-Object { [int64]$_ })'
        ('    $need = @({0})' -f (($release.Build.Split('.') | ForEach-Object { [int64]$_ }) -join ', '))
        '}'
        'catch { exit 0 }'
        '$result = 0'
        'for ($i = 0; $i -lt [Math]::Max($have.Count, $need.Count); $i++) {'
        '    $a = if ($i -lt $have.Count) { $have[$i] } else { 0 }'
        '    $b = if ($i -lt $need.Count) { $need[$i] } else { 0 }'
        '    if ($a -ne $b) { $result = if ($a -gt $b) { 1 } else { -1 }; break }'
        '}'
        'if ($result -ge 0) { Write-Output ''Installed'' }'
        'exit 0'
    ) -join "`r`n"

    Write-Log ""
    Write-Log "Detection                    : product-info.json buildNumber >= $($release.Build)"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "Android Studio"
        Publisher        = "Google"
        SoftwareVersion  = $release.Version
        InstallerFile    = $release.FileName
        InstallerType    = "EXE"
        InstallArgs      = "/S"
        UninstallArgs    = "/S"
        UninstallCommand = $uninstallExe
        RunningProcess   = @("studio64")
        Detection        = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $release.Version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageAndroidStudio {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Android Studio (x64) - PACKAGE phase"
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
        $release = Resolve-AndroidStudioRelease -Quiet
        if (-not $release) { exit 1 }
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Android Studio GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Android Studio (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseFeedUrl               : $ReleaseFeedUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAndroidStudio
    }
    elseif ($PackageOnly) {
        Invoke-PackageAndroidStudio
    }
    else {
        Invoke-StageAndroidStudio
        Invoke-PackageAndroidStudio
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-androidstudio'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
