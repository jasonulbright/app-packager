<#
Vendor: Mozilla
App: Mozilla Firefox
CMName: Mozilla Firefox
SupportsVariants: Architecture, Language
VendorUrl: https://www.mozilla.org/firefox/enterprise/
CPE: cpe:2.3:a:mozilla:firefox:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.mozilla.org/en-US/firefox/releases/
DownloadPageUrl: https://www.mozilla.org/en-US/firefox/enterprise/

.SYNOPSIS
    Packages Mozilla Firefox (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest Mozilla Firefox x64 MSI from the official Mozilla
    release servers, stages content to a versioned local folder with file-based
    detection metadata, and creates an MECM Application with file-version-based
    detection.
    Detection uses firefox.exe version >= packaged version in the Program Files
    install path.

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
    Content is staged under: <FileServerPath>\Applications\Mozilla\Firefox\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Firefox).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download MSI, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Firefox version string and exits.

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
$VersionsJsonUrl  = "https://product-details.mozilla.org/1.0/firefox_versions.json"
$DownloadBase     = "https://releases.mozilla.org/pub/firefox/releases"

$VendorFolder = "Mozilla"
$AppFolder    = "Firefox"

$BaseDownloadRoot = Join-Path $DownloadRoot "Firefox"

# --- Functions ---


function Get-LatestFirefoxVersion {
    param([switch]$Quiet)

    Write-Log "Versions JSON URL            : $VersionsJsonUrl" -Quiet:$Quiet

    try {
        $jsonText = (curl.exe -L --fail --silent --show-error $VersionsJsonUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Firefox version info: $VersionsJsonUrl" }

        $json = ConvertFrom-Json $jsonText
        $version = $json.LATEST_FIREFOX_VERSION
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "LATEST_FIREFOX_VERSION field was empty."
        }

        Write-Log "Latest Firefox version       : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Firefox version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageFirefox {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Mozilla Firefox (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestFirefoxVersion
    if (-not $version) { throw "Could not resolve Firefox version." }

    $msiFileName = "Firefox Setup $version.msi"
    $downloadUrl = "$DownloadBase/$version/win64/en-US/" + ($msiFileName -replace ' ', '%20')

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $msiFileName"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # Install uses standard MSI wrapper
    $wrappers = New-MsiWrapperContent -MsiFileName $msiFileName

    # Firefox MSI is a thin wrapper around the EXE installer. msiexec /x
    # returns 1605 because the product isn't registered as an MSI install.
    # Use the native uninstaller (helper.exe /S) instead.
    $customUninstall = (
        '$helperPath = Join-Path $env:ProgramFiles ''Mozilla Firefox\uninstall\helper.exe''',
        'if (-not (Test-Path -LiteralPath $helperPath)) { exit 0 }',
        '$proc = Start-Process -FilePath $helperPath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrappers.Install `
        -UninstallPs1Content $customUninstall

    # --- Optional language variants (Language split) ---
    # Mozilla locale names differ from culture codes (de, fr, es-ES,
    # pt-BR); try the full culture first, then the primary subtag.
    function Resolve-FirefoxLocaleUrl {
        param([Parameter(Mandatory)][string]$Culture, [Parameter(Mandatory)][string]$FileName)
        $candidates = @($Culture)
        $primary = ($Culture -split '-')[0]
        if ($primary -ne $Culture) { $candidates += $primary }
        foreach ($locale in $candidates) {
            $url = "$DownloadBase/$version/win64/$locale/" + ($FileName -replace ' ', '%20')
            try {
                $resp = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -ErrorAction Stop
                if ($resp.StatusCode -eq 200) { return [pscustomobject]@{ Locale = $locale; Url = $url } }
            } catch { continue }
        }
        return $null
    }

    $deploymentTypes = $null
    $variants = Get-RequestedPackagerVariants
    if ($variants -and $variants.Split -eq 'Language') {
        Write-Log ""
        Write-Log ("Language split requested     : {0}" -f (@($variants.Languages) -join ', '))

        $msiWrappersForLang = New-MsiWrapperContent -MsiFileName $msiFileName
        $langEntries = @()
        foreach ($culture in @($variants.Languages)) {
            if ($culture -match '^en(-US)?$') { Write-Log "Skipping $culture (the en-US fallback already covers it)."; continue }
            $resolved = Resolve-FirefoxLocaleUrl -Culture $culture -FileName $msiFileName
            if (-not $resolved) { throw "No Firefox $version installer found for language '$culture' (tried the culture and its primary subtag)." }

            $langFolder = $resolved.Locale
            $localLangMsi = Join-Path $BaseDownloadRoot ("{0}-{1}" -f $resolved.Locale, $msiFileName)
            if (-not (Test-Path -LiteralPath $localLangMsi)) {
                Write-Log "Language download URL        : $($resolved.Url)"
                Invoke-DownloadWithRetry -Url $resolved.Url -OutFile $localLangMsi
            }

            $langContentPath = Join-Path $localContentPath $langFolder
            Initialize-Folder -Path $langContentPath
            Copy-Item -LiteralPath $localLangMsi -Destination (Join-Path $langContentPath $msiFileName) -Force -ErrorAction Stop
            Write-Log "Copied $langFolder MSI to staged : $langContentPath"

            Write-ContentWrappers -OutputPath $langContentPath `
                -InstallPs1Content $msiWrappersForLang.Install `
                -UninstallPs1Content $customUninstall

            $langEntries += @{
                NameSuffix     = $culture
                ContentSubpath = $langFolder
                Requirements   = @(@{ ConditionId = 'os-language'; Cultures = @($culture) })
            }
        }
        if ($langEntries.Count -eq 0) { throw "Language split requested but every language reduced to the en-US fallback." }

        # en-US payload at the content root is the unconditional fallback.
        $deploymentTypes = $langEntries + @(@{ NameSuffix = 'en-US' })
    }

    # --- Optional ARM64 variant (Architecture split) ---
    # Mozilla ships win64-aarch64 as an exe installer only. Both
    # architectures install to the same Program Files path, so the file
    # detection and the helper.exe uninstall carry over unchanged.
    if ($variants -and $variants.Split -eq 'Architecture') {
        Write-Log ""
        Write-Log "Architecture split requested : staging ARM64 variant"

        $armExeFileName = "Firefox Setup $version.exe"
        $armUrl = "$DownloadBase/$version/win64-aarch64/en-US/" + ($armExeFileName -replace ' ', '%20')
        $localArmExe = Join-Path $BaseDownloadRoot $armExeFileName
        if (-not (Test-Path -LiteralPath $localArmExe)) {
            Write-Log "ARM64 download URL           : $armUrl"
            Invoke-DownloadWithRetry -Url $armUrl -OutFile $localArmExe
        }

        $armContentPath = Join-Path $localContentPath "arm64"
        Initialize-Folder -Path $armContentPath
        Copy-Item -LiteralPath $localArmExe -Destination (Join-Path $armContentPath $armExeFileName) -Force -ErrorAction Stop
        Write-Log "Copied ARM64 exe to staged   : $armContentPath"

        $armWrappers = New-ExeWrapperContent -InstallerFileName $armExeFileName -InstallArgs "'/S'" `
            -UninstallCommand 'unused'
        Write-ContentWrappers -OutputPath $armContentPath `
            -InstallPs1Content $armWrappers.Install `
            -UninstallPs1Content $customUninstall

        # Both entries are gated: a machine that is neither x64 nor ARM64
        # should install neither variant. Detection comes from the base
        # manifest entry; the install path is identical on both.
        $deploymentTypes = @(
            @{
                NameSuffix     = 'ARM64'
                ContentSubpath = 'arm64'
                Requirements   = @(@{ ConditionId = 'cpu-arch'; Value = 'ARM64' })
            },
            @{
                NameSuffix   = 'x64'
                Requirements = @(@{ ConditionId = 'cpu-arch'; Value = 'x64' })
            }
        )
    }

    # --- Write stage manifest ---
    $detectionPath = "{0}\Mozilla Firefox" -f $env:ProgramFiles

    # A split app's name drops only the marker its deployment types now
    # cover: the architecture split stays en-US, the language split stays
    # x64-only.
    $appName = if ($variants -and $variants.Split -eq 'Architecture') { "Mozilla Firefox (en-US)" }
               elseif ($variants -and $variants.Split -eq 'Language') { "Mozilla Firefox (x64)" }
               else { "Mozilla Firefox (x64 en-US)" }
    $publisher = "Mozilla"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : firefox.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        RunningProcess  = @("firefox")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "firefox.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
        }
    }
    if ($deploymentTypes) { $manifestData['DeploymentTypes'] = $deploymentTypes }
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageFirefox {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Mozilla Firefox (x64) - PACKAGE phase"
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
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
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
        $v = Get-LatestFirefoxVersion -Quiet
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
    Write-Log "Mozilla Firefox (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionsJsonUrl              : $VersionsJsonUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageFirefox
    }
    elseif ($PackageOnly) {
        Invoke-PackageFirefox
    }
    else {
        Invoke-StageFirefox
        Invoke-PackageFirefox
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-firefox'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
