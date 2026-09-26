<#
Vendor: Cloud Software Group
App: Citrix Workspace app for Windows (LTSR x86)
CMName: Citrix Workspace LTSR x86
VendorUrl: https://www.citrix.com/downloads/workspace-app/
CPE: cpe:2.3:a:citrix:workspace_app:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.citrix.com/en-us/citrix-workspace-app-for-windows/whats-new.html
DownloadPageUrl: https://www.citrix.com/downloads/workspace-app/workspace-app-for-windows-long-term-service-release/workspace-app-for-windows-LTSR-Latest1.html
IconSource: Installer
UpdateCadenceDays: 90
LocalSource: Optional

.SYNOPSIS
    Packages Citrix Workspace app for Windows, LTSR (x86), for ConfigMgr.

.DESCRIPTION
    Reads the latest LTSR version from the vendor's update catalog and stages
    the highest-version online (CitrixWorkspaceApp) or offline
    (CitrixWorkspaceFullInstaller) installer for this architecture from the
    local source folder. The catalog serves only the x86 build, so the x86
    packager downloads it, verified against the catalog SHA-256, when no
    folder is set; the x64 and ARM64 builds are behind the LTSR download page
    sign-in and always come from the folder. Creates a ConfigMgr Application with
    registry-based detection.

    Supports two-phase operation:
      -StageOnly    Resolve the installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    Install switches come from Packagers\citrix-workspace-switches.json, the
    same file the GUI edits. Keys with no documented installer switch are
    logged as ignored rather than guessed at. When Components.Customize is
    false no ADDLOCAL is emitted and the installer picks its own component set.

    GetLatestVersionOnly reads the catalog XML (a few hundred KB) and exits
    without downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Citrix\Citrix Workspace LTSR x86\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\CitrixWorkspaceLTSRx86).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER SourceFolder
    Folder that holds the online or offline installer. Overrides the folder
    chosen in Options > Packager Preferences > Local Installer Sources.

.PARAMETER StageOnly
    Runs only the Stage phase: resolve the installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available LTSR version string and exits.

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
$CatalogUrl      = "https://downloadplugins.citrix.com/ReceiverUpdates/Prod/catalog_win.xml"
# Catalog DownloadURL values are relative to the catalog's own folder.
$CatalogBaseUrl  = "https://downloadplugins.citrix.com/ReceiverUpdates/Prod"

$PackagerName     = "package-citrixworkspace-ltsr-x86"
$Architecture     = "x86"
# Online and offline installers share the pattern; the other builds carry an
# architecture suffix and a browser may add " (1)" to repeated names.
$InstallerFilter  = "CitrixWorkspace*.exe"
$InstallerExclude = @('*_x64*', '*_ARM64*')
$CatalogDownload  = $true

$VendorFolder = "Citrix"
$AppFolder    = "Citrix Workspace LTSR x86"

$BaseDownloadRoot = Join-Path $DownloadRoot "CitrixWorkspaceLTSRx86"
$SwitchesFile     = Join-Path $PSScriptRoot "citrix-workspace-switches.json"
$CwaStream        = 'LTSR'

# The x86 build writes its ARP entry to the 32-bit registry view and the x64
# and ARM64 builds to the 64-bit view; ConfigMgr adds the WOW6432Node segment
# itself when the clause is not marked 64-bit.
$DetectionRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CitrixOnlinePluginPackWeb"

# --- Functions ---


function Get-CitrixWorkspaceLtsrFromCatalog {
    <#
    .SYNOPSIS
        Returns the newest LTSR installer entry from the update catalog XML.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Xml)

    $catalog = [xml]$Xml
    $entries = @($catalog.Catalog.Installers.Installer | Where-Object { [string]$_.Stream -eq 'LTSR' })
    $parsed = foreach ($entry in $entries) {
        $v = $null
        if ([version]::TryParse([string]$entry.Version, [ref]$v) -and [string]$entry.DownloadURL -and [string]$entry.Hash -match '^[0-9A-Fa-f]{64}$') {
            [pscustomobject]@{ Parsed = $v; Entry = $entry }
        }
    }
    $newest = @($parsed | Sort-Object Parsed -Descending | Select-Object -First 1)
    if ($newest.Count -eq 0) { return $null }

    $entry = $newest[0].Entry
    return [pscustomobject]@{
        Version     = [string]$entry.Version
        DownloadUrl = $CatalogBaseUrl + '/' + ([string]$entry.DownloadURL).TrimStart('/')
        Sha256      = ([string]$entry.Hash).ToLowerInvariant()
    }
}


function Get-CitrixWorkspaceLtsrRelease {
    param([switch]$Quiet)

    Write-Log "Citrix update catalog        : $CatalogUrl" -Quiet:$Quiet

    $xml = (curl.exe -L --fail --silent --show-error $CatalogUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Citrix update catalog: $CatalogUrl" }

    $release = Get-CitrixWorkspaceLtsrFromCatalog -Xml $xml
    if (-not $release) {
        throw "Could not find an LTSR installer with a version, download path and SHA-256 in the Citrix update catalog."
    }

    Write-Log "Latest LTSR                  : $($release.Version)" -Quiet:$Quiet
    return $release
}


function ConvertTo-CwaInstallArguments {
    <#
    .SYNOPSIS
        Builds the CitrixWorkspaceApp.exe argument list from a parsed
        citrix-workspace-switches.json object.

    .DESCRIPTION
        Returns an object with Arguments (ordered string array) and Warnings
        (settings skipped because their value is invalid or the vendor install
        page for the stream documents no switch for them). A missing or empty
        setting emits no switch, so the installer default applies.
    #>
    param(
        [AllowNull()][object]$Config,
        [Parameter(Mandatory)][ValidateSet('Current','LTSR')][string]$Stream
    )

    $arguments = [System.Collections.Generic.List[string]]::new()
    $warnings  = [System.Collections.Generic.List[string]]::new()
    $arguments.Add('/silent')

    if ($null -ne $Config) {
        $install = $Config.Installation
        if ($null -ne $install) {
            if ([bool]$install.CleanInstall) { $arguments.Add('/CleanInstall') }
            # ENABLE_SSON takes effect only together with /includeSSON.
            if ([bool]$install.IncludeSSON) {
                $arguments.Add('/includeSSON')
                if ($null -ne $install.EnableSSON) {
                    $arguments.Add('ENABLE_SSON=' + $(if ([bool]$install.EnableSSON) { 'Yes' } else { 'No' }))
                }
            }
            if ([bool]$install.AppProtection) { $arguments.Add('startAppProtection') }
            if ($null -ne $install.SessionPreLaunch) {
                $arguments.Add('ENABLEPRELAUNCH=' + $(if ([bool]$install.SessionPreLaunch) { 'True' } else { 'False' }))
            }
            if ($null -ne $install.SelfServiceMode) {
                $arguments.Add('SELFSERVICEMODE=' + $(if ([bool]$install.SelfServiceMode) { 'True' } else { 'False' }))
            }
        }

        $plugins = $Config.Plugins
        if ($null -ne $plugins) {
            $addons = [System.Collections.Generic.List[string]]::new()
            if ($Stream -eq 'Current') {
                if ($null -ne $plugins.MSTeamsPlugin) {
                    $arguments.Add('/InstallMSTeamsPlugin=' + $(if ([bool]$plugins.MSTeamsPlugin) { 'Y' } else { 'N' }))
                }
                if ($null -ne $plugins.ZoomPlugin -and -not [bool]$plugins.ZoomPlugin) { $arguments.Add('Installzoomplugin=N') }
            }
            else {
                if ($null -ne $plugins.MSTeamsPlugin) {
                    $warnings.Add('Plugins.MSTeamsPlugin has no documented LTSR installer switch; ignored.')
                }
                if ([bool]$plugins.ZoomPlugin) { $addons.Add('ZoomVDIPlugin') }
            }
            if ($null -ne $plugins.EPAClient -and -not [bool]$plugins.EPAClient) { $arguments.Add('InstallEPAClient=N') }
            if ([bool]$plugins.WebExPlugin) { $addons.Add('WebexVDIPlugin') }
            if ($addons.Count -gt 0) { $arguments.Add('ADDONS=' + ($addons -join ',')) }

            if ([bool]$plugins.UberAgent) {
                $arguments.Add('/InstallUberAgent')
                if ([bool]$plugins.UberAgentSkipUpgrade) { $arguments.Add('/SkipUberAgentUpgrade') }
            }
            elseif ([bool]$plugins.UberAgentSkipUpgrade) {
                $warnings.Add('Plugins.UberAgentSkipUpgrade applies only with Plugins.UberAgent; ignored.')
            }

            if ([bool]$plugins.SessionRecording) {
                if ($Stream -eq 'Current') { $arguments.Add('/InstallSRAgent') }
                else { $warnings.Add('Plugins.SessionRecording has no documented LTSR installer switch; ignored.') }
            }
        }

        $update = $Config.UpdateAndTelemetry
        if ($null -ne $update) {
            $autoUpdate = ([string]$update.AutoUpdateCheck).Trim().ToLowerInvariant()
            if ($autoUpdate) {
                if ($autoUpdate -in @('auto','manual','disabled')) { $arguments.Add('AutoUpdateCheck=' + $autoUpdate) }
                else { $warnings.Add(("UpdateAndTelemetry.AutoUpdateCheck '{0}' is not auto, manual or disabled; ignored." -f $update.AutoUpdateCheck)) }
            }
            if ($null -ne $update.EnableCEIP) {
                $arguments.Add('EnableCEIP=' + $(if ([bool]$update.EnableCEIP) { 'True' } else { 'False' }))
            }
            if ($null -ne $update.EnableTracing) {
                $arguments.Add('EnableTracing=' + $(if ([bool]$update.EnableTracing) { 'true' } else { 'false' }))
            }
        }

        $policy = $Config.StorePolicy
        if ($null -ne $policy) {
            foreach ($pair in @(@('AllowAddStore','ALLOWADDSTORE'), @('AllowSavePwd','ALLOWSAVEPWD'))) {
                $value = ([string]$policy.($pair[0])).Trim().ToUpperInvariant()
                if (-not $value) { continue }
                if ($value -in @('S','A','N')) { $arguments.Add($pair[1] + '=' + $value) }
                else { $warnings.Add(("StorePolicy.{0} '{1}' is not S, A or N; ignored." -f $pair[0], $policy.($pair[0]))) }
            }
        }

        $store = $Config.Store
        if ($null -ne $store) {
            $storeName = ([string]$store.Name).Trim()
            $storeUrl  = ([string]$store.Url).Trim()
            if ($storeName -or $storeUrl) {
                $parsedUrl = $null
                if (-not $storeName -or -not $storeUrl) {
                    $warnings.Add('Store needs both Name and Url; STORE0 omitted.')
                }
                elseif ($storeName -match '[;"]') {
                    $warnings.Add('Store.Name contains a semicolon or double quote; STORE0 omitted.')
                }
                elseif (-not [uri]::TryCreate($storeUrl, [System.UriKind]::Absolute, [ref]$parsedUrl) -or $parsedUrl.Scheme -notin @('https','http') -or $storeUrl -match '[;"\s]') {
                    $warnings.Add(("Store.Url '{0}' is not an absolute http or https URL; STORE0 omitted." -f $storeUrl))
                }
                else {
                    $storeValue = '{0};{1};On;{0}' -f $storeName, $storeUrl
                    # Start-Process joins ArgumentList elements with spaces and adds no quotes.
                    if ($storeValue -match '\s') { $storeValue = '"' + $storeValue + '"' }
                    $arguments.Add('STORE0=' + $storeValue)
                }
            }
        }

        # ADDLOCAL restricts the install to the listed components, so it is
        # emitted only when the operator opted in.
        if ($null -ne $Config.Components -and [bool]$Config.Components.Customize) {
            $componentNames = @('ReceiverInside','ICA_Client','AM','SelfService','DesktopViewer','WebHelper','BCR_Client','USB','SSON')
            $selected = @($componentNames | Where-Object { [bool]$Config.Components.$_ })
            if ($selected.Count -gt 0) { $arguments.Add('ADDLOCAL=' + ($selected -join ',')) }
            else { $warnings.Add('Components.Customize is set but no component is enabled; ADDLOCAL omitted.') }
        }
    }

    return [pscustomobject]@{
        Arguments = $arguments.ToArray()
        Warnings  = $warnings.ToArray()
    }
}


function Get-CwaInstallArguments {
    <#
    .SYNOPSIS
        Reads citrix-workspace-switches.json and returns the installer
        arguments for this packager's stream. Missing file or unreadable JSON
        yields /silent only.
    #>
    $cfg = $null
    if (-not (Test-Path -LiteralPath $SwitchesFile)) {
        Write-Log "Switch config not found      : $SwitchesFile (using the silent baseline)" -Level WARN
    }
    else {
        try {
            $cfg = Get-Content -LiteralPath $SwitchesFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            Write-Log "Switch config                : $SwitchesFile"
        }
        catch {
            Write-Log "Switch config unreadable     : $($_.Exception.Message) (using the silent baseline)" -Level WARN
        }
    }

    $result = ConvertTo-CwaInstallArguments -Config $cfg -Stream $CwaStream
    foreach ($warning in $result.Warnings) { Write-Log $warning -Level WARN }
    return $result.Arguments
}


function ConvertTo-WrapperArgumentList {
    <#
    .SYNOPSIS
        Renders an argument array as a PowerShell element list for the
        generated wrapper's @(...) ArgumentList.
    #>
    param([Parameter(Mandatory)][string[]]$Arguments)

    return (($Arguments | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', ')
}


function Get-CitrixWorkspaceUninstallContent {
    param([Parameter(Mandatory)][string]$InstallerFileName)

    $escaped = $InstallerFileName -replace "'", "''"
    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $escaped),
        'if (-not (Test-Path -LiteralPath $exePath)) { exit 1 }',
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/silent'', ''/uninstall'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCitrixWorkspaceLTSR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace LTSR ($Architecture) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Version and installer ---
    $release = Get-CitrixWorkspaceLtsrRelease
    $sourceFolderPath = Get-LocalSourceFolder -PackagerName $PackagerName -Override $SourceFolder

    if ($sourceFolderPath -or -not $CatalogDownload) {
        $localExe = Resolve-LocalSourceInstaller -PackagerName $PackagerName -Filter $InstallerFilter -Exclude $InstallerExclude -Override $SourceFolder
        Write-Log "Local installer              : $localExe"

        $signature = Get-AuthenticodeSignature -LiteralPath $localExe
        if ($signature.Status -ne 'Valid') {
            throw "Installer signature is $($signature.Status), not Valid: $localExe"
        }
        Write-Log "Installer signer             : $($signature.SignerCertificate.Subject)"

        $version = ([string][System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion).Trim()
        $parsed = $null
        if (-not [version]::TryParse($version, [ref]$parsed)) {
            throw "Installer carries no numeric file version: $localExe"
        }
        if ($parsed -lt [version]$release.Version) {
            Write-Log "Local installer $version is older than the catalog LTSR $($release.Version); staging it as found." -Level WARN
        }
        $kind = if ((Split-Path -Leaf $localExe) -like '*FullInstaller*') { 'CitrixWorkspaceFullInstaller' } else { 'CitrixWorkspaceApp' }
        $installerFileName = "$kind-LTSR-$Architecture-$version.exe"
    }
    else {
        $version = $release.Version
        $installerFileName = "CitrixWorkspaceApp-LTSR-$Architecture-$version.exe"
        $localExe = Join-Path $BaseDownloadRoot $installerFileName
        Write-Log "Download URL                 : $($release.DownloadUrl)"
        $cachedHash = ''
        if (Test-Path -LiteralPath $localExe) {
            $cachedHash = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256).Hash.ToLowerInvariant()
        }
        if ($cachedHash -ne $release.Sha256) {
            Write-Log "Downloading Citrix Workspace app installer..."
            Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localExe
            $actualHash = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actualHash -ne $release.Sha256) {
                throw "Downloaded installer SHA-256 $actualHash does not match the catalog hash $($release.Sha256)."
            }
        }
        else {
            Write-Log "Local installer matches the catalog hash. Skipping download."
        }
        Write-Log "Installer SHA-256            : $($release.Sha256) (verified)"
    }

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

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

    # --- Install arguments from the GUI-managed switch file ---
    $installArguments = Get-CwaInstallArguments
    Write-Log ""
    Write-Log ("Install arguments            : {0}" -f ($installArguments -join ' '))
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs (ConvertTo-WrapperArgumentList -Arguments $installArguments) `
        -UninstallCommand 'unused'

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content (Get-CitrixWorkspaceUninstallContent -InstallerFileName $installerFileName)

    # --- Write stage manifest ---
    Write-Log "Detection key                : HKLM\$DetectionRegistryKey ($(if ($Architecture -eq 'x86') { '32-bit' } else { '64-bit' }) view)"
    Write-Log "Detection value              : DisplayVersion >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Citrix Workspace LTSR $Architecture $version"
        Publisher       = "Cloud Software Group"
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = ($installArguments -join ' ')
        UninstallArgs   = "/silent /uninstall"
        RunningProcess  = @("SelfService", "AuthManSvr", "Receiver")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $DetectionRegistryKey
            ValueName           = "DisplayVersion"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
            Is64Bit             = ($Architecture -ne 'x86')
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

function Invoke-PackageCitrixWorkspaceLTSR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace LTSR ($Architecture) - PACKAGE phase"
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
        Write-Output (Get-CitrixWorkspaceLtsrRelease -Quiet).Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Citrix Workspace LTSR GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace LTSR ($Architecture) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CatalogUrl                   : $CatalogUrl"
    Write-Log "CatalogUrl                   : $CatalogUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCitrixWorkspaceLTSR
    }
    elseif ($PackageOnly) {
        Invoke-PackageCitrixWorkspaceLTSR
    }
    else {
        Invoke-StageCitrixWorkspaceLTSR
        Invoke-PackageCitrixWorkspaceLTSR
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-citrixworkspace-ltsr'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
