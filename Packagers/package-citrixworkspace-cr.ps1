<#
Vendor: Cloud Software Group
App: Citrix Workspace app for Windows (Current Release)
CMName: Citrix Workspace CR
VendorUrl: https://www.citrix.com/downloads/workspace-app/
CPE: cpe:2.3:a:citrix:workspace_app:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.citrix.com/en-us/citrix-workspace-app-for-windows/whats-new.html
DownloadPageUrl: https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html
IconSource: Installer
UpdateCadenceDays: 60

.SYNOPSIS
    Packages Citrix Workspace app for Windows, Current Release (x64), for ConfigMgr.

.DESCRIPTION
    Reads the Current stream version from the vendor's update catalog, downloads
    the matching CitrixWorkspaceApp.exe, stages content to a versioned local
    folder, and creates a ConfigMgr Application with registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    Scope: Current Release only. The LTSR build is served from an account-gated
    download page and cannot be fetched unattended, so this packager does not
    attempt it.

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
    Content is staged under: <FileServerPath>\Applications\Citrix\Citrix Workspace CR\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\CitrixWorkspaceCR).
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
    create ConfigMgr application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Current Release version string and exits.

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
$CatalogUrl      = "https://downloadplugins.citrix.com/ReceiverUpdates/Prod/catalog_win.xml"
$CatalogPattern  = '<Stream>Current</Stream><Version>([\d.]+)</Version>'
$ExeDownloadUrl  = "https://downloadplugins.citrix.com/Windows/CitrixWorkspaceApp.exe"
$CacheFileName   = "CitrixWorkspaceApp.exe"

$VendorFolder = "Citrix"
$AppFolder    = "Citrix Workspace CR"

$BaseDownloadRoot = Join-Path $DownloadRoot "CitrixWorkspaceCR"
$SwitchesFile     = Join-Path $PSScriptRoot "citrix-workspace-switches.json"
$CwaStream        = 'Current'

# The installer writes its ARP entry to the 32-bit registry view; ConfigMgr adds
# the WOW6432Node segment itself when the clause is not marked 64-bit.
$DetectionRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CitrixOnlinePluginPackWeb"

# --- Functions ---


function Get-CitrixWorkspaceCurrentVersion {
    param([switch]$Quiet)

    Write-Log "Citrix update catalog        : $CatalogUrl" -Quiet:$Quiet

    $xml = (curl.exe -L --fail --silent --show-error $CatalogUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch the Citrix update catalog: $CatalogUrl" }

    $m = [regex]::Match($xml, $CatalogPattern)
    if (-not $m.Success) {
        throw "Could not find a Current stream version in the Citrix update catalog."
    }

    $version = $m.Groups[1].Value
    Write-Log "Latest Current Release       : $version" -Quiet:$Quiet
    return $version
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

function Invoke-StageCitrixWorkspaceCR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace CR (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Version ---
    $version = Get-CitrixWorkspaceCurrentVersion
    $installerFileName = "CitrixWorkspaceApp-CR-$version.exe"

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $CacheFileName
    Write-Log "Download URL                 : $ExeDownloadUrl"
    Write-Log "Local EXE path               : $localExe"
    Write-Log ""
    Write-Log "Downloading Citrix Workspace app installer..."
    Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe

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
    Write-Log "Detection key                : HKLM\$DetectionRegistryKey (32-bit view)"
    Write-Log "Detection value              : DisplayVersion >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Citrix Workspace CR $version"
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
            Is64Bit             = $false
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

function Invoke-PackageCitrixWorkspaceCR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace CR (x64) - PACKAGE phase"
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
        Write-Output (Get-CitrixWorkspaceCurrentVersion -Quiet)
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Citrix Workspace CR GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace CR (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CatalogUrl                   : $CatalogUrl"
    Write-Log "ExeDownloadUrl               : $ExeDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCitrixWorkspaceCR
    }
    elseif ($PackageOnly) {
        Invoke-PackageCitrixWorkspaceCR
    }
    else {
        Invoke-StageCitrixWorkspaceCR
        Invoke-PackageCitrixWorkspaceCR
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-citrixworkspace-cr'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
