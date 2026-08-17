<#
.SYNOPSIS
    Shared core module for the MECM tool suite.

.DESCRIPTION
    Vendored into each consumer tool at Lib\SuiteCommon\ by
    sync-suitecommon.ps1. Provides:
      - Structured logging (Initialize-Logging, Write-Log, Write-LogErrorRecord)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite,
        Test-CMConnection, Get-CMConnectionInfo)
      - Flat JSON settings persistence (Read-SuiteSettings, Save-SuiteSettings)

    Consumers never edit the vendored copy; changes flow through the
    suite-core repository and re-sync.

.EXAMPLE
    # Top of a consumer tool module (Module\<Tool>Common.psm1):
    if (-not (Get-Module SuiteCommon)) {
        Import-Module (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
    }
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__SuiteLogPath         = $null
$script:__SuiteVerbose         = $false
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null
$script:ConnectedAt            = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    <#
    .SYNOPSIS
        Sets the log file path and verbose behavior for Write-Log.

    .DESCRIPTION
        -Attach appends to an existing file instead of writing a fresh
        header (used by background runspaces that share the shell's log).
        Verbose diagnostics: the explicit switch wins; otherwise the
        SUITE_VERBOSE environment variable enables DEBUG console output
        unless it holds an explicit off value (0/false/no).
    #>
    param(
        [string]$LogPath,
        [switch]$Attach,
        [switch]$VerboseLogging
    )

    # A whitespace-only path reaches Set-Content as a filename Windows strips
    # to nothing and throws DirectoryNotFoundException; treat it as no path.
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { $LogPath = $LogPath.Trim() } else { $LogPath = $null }

    $script:__SuiteLogPath = $LogPath

    $envVerbose = $env:SUITE_VERBOSE
    $script:__SuiteVerbose = [bool]$VerboseLogging -or
        (-not [string]::IsNullOrWhiteSpace($envVerbose) -and $envVerbose -notin @('0', 'false', 'no'))

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        if (-not $Attach) {
            $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
        }
    }

    if ($script:__SuiteVerbose) {
        Write-Log "Verbose logging enabled (DEBUG lines on console and in log file)." -Level DEBUG
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.

    .DESCRIPTION
        INFO  -> Write-Host (stdout)
        WARN  -> Write-Host (stdout)
        ERROR -> Write-Host (stdout) + $host.UI.WriteErrorLine (stderr)
        DEBUG -> log file always; console only when verbose logging is
                 enabled (Initialize-Logging -VerboseLogging or the
                 SUITE_VERBOSE environment variable).

        -Quiet suppresses all console output but still writes to the log file.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write-Log is the single console-surfacing path; Write-Host is the deliberate contract so lines reach both the host and the file log.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification='Write-Log is the suite-wide logging contract on Windows PowerShell 5.1, which ships no Write-Log.')]
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    $suppressConsole = $Quiet -or ($Level -eq 'DEBUG' -and -not $script:__SuiteVerbose)

    if (-not $suppressConsole) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__SuiteLogPath) {
        # Concurrent writers (a background runspace sharing the shell's log
        # via -Attach) collide on the per-line open/close and the sharing
        # violation silently drops the line - measured at 35-86% loss under
        # contention. A named mutex keyed on the path serializes every
        # SuiteCommon writer in the session; the retry loop still covers
        # non-SuiteCommon writers holding the file.
        $mutexName = 'SuiteCommonLog_' + ($script:__SuiteLogPath.ToLowerInvariant() -replace '[\\/:]', '_')
        $mutex = $null
        $owned = $false
        try {
            $mutex = New-Object System.Threading.Mutex($false, $mutexName)
            try { $owned = $mutex.WaitOne(2000) }
            catch [System.Threading.AbandonedMutexException] { $owned = $true }
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Add-Content -LiteralPath $script:__SuiteLogPath -Value $formatted -Encoding UTF8 -ErrorAction Stop
                    break
                }
                catch {
                    if ($attempt -lt 3) { Start-Sleep -Milliseconds 5 }
                }
            }
        }
        catch { $null = $_ }
        finally {
            if ($mutex) {
                if ($owned) { try { $mutex.ReleaseMutex() } catch { $null = $_ } }
                $mutex.Dispose()
            }
        }
    }
}

function Write-LogErrorRecord {
    <#
    .SYNOPSIS
        Logs full diagnostics for an ErrorRecord: exception chain, error id,
        failing file/line/statement, and the script stack trace.

    .DESCRIPTION
        Designed for catch blocks. The InvocationInfo position and
        ScriptStackTrace identify the exact line that threw, which a bare
        $_.Exception.Message never reveals.

    .PARAMETER Level
        Log level for the diagnostic lines. Default ERROR. Pass DEBUG when
        the caller already emits its own single ERROR summary line and the
        detail should only land in the log file.
    #>
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [string]$Context = '',

        [ValidateSet('ERROR', 'WARN', 'DEBUG')]
        [string]$Level = 'ERROR'
    )

    if ($Context) {
        Write-Log ("Failure context              : {0}" -f $Context) -Level $Level
    }

    $ex = $ErrorRecord.Exception
    if ($ex) {
        Write-Log ("Exception                    : {0}: {1}" -f $ex.GetType().FullName, $ex.Message) -Level $Level
        $inner = $ex.InnerException
        while ($inner) {
            Write-Log ("Inner exception              : {0}: {1}" -f $inner.GetType().FullName, $inner.Message) -Level $Level
            $inner = $inner.InnerException
        }
    }

    Write-Log ("FullyQualifiedErrorId        : {0}" -f $ErrorRecord.FullyQualifiedErrorId) -Level $Level

    $ii = $ErrorRecord.InvocationInfo
    if ($ii) {
        if ($ii.ScriptName) {
            Write-Log ("Failing location             : {0}:{1}" -f $ii.ScriptName, $ii.ScriptLineNumber) -Level $Level
        }
        if ($ii.Line) {
            Write-Log ("Failing statement            : {0}" -f $ii.Line.Trim()) -Level $Level
        }
    }

    if ($ErrorRecord.ScriptStackTrace) {
        foreach ($frame in ($ErrorRecord.ScriptStackTrace -split "`r?`n")) {
            Write-Log ("Stack                        : {0}" -f $frame) -Level $Level
        }
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Resolve-ConfigurationManagerModulePath {
    <#
    .SYNOPSIS
        Locates ConfigurationManager.psd1 from the console install.

    .DESCRIPTION
        SMS_ADMIN_UI_PATH points at ...\AdminConsole\bin\i386; the module
        manifest lives at ...\AdminConsole\bin\ConfigurationManager.psd1.
        Split-Path gives a normalized parent: a literal '..' in the module
        path breaks the CM manifest's relative resolution of its nested
        submodules, and the first Get-CM* autoload then fails with
        "module could not be loaded".
    #>
    $candidates = @()

    if ($env:SMS_ADMIN_UI_PATH) {
        $candidates += (Join-Path (Split-Path -Parent $env:SMS_ADMIN_UI_PATH) 'ConfigurationManager.psd1')
    }

    $candidates += @(
        'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1',
        'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1',
        'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1',
        'C:\Program Files\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1'
    )

    foreach ($candidate in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }

    return $null
}

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, ensures the site PSDrive,
        and sets location onto it.

    .DESCRIPTION
        The PSDrive is created with -Scope Global: a drive created inside
        this module's session state is invisible to callers and to sibling
        tool modules resolving the CMSite location (current location is
        runspace-shared; drive visibility is session-state-scoped). An
        existing drive that fails Set-Location (provider restart) is torn
        down and rebuilt from the provider. Saves the original location for
        Disconnect-CMSite. Returns $true on success, $false on failure.

    .PARAMETER SMSProvider
        SMS Provider machine. Optional: falls back to the SUITE_CM_PROVIDER
        environment variable, then to an existing drive's root.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Alias('ProviderMachineName')][string]$SMSProvider = '',
        [string]$CMModulePath = '',
        [switch]$SkipSiteVerification
    )

    # Capture only on a fresh connect: a reconnect/site-switch without an
    # intervening Disconnect-CMSite would otherwise overwrite the saved
    # location with the previous site's drive and lose the real origin.
    if (-not $script:ConnectedSiteCode) {
        $script:OriginalLocation = Get-Location
    }

    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $resolved = if (-not [string]::IsNullOrWhiteSpace($CMModulePath)) { $CMModulePath } else { Resolve-ConfigurationManagerModulePath }

        if (-not $resolved -or -not (Test-Path -LiteralPath $resolved)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            # -Global: this import runs in SuiteCommon's session state; without
            # it the CM cmdlets are visible only inside SuiteCommon, and every
            # consumer tool module's own Get-CM* calls fail with "term not
            # recognized" - the same visibility rule that forces -Scope Global
            # on the PSDrive below.
            Import-Module $resolved -Force -DisableNameChecking -Global -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module from $resolved"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    $provider = $SMSProvider
    if ([string]::IsNullOrWhiteSpace($provider) -and -not [string]::IsNullOrWhiteSpace($env:SUITE_CM_PROVIDER)) {
        $provider = $env:SUITE_CM_PROVIDER.Trim()
    }

    # The CM module auto-creates the site PSDrive on import when the install
    # baked in DefaultSiteServerName. Only fall back to New-PSDrive when the
    # drive truly is not registered yet -- creating it while the CM module
    # is still half-initialized races the module's own provider setup and
    # surfaces as the Get-CM* autoload failure.
    $siteDrive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue

    if ($siteDrive -and -not [string]::IsNullOrWhiteSpace($provider) -and ([string]$siteDrive.Root -ne $provider)) {
        # Same site code bound to a different SMS Provider (the provider was
        # changed in a tool's options): rebind to the requested one. Stepping
        # off the drive first is mandatory - the drive you are located in
        # cannot be removed.
        Write-Log "PSDrive ${SiteCode}: points at '$($siteDrive.Root)', not '$provider' -- recreating." -Level WARN
        try { Set-Location -LiteralPath "$env:SystemDrive\" -ErrorAction SilentlyContinue } catch { $null = $_ }
        Remove-PSDrive -Name $SiteCode -PSProvider CMSite -Force -ErrorAction SilentlyContinue
        $siteDrive = $null
    }

    if ($siteDrive) {
        try {
            Set-Location "${SiteCode}:" -ErrorAction Stop
        }
        catch {
            # Entering an existing drive fails when its provider connection
            # is gone (e.g. provider restart). Tear it down and rebuild.
            Write-Log ("Set-Location {0}: failed on existing drive ({1}); recreating it." -f $SiteCode, $_.Exception.Message) -Level WARN
            if ([string]::IsNullOrWhiteSpace($provider)) { $provider = [string]$siteDrive.Root }
            try { Set-Location -LiteralPath "$env:SystemDrive\" -ErrorAction SilentlyContinue } catch { $null = $_ }
            Remove-PSDrive -Name $SiteCode -Force -ErrorAction SilentlyContinue
            $siteDrive = $null
        }
    }

    if (-not $siteDrive) {
        if ([string]::IsNullOrWhiteSpace($provider)) {
            Write-Log "No CMSite drive '$SiteCode' exists and no SMS provider machine was given (parameter or SUITE_CM_PROVIDER)." -Level ERROR
            return $false
        }

        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $provider -Scope Global -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode -> $provider"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }

        try {
            Set-Location "${SiteCode}:" -ErrorAction Stop
        }
        catch {
            Write-Log "Failed to enter site drive $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    if (-not $SkipSiteVerification) {
        try {
            $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
            Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        }
        catch {
            Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
            return $false
        }
    }
    else {
        Write-Log "Connected to CM site: $SiteCode"
    }

    if ([string]::IsNullOrWhiteSpace($provider)) {
        $drive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue
        if ($drive) { $provider = [string]$drive.Root }
    }

    # Verbose-only sanity check: a drive named for the wrong site code
    # connects fine but every subsequent CM cmdlet dies with "Key cannot
    # be null. Parameter name: key". Surface the mismatch here so the log
    # explains the failure instead of the cmdlet's opaque message.
    if ($script:__SuiteVerbose -and -not [string]::IsNullOrWhiteSpace($provider)) {
        try {
            $check = Test-CMSiteCodeMatchesProvider -SiteCode $SiteCode -ProviderMachineName $provider
            if ($check -and -not $check.Match) {
                Write-Log ("Site code mismatch: drive is named '{0}' but provider {1} serves site code(s) {2}. CM cmdlets typically fail with 'Key cannot be null. Parameter name: key' in this state - correct the configured site code." -f $SiteCode, $provider, ($check.SiteCodes -join ', ')) -Level WARN
            }
            elseif ($check) {
                Write-Log ("Provider confirms site code  : {0} on {1}" -f $SiteCode, $provider) -Level DEBUG
            }
        } catch { $null = $_ }
    }

    $script:ConnectedSiteCode    = $SiteCode
    $script:ConnectedSMSProvider = $provider
    $script:ConnectedAt          = Get-Date
    return $true
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location saved by Connect-CMSite and clears
        the recorded connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { $null = $_ }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    $script:ConnectedAt          = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

function Get-CMConnectionInfo {
    <#
    .SYNOPSIS
        Returns the recorded connection (SiteCode, SMSProvider, ConnectedAt)
        or $null when not connected.

    .DESCRIPTION
        Consumers that need the connected site code or provider for their
        own CIM/WMI queries read it here instead of reaching into module
        state that a module split would hide from them.
    #>
    if (-not $script:ConnectedSiteCode) { return $null }

    return [pscustomobject]@{
        SiteCode    = $script:ConnectedSiteCode
        SMSProvider = $script:ConnectedSMSProvider
        ConnectedAt = $script:ConnectedAt
    }
}

function Test-CMSiteCodeMatchesProvider {
    <#
    .SYNOPSIS
        Asks the SMS Provider which site code(s) it actually serves.

    .DESCRIPTION
        Queries root\sms:SMS_ProviderLocation on the provider machine. A
        CMSite PSDrive whose NAME does not match a real site code on its
        Root server connects fine, but every subsequent CM cmdlet fails
        with "Key cannot be null. Parameter name: key". Returns $null when
        the query itself fails (offline provider, no WMI rights).
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$ProviderMachineName
    )

    try {
        $locations = @(Get-CimInstance -ComputerName $ProviderMachineName -Namespace 'root\sms' -ClassName 'SMS_ProviderLocation' -ErrorAction Stop)
        $codes = @($locations | ForEach-Object { [string]$_.SiteCode } | Where-Object { $_ } | Select-Object -Unique)
        if ($codes.Count -eq 0) { return $null }

        return [pscustomobject]@{
            Match     = ($codes -contains $SiteCode)
            SiteCodes = $codes
        }
    }
    catch {
        Write-Log ("Provider site-code query failed for {0}: {1}" -f $ProviderMachineName, $_.Exception.Message) -Level DEBUG
        return $null
    }
}

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

function Read-SuiteSettings {
    <#
    .SYNOPSIS
        Reads a flat JSON settings file over a defaults hashtable.

    .DESCRIPTION
        Returns a hashtable seeded from Defaults; any top-level key present
        in both the file and Defaults overwrites the default. Keys in the
        file that Defaults does not declare are ignored, so stale settings
        cannot smuggle unexpected state into a tool. A missing or
        unreadable file returns the defaults unchanged.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Reads the full settings hashtable by design.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Defaults
    )

    $settings = @{}
    foreach ($k in $Defaults.Keys) { $settings[$k] = $Defaults[$k] }

    if (Test-Path -LiteralPath $Path) {
        try {
            # -Encoding UTF8: on 5.1 the default decoder mangles non-ASCII in
            # BOM-less UTF-8 files (hand-edited settings) without any error.
            $loaded = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @($settings.Keys)) {
                $val = $loaded.$k
                if ($null -ne $val) { $settings[$k] = $val }
            }
        } catch { $null = $_ }
    }

    return $settings
}

function Save-SuiteSettings {
    <#
    .SYNOPSIS
        Writes a settings hashtable to a JSON file. Returns $true on
        success, $false on failure (failure is logged, never thrown -
        losing a preference write must not take the tool down).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full settings hashtable by design.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Settings,
        [int]$Depth = 5
    )

    try {
        $parentDir = Split-Path -Path $Path -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
        }
        # -ErrorAction Stop: Set-Content fails non-terminating by default,
        # which would skip the catch and report success on a failed write.
        $Settings | ConvertTo-Json -Depth $Depth | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log ("Failed to save settings to {0}: {1}" -f $Path, $_.Exception.Message) -Level WARN -Quiet
        return $false
    }
}

# ---------------------------------------------------------------------------
# Module identity
# ---------------------------------------------------------------------------

function Get-SuiteCommonVersion {
    <#
    .SYNOPSIS
        Returns this vendored copy's ModuleVersion string.
    #>
    $psd1 = Join-Path $PSScriptRoot 'SuiteCommon.psd1'
    try {
        return [string](Import-PowerShellDataFile -LiteralPath $psd1).ModuleVersion
    }
    catch {
        return ''
    }
}

# =============================================================================
# Window chrome: native title-bar drag + window-state persistence
# =============================================================================
# Hook delegates and their windows are keyed by HWND so multiple windows
# (main shell + modal dialogs) can hold hooks concurrently. The delegate
# and window references must stay rooted here for the window's lifetime;
# Remove-NativeTitleBarHitTestHook (wired via Add_Closed) is what prevents
# a per-window leak.
$script:TitleBarHitTestHooks   = @{}
$script:TitleBarHitTestWindows = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $h = [double]$Window.TitleBarHeight
        if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h }
    } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    # Template lookup can fail before the first layout pass. Keep the
    # right-side caption buttons available with a conservative fallback.
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch {
        # A silent failure here leaves the window undraggable with no trace;
        # the log line is the only signal an installation problem exists.
        try { Write-Log -Message ("Title-bar native hit-test hook failed: {0}" -f $_.Exception.Message) -Level WARN } catch { $null = $_ }
    }
}

function Remove-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($key -ne '0' -and $script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
            $script:TitleBarHitTestWindows.Remove($key)
            return
        }
        # By the time Closed fires the HWND is destroyed and Handle reads
        # zero, so the key lookup above cannot match. Without this
        # reference-based sweep every hooked window's delegate and window
        # object stay rooted for the process lifetime.
        $stale = @($script:TitleBarHitTestWindows.Keys | Where-Object {
            [object]::ReferenceEquals($script:TitleBarHitTestWindows[$_], $Window)
        })
        foreach ($k in $stale) {
            $script:TitleBarHitTestHooks.Remove($k)
            $script:TitleBarHitTestWindows.Remove($k)
        }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    # Without the Closed teardown every hooked window leaks its delegate and
    # a strong window reference for the process lifetime.
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

function Save-WindowState {
    param(
        [Parameter(Mandatory)]$Window,
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$ExtraState
    )
    try {
        $state = @{}
        # RestoreBounds, not Left/Width, when not Normal: a maximized close
        # would otherwise persist full-screen extents as normal geometry.
        if ($Window.WindowState -eq [System.Windows.WindowState]::Normal) {
            $raw = @([double]$Window.Left, [double]$Window.Top, [double]$Window.Width, [double]$Window.Height)
        }
        else {
            $rb = $Window.RestoreBounds
            $raw = @([double]$rb.Left, [double]$rb.Top, [double]$rb.Width, [double]$rb.Height)
        }
        # A never-shown window reads NaN (or an Empty rect's infinities);
        # casting those to [int] throws inside the caller's Closing handler.
        # Nothing meaningful exists to persist, so skip the save entirely.
        foreach ($v in $raw) {
            if ([double]::IsNaN($v) -or [double]::IsInfinity($v)) { return }
        }
        $state.Left   = [int]$raw[0]
        $state.Top    = [int]$raw[1]
        $state.Width  = [int]$raw[2]
        $state.Height = [int]$raw[3]
        $state.Maximized = ($Window.WindowState -eq [System.Windows.WindowState]::Maximized)
        if ($ExtraState) {
            foreach ($k in $ExtraState.Keys) { $state[$k] = $ExtraState[$k] }
        }
        Set-Content -LiteralPath $Path -Value ($state | ConvertTo-Json -Depth 5) -Encoding UTF8
    } catch { $null = $_ }
}

function Restore-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Reads the JSON state file and applies geometry; idempotent.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Restore is intentional and reads as a single action.')]
    param(
        [Parameter(Mandatory)]$Window,
        [Parameter(Mandatory)][string]$Path,
        [scriptblock]$OnStateLoaded
    )
    if (-not (Test-Path -LiteralPath $Path)) { return }
    try {
        $s = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -ErrorAction Stop

        # Schema bridge: WinForms-era files used X / Y; WPF uses Left / Top.
        # Legacy files would otherwise snap the window to (0,0) at MinSize.
        $left = if ($null -ne $s.Left) { [int]$s.Left } elseif ($null -ne $s.X) { [int]$s.X } else { $null }
        $top  = if ($null -ne $s.Top)  { [int]$s.Top  } elseif ($null -ne $s.Y) { [int]$s.Y } else { $null }
        $w    = if ($null -ne $s.Width)  { [int]$s.Width  } else { $null }
        $h    = if ($null -ne $s.Height) { [int]$s.Height } else { $null }

        if ($s.Maximized) {
            $Window.WindowState = [System.Windows.WindowState]::Maximized
        } elseif ($null -ne $left -and $null -ne $top -and $null -ne $w -and $null -ne $h) {
            # A saved position off every connected screen is clamped into the
            # nearest monitor's working area. Discarding the whole geometry
            # instead loses a still-valid window size whenever a monitor
            # detaches between sessions.
            Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue
            $screen = [System.Windows.Forms.Screen]::FromPoint([System.Drawing.Point]::new($left, $top))
            $bounds = $screen.WorkingArea
            $left = [Math]::Max($bounds.X, [Math]::Min($left, $bounds.Right - 200))
            $top  = [Math]::Max($bounds.Y, [Math]::Min($top,  $bounds.Bottom - 100))
            $Window.Left   = $left
            $Window.Top    = $top
            $Window.Width  = [Math]::Max($Window.MinWidth,  $w)
            $Window.Height = [Math]::Max($Window.MinHeight, $h)
        }

        if ($OnStateLoaded) { & $OnStateLoaded $s }
    } catch { $null = $_ }
}

# =============================================================================
# Theming: shared palette + title bar, sidebar, action button, dialog setters
# =============================================================================
$script:SuiteTheme = $null

function Initialize-SuiteTheme {
    <#
    .SYNOPSIS
        Stores the per-shell theming context the Update-*/Set-* theme
        functions read when their optional parameters are omitted.

    .DESCRIPTION
        IsDarkGetter and ActiveViewGetter are scriptblocks created in the
        shell's scope; they are invoked (not dot-sourced) so they evaluate
        against the shell's own state ($global:Prefs, a toggle control,
        $script:ActiveView) without this module naming those variables.
        Brushes defaults to the suite's standard palette; pass -Brushes to
        override individual entries only.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Stores module-scope theming context only.')]
    param(
        [Parameter(Mandatory)]$Window,
        [Parameter(Mandatory)][scriptblock]$IsDarkGetter,
        [scriptblock]$ActiveViewGetter,
        [array]$ViewButtons,
        $OptionsButton,
        $LogLabel,
        [array]$WorkflowButtons,
        [array]$OptionsButtons,
        [ValidateSet('Fill','Border')][string]$ActiveMode = 'Fill',
        [hashtable]$Brushes
    )
    $bc = [System.Windows.Media.BrushConverter]::new()
    $b = @{
        TitleBarActive    = $bc.ConvertFrom('#0078D4')
        TitleBarInactive  = $bc.ConvertFrom('#4BA3E0')
        DarkButtonBg      = $bc.ConvertFrom('#1E1E1E')
        DarkButtonBorder  = $bc.ConvertFrom('#555555')
        DarkActiveBg      = $bc.ConvertFrom('#3A3A3A')
        DarkActiveBorder  = $bc.ConvertFrom('#F9F9F9')
        LightWfBg         = $bc.ConvertFrom('#0078D4')
        LightWfBorder     = $bc.ConvertFrom('#006CBE')
        LightActiveBg     = $bc.ConvertFrom('#005A9E')
        LightActiveBorder = $bc.ConvertFrom('#FFFFFF')
        LightOptBg        = $bc.ConvertFrom('#0078D4')
        LightOptBorder    = $bc.ConvertFrom('#006CBE')
        LogLabelDark      = $bc.ConvertFrom('#B0B0B0')
        LogLabelLight     = $bc.ConvertFrom('#595959')
    }
    if ($Brushes) { foreach ($k in $Brushes.Keys) { $b[$k] = $Brushes[$k] } }
    $script:SuiteTheme = @{
        Window           = $Window
        IsDarkGetter     = $IsDarkGetter
        ActiveViewGetter = $ActiveViewGetter
        ViewButtons      = $ViewButtons
        OptionsButton    = $OptionsButton
        LogLabel         = $LogLabel
        WorkflowButtons  = $WorkflowButtons
        OptionsButtons   = $OptionsButtons
        ActiveMode       = $ActiveMode
        Brushes          = $b
    }
}

function Get-SuiteThemeContext {
    if (-not $script:SuiteTheme) {
        throw 'Initialize-SuiteTheme has not been called; the theme context is empty.'
    }
    return $script:SuiteTheme
}

function Update-TitleBarBrushes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only; no external state.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Sets both the WindowTitleBrush and NonActiveWindowTitleBrush per theme.')]
    param(
        $Window,
        [Nullable[bool]]$IsDark
    )
    $ctx = Get-SuiteThemeContext
    if ($null -eq $Window) { $Window = $ctx.Window }
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    if ($IsDark) {
        # Clearing lets the dark MahApps theme supply its own brushes.
        $Window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $Window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $Window.WindowTitleBrush          = $ctx.Brushes.TitleBarActive
        $Window.NonActiveWindowTitleBrush = $ctx.Brushes.TitleBarInactive
    }
}

function Update-SidebarButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param(
        [Nullable[bool]]$IsDark,
        [string]$ActiveView
    )
    $ctx = Get-SuiteThemeContext
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    if (-not $ActiveView -and $ctx.ActiveViewGetter) { $ActiveView = [string](& $ctx.ActiveViewGetter) }
    $b = $ctx.Brushes
    $thickness = [System.Windows.Thickness]::new(1)

    if ($ctx.ActiveMode -eq 'Border') {
        # Active-state visual is BorderBrush color only; thickness stays 1px.
        $bg           = if ($IsDark) { $b.DarkButtonBg }     else { $b.LightWfBg }
        $border       = if ($IsDark) { $b.DarkButtonBorder } else { $b.LightWfBorder }
        $activeBorder = if ($IsDark) { $b.DarkActiveBorder } else { $b.LightActiveBorder }
        foreach ($v in $ctx.ViewButtons) {
            if (-not $v.Button) { continue }
            $v.Button.Background      = $bg
            $v.Button.BorderBrush     = if ($v.Name -eq $ActiveView) { $activeBorder } else { $border }
            $v.Button.BorderThickness = $thickness
        }
        $idleBg = $bg
    }
    else {
        $idleBg   = if ($IsDark) { $b.DarkButtonBg }     else { $b.LightWfBg }
        $activeBg = if ($IsDark) { $b.DarkActiveBg }     else { $b.LightActiveBg }
        $border   = if ($IsDark) { $b.DarkButtonBorder } else { $b.LightWfBorder }
        foreach ($v in $ctx.ViewButtons) {
            if (-not $v.Button) { continue }
            $isActive = ($v.Name -eq $ActiveView)
            $v.Button.Background      = if ($isActive) { $activeBg } else { $idleBg }
            $v.Button.BorderBrush     = $border
            $v.Button.BorderThickness = $thickness
        }
    }
    if ($ctx.OptionsButton) {
        $ctx.OptionsButton.Background      = $idleBg
        $ctx.OptionsButton.BorderBrush     = $border
        $ctx.OptionsButton.BorderThickness = $thickness
    }
    if ($ctx.LogLabel) {
        $ctx.LogLabel.Foreground = if ($IsDark) { $b.LogLabelDark } else { $b.LogLabelLight }
    }
}

function Set-ButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param([Nullable[bool]]$IsDark)
    $ctx = Get-SuiteThemeContext
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    $b = $ctx.Brushes
    if ($IsDark) {
        foreach ($btn in @($ctx.WorkflowButtons)) { $btn.Background = $b.DarkButtonBg; $btn.BorderBrush = $b.DarkButtonBorder }
        foreach ($btn in @($ctx.OptionsButtons))  { $btn.Background = $b.DarkButtonBg; $btn.BorderBrush = $b.DarkButtonBorder }
        if ($ctx.LogLabel) { $ctx.LogLabel.Foreground = $b.LogLabelDark }
    }
    else {
        foreach ($btn in @($ctx.WorkflowButtons)) { $btn.Background = $b.LightWfBg;  $btn.BorderBrush = $b.LightWfBorder }
        foreach ($btn in @($ctx.OptionsButtons))  { $btn.Background = $b.LightOptBg; $btn.BorderBrush = $b.LightOptBorder }
        if ($ctx.LogLabel) { $ctx.LogLabel.Foreground = $b.LogLabelLight }
    }
}

function Set-DialogTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Window]$Dialog,
        [Nullable[bool]]$IsDark
    )
    $ctx = Get-SuiteThemeContext
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    if ($IsDark) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Dark.Steel') }
    else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Light.Blue')
        $Dialog.WindowTitleBrush          = $ctx.Brushes.TitleBarActive
        $Dialog.NonActiveWindowTitleBrush = $ctx.Brushes.TitleBarInactive
    }
}

# =============================================================================
# Themed dialogs
# =============================================================================
$script:ThemedMessageResult = $null

function Show-ThemedMessage {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Shows a modal MetroWindow and returns the chosen string; mutates no external state.')]
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('OK','OKCancel','YesNo')][string]$Buttons = 'OK',
        [ValidateSet('None','Info','Warn','Warning','Error','Question')][string]$Icon = 'None'
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Message"
    SizeToContent="Height"
    Width="460" MinHeight="160"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="NoResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,18,20,14">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="txtIcon" Grid.Row="0" Grid.Column="0" FontSize="20" VerticalAlignment="Top" Margin="0,0,14,0"/>
        <TextBlock x:Name="txtMsg"  Grid.Row="0" Grid.Column="1" FontSize="13" TextWrapping="Wrap" VerticalAlignment="Center"/>
        <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
            <Button x:Name="btnPrimary"   MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnSecondary" MinWidth="90" Height="32" IsCancel="True" Visibility="Collapsed"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $xmlReader = New-Object System.Xml.XmlNodeReader $xml
    $dlg = [System.Windows.Markup.XamlReader]::Load($xmlReader)
    Install-TitleBarDragFallback -Window $dlg

    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    $dlg.Owner = $Owner
    # Inactive brushes copy from the owner's own inactive properties; copying
    # the active ones renders the wrong title bar whenever the dialog loses
    # focus while the owner is in a different active/inactive state.
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.NonActiveWindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.NonActiveGlowBrush
    } catch { $null = $_ }

    $dlg.Title   = $Title
    $txtIcon     = $dlg.FindName('txtIcon')
    $txtMsg      = $dlg.FindName('txtMsg')
    $btn1        = $dlg.FindName('btnPrimary')
    $btn2        = $dlg.FindName('btnSecondary')
    $txtMsg.Text = $Message

    # Glyph-only state (no red/green); inherits ThemeForeground for AAA
    # contrast on both themes.
    $glyph = switch ($Icon) {
        'Info'     { [char]0x2139 }
        'Warn'     { [char]0x26A0 }
        'Warning'  { [char]0x26A0 }
        'Error'    { [char]0x2716 }
        'Question' { [char]0x003F }
        default    { '' }
    }
    $txtIcon.Text = [string]$glyph

    switch ($Buttons) {
        'OK' {
            $btn1.Content = 'OK'
            $btn2.Visibility = [System.Windows.Visibility]::Collapsed
        }
        'OKCancel' {
            $btn1.Content = 'OK'
            $btn2.Content = 'Cancel'
            $btn2.Visibility = [System.Windows.Visibility]::Visible
        }
        'YesNo' {
            $btn1.Content = 'Yes'
            $btn2.Content = 'No'
            $btn2.Visibility = [System.Windows.Visibility]::Visible
        }
    }

    $script:ThemedMessageResult = switch ($Buttons) { 'YesNo' { 'No' } default { 'Cancel' } }

    # A collapsed IsCancel button never registers its access key, so Escape
    # is dead on OK-only dialogs without this handler.
    $dlg.Add_PreviewKeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
            $e.Handled = $true
            $s.Close()
        }
    })

    # No GetNewClosure -- Show-ThemedMessage is still on the stack blocked on
    # ShowDialog, so $script: writes to $ThemedMessageResult reach this
    # module's script scope naturally. GetNewClosure would silently drop them.
    $btn1.Add_Click({
        $script:ThemedMessageResult = switch ($Buttons) { 'YesNo' { 'Yes' } default { 'OK' } }
        $dlg.Close()
    })
    $btn2.Add_Click({
        $script:ThemedMessageResult = switch ($Buttons) { 'YesNo' { 'No' } default { 'Cancel' } }
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()
    return $script:ThemedMessageResult
}

function Show-ConfirmDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)]$Owner
    )
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="480" SizeToContent="Height" MinWidth="380"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ResizeMode="NoResize" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtMsg" Grid.Row="0" TextWrapping="Wrap" FontSize="13" Margin="0,8,0,16"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnYes" Content="Yes" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnNo"  Content="No"  Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $Owner
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    # Owner-driven theming, same as Show-ThemedMessage: this dialog stays
    # safe to call before Initialize-SuiteTheme has run.
    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.NonActiveWindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.NonActiveGlowBrush
    } catch { $null = $_ }
    $dlg.FindName('txtMsg').Text = $Message
    $dlg.FindName('btnYes').Add_Click({ $dlg.DialogResult = $true;  $dlg.Close() })
    $dlg.FindName('btnNo').Add_Click({  $dlg.DialogResult = $false; $dlg.Close() })
    return [bool]$dlg.ShowDialog()
}


# =============================================================================
# Background work: runspace lifecycle helpers
# =============================================================================
# Explicit-argument by design: the shells keep their own $script:Bg* state
# (their scan/refresh wiring reads it directly, 30-80 sites per shell);
# only the lifecycle logic lives here.

function New-SuiteBgRunspace {
    <#
    .SYNOPSIS
        Creates and opens the STA background runspace, pre-importing the
        tool module and attaching its logging.
    #>
    param(
        [Parameter(Mandatory)][string]$ModulePath,
        [string]$LogPath
    )
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $initPS = [powershell]::Create()
    $initPS.Runspace = $rs
    [void]$initPS.AddScript({
        param($ModulePath, $LogPath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
        if ($LogPath) { Initialize-Logging -LogPath $LogPath -Attach }
    }).AddArgument($ModulePath).AddArgument($LogPath)
    [void]$initPS.Invoke()
    $initPS.Dispose()
    return $rs
}

function Stop-SuiteBgWork {
    <#
    .SYNOPSIS
        Non-blocking teardown of an in-flight background pipeline.

    .DESCRIPTION
        BeginStop, not Stop: a synchronous Stop() blocks the UI thread for
        as long as the pipeline is stuck inside a CM/CIM call against an
        unresponsive provider. The stopping pipeline is parked in the
        graveyard and reaped on a later pass once it has actually stopped.
        Returns the surviving graveyard; wrap the result in @() at the
        call site (an emptied graveyard returns nothing).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing.')]
    param(
        $PowerShell,
        $Timer,
        [array]$Graveyard = @()
    )
    if ($Timer) { try { $Timer.Stop() } catch { $null = $_ } }
    $yard = @($Graveyard)
    if ($PowerShell) {
        try { [void]$PowerShell.BeginStop($null, $null) } catch { $null = $_ }
        $yard += ,$PowerShell
    }
    $kept = @($yard | Where-Object {
        if ($_.InvocationStateInfo.State -in @('Stopped', 'Completed', 'Failed')) {
            try { $_.Dispose() } catch { $null = $_ }
            $false
        }
        else { $true }
    })
    return $kept
}

function Close-SuiteBgRunspace {
    <#
    .SYNOPSIS
        Asynchronously closes the background runspace at shell exit.

    .DESCRIPTION
        CloseAsync: a blocking Close() waits for a still-stopping pipeline
        and freezes shutdown for as long as that pipeline is stuck.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing.')]
    param($Runspace)
    if ($Runspace) { try { $Runspace.CloseAsync() } catch { $null = $_ } }
}

# =============================================================================
# CM collection widgets: folder tree + picker dialog
# =============================================================================
$script:CollectionTreeLeafCount = 0
$script:PickerResult = $null

function Build-CollectionTree {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Build is the natural verb for tree assembly.')]
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TreeView]$TreeView,
        [Parameter(Mandatory)]$AllCollections,
        $AllFolders,
        [string]$Needle = ''
    )

    $colls   = @($AllCollections)
    $folders = @($AllFolders)

    $needleLower = ([string]$Needle).Trim().ToLowerInvariant()
    $hasFilter   = -not [string]::IsNullOrWhiteSpace($needleLower)

    $foldersByParent = @{}
    foreach ($f in $folders) {
        $parentId = [int]$f.ParentID
        if (-not $foldersByParent.ContainsKey($parentId)) { $foldersByParent[$parentId] = @() }
        $foldersByParent[$parentId] += $f
    }
    $collectionsByFolder = @{}
    foreach ($c in $colls) {
        $fid = [int]$c.FolderID
        if (-not $collectionsByFolder.ContainsKey($fid)) { $collectionsByFolder[$fid] = @() }
        $collectionsByFolder[$fid] += $c
    }

    $TreeView.Items.Clear()
    $script:CollectionTreeLeafCount = 0

    $matchesNeedle = {
        param($Coll)
        if (-not $hasFilter) { return $true }
        $name = ([string]$Coll.Name).ToLowerInvariant()
        $id   = ([string]$Coll.CollectionID).ToLowerInvariant()
        return ($name.Contains($needleLower) -or $id.Contains($needleLower))
    }

    $populate = {
        param($ParentNode, [int]$FolderID)
        $any = $false

        $childFolders = if ($foldersByParent.ContainsKey($FolderID)) { @($foldersByParent[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($f in $childFolders) {
            $folderNode = New-Object System.Windows.Controls.TreeViewItem
            $folderNode.Header = ('[+] {0}' -f $f.Name)
            $folderNode.Tag = @{ Type = 'Folder'; Object = $f }
            $folderNode.FontWeight = [System.Windows.FontWeights]::SemiBold
            if ($hasFilter) { $folderNode.IsExpanded = $true }
            $hadAny = & $populate $folderNode ([int]$f.FolderID)
            if ($hadAny -or -not $hasFilter) {
                [void]$ParentNode.Items.Add($folderNode)
                $any = $true
            }
        }

        $childColls = if ($collectionsByFolder.ContainsKey($FolderID)) { @($collectionsByFolder[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($c in $childColls) {
            if (-not (& $matchesNeedle $c)) { continue }
            $collNode = New-Object System.Windows.Controls.TreeViewItem
            $collNode.Header = ('{0}  ({1}, {2} members)' -f $c.Name, $c.CollectionID, $c.MemberCount)
            $collNode.Tag = @{ Type = 'Collection'; Object = $c }
            [void]$ParentNode.Items.Add($collNode)
            $script:CollectionTreeLeafCount++
            $any = $true
        }
        return $any
    }

    # Void the invocation: the populate scriptblock returns its any-added
    # flag, which otherwise leaks into the pipeline ahead of the count and
    # turns the return value into @(bool, int).
    $null = & $populate $TreeView 0
    return $script:CollectionTreeLeafCount
}

function Show-CollectionPickerDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)]$Collections,
        $Folders,
        [string]$Title = 'Pick Collection',
        [bool]$IncludeBuiltIn = $true
    )

    if (-not $Collections -or @($Collections).Count -eq 0) {
        try { Write-Log -Message 'Collection picker opened with no collections loaded; refresh first.' -Level WARN } catch { $null = $_ }
        return $null
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title=""
    Width="720" Height="640"
    MinWidth="540" MinHeight="420"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="txtPickerFilter" Grid.Row="0" FontSize="12" Padding="6,4,6,4" Margin="0,4,0,8"
                 Controls:TextBoxHelper.Watermark="Filter by collection name or ID..."/>
        <Border Grid.Row="1" BorderThickness="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                Background="{DynamicResource MahApps.Brushes.ThemeBackground}">
            <TreeView x:Name="treePicker" FontSize="12"
                      VirtualizingStackPanel.IsVirtualizing="True"
                      VirtualizingStackPanel.VirtualizationMode="Recycling"
                      Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                      Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                      BorderThickness="0"/>
        </Border>
        <TextBlock x:Name="txtPickerStatus" Grid.Row="2" FontSize="11" Margin="0,8,0,0"
                   Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                   Text="Pick a collection (folders are not selectable)."/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True" IsEnabled="False"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $Owner
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtPickerFilter = $dlg.FindName('txtPickerFilter')
    $treePicker      = $dlg.FindName('treePicker')
    $txtPickerStatus = $dlg.FindName('txtPickerStatus')
    $btnOk           = $dlg.FindName('btnOk')
    $btnCancel       = $dlg.FindName('btnCancel')

    $allCollections = if ($IncludeBuiltIn) { $Collections } else { $Collections | Where-Object { -not $_.IsBuiltIn } }
    $allCollections = @($allCollections)
    $allFolders     = @($Folders)

    $rebuildTree = {
        param([string]$Needle)
        $count = Build-CollectionTree -TreeView $treePicker `
            -AllCollections $allCollections `
            -AllFolders     $allFolders `
            -Needle         $Needle
        if ([string]::IsNullOrWhiteSpace($Needle)) {
            $totalColls = @($allCollections).Count
            $totalFolders = @($allFolders).Count
            $txtPickerStatus.Text = ('{0} collections across {1} folders. Pick one (folders are not selectable).' -f $totalColls, $totalFolders)
        } else {
            $txtPickerStatus.Text = ('{0} collections match "{1}".' -f $count, $Needle.Trim())
        }
    }

    & $rebuildTree ''

    $script:PickerResult = $null
    $treePicker.Add_SelectedItemChanged({
        $node = $treePicker.SelectedItem
        if (-not $node -or -not $node.Tag -or $node.Tag.Type -ne 'Collection') {
            $btnOk.IsEnabled = $false
            $script:PickerResult = $null
            return
        }
        $btnOk.IsEnabled = $true
        $script:PickerResult = $node.Tag.Object
    })

    $txtPickerFilter.Add_TextChanged({ & $rebuildTree ([string]$txtPickerFilter.Text) })

    $btnOk.Add_Click({
        if ($script:PickerResult) {
            $dlg.DialogResult = $true
            $dlg.Close()
        }
    })
    $btnCancel.Add_Click({
        $script:PickerResult = $null
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()
    return $script:PickerResult
}

