<#
.SYNOPSIS
    Shared core module for the Configuration Manager tool suite.

.DESCRIPTION
    Vendored into each consumer tool at Lib\SuiteCommon\ by
    sync-suitecommon.ps1. Provides:
      - Structured logging (Initialize-Logging, Write-Log, Write-LogErrorRecord)
      - ConfigMgr site connection management (Connect-CMSite, Disconnect-CMSite,
        Test-CMConnection, Get-CMConnectionInfo)
      - Flat JSON settings persistence (Read-SuiteSettings, Save-SuiteSettings)

    Consumers never edit the vendored copy; changes flow through the
    app-packager-suite repository and synchronize again.

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
# Host environment
# ---------------------------------------------------------------------------

function Repair-WindowsPowerShellModulePath {
    <#
    .SYNOPSIS
        Removes PowerShell 7 module directories from the process PSModulePath.

    .DESCRIPTION
        A Windows PowerShell process launched from PowerShell 7 inherits the
        7.x module directories ahead of the 5.1 ones. The launching runspace
        copes, but every runspace opened later in the process and every child
        powershell.exe autoloads Microsoft.PowerShell.Utility from the 7.x
        manifest, which carries no Get-FileHash or ConvertFrom-Json under
        5.1. The process environment is rewritten once, so anything created
        afterwards inherits the 5.1 roots. Runs at module load; safe to call
        again.

    .OUTPUTS
        [bool] True when the path changed.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Process-environment hygiene at startup; no user-visible state to confirm.')]
    param()

    # The inheritance problem exists only in a Windows PowerShell process;
    # under PowerShell 7 the 7.x roots are the host's own and must stay.
    if ($PSVersionTable.PSEdition -eq 'Core') { return $false }
    $current = [string]$env:PSModulePath
    $winPsModules = Join-Path $PSHOME 'Modules'
    $roots = @($current -split ';' | Where-Object { $_ -and $_ -notmatch '(?i)[\\/]PowerShell[\\/](7[\\/]|Modules)|microsoft\.powershell_' })
    if ($roots -notcontains $winPsModules) { $roots = @($winPsModules) + $roots }
    $repaired = ($roots -join ';')
    if ($repaired -eq $current) { return $false }
    $env:PSModulePath = $repaired
    return $true
}

# Import-time repair: consumers import this module before opening any
# background runspace or child process.
[void](Repair-WindowsPowerShellModulePath)

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
        SMS Provider computer. Optional: falls back to the SUITE_CM_PROVIDER
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
            Write-Log "ConfigurationManager module not found. Make sure that the Configuration Manager console is installed." -Level ERROR
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
            Write-Log "No CMSite drive '$SiteCode' exists and no SMS Provider computer was given (parameter or SUITE_CM_PROVIDER)." -Level ERROR
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
        Write-Log "Connected to ConfigMgr site: $SiteCode"
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
    Write-Log "Disconnected from ConfigMgr site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a ConfigMgr site.
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
        Asks the SMS Provider which site code(s) it serves.

    .DESCRIPTION
        Queries root\sms:SMS_ProviderLocation on the SMS Provider computer. A
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

    # Never use .Keys / .Values / .Count on a settings hashtable: a setting
    # with that name makes the dot access return the setting's value, not the
    # collection. GetEnumerator() and the indexer are not shadowed.
    $settings = @{}
    foreach ($entry in $Defaults.GetEnumerator()) { $settings[$entry.Key] = $entry.Value }

    if (Test-Path -LiteralPath $Path) {
        try {
            # -Encoding UTF8: on 5.1 the default decoder mangles non-ASCII in
            # BOM-less UTF-8 files (hand-edited settings) without any error.
            $loaded = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            # PSObject.Properties reads only real JSON properties. A dot access
            # such as $loaded.Count returns 1 on PowerShell 7 when the file has
            # no such property.
            foreach ($k in @($Defaults.GetEnumerator() | ForEach-Object { $_.Key })) {
                $prop = $loaded.PSObject.Properties[$k]
                if ($prop -and $null -ne $prop.Value) { $settings[$k] = $prop.Value }
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

    $tempPath = $null
    try {
        $parentDir = Split-Path -Path $Path -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
        }
        # Write the new text to a temp file in the same folder, then swap it in.
        # A crash or a full disk during the write leaves the old file whole,
        # where a direct Set-Content onto the target leaves a truncated one.
        # The temp file shares the volume, so the swap is one file system call.
        $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $json = $Settings | ConvertTo-Json -Depth $Depth
        $tempPath = '{0}.{1}.tmp' -f $fullPath, [guid]::NewGuid().ToString('N')
        # -ErrorAction Stop: Set-Content fails non-terminating by default,
        # which would skip the catch and report success on a failed write.
        Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8 -ErrorAction Stop
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            # [NullString]::Value: a PowerShell $null becomes an empty string here, and File.Replace rejects it as a path.
            [System.IO.File]::Replace($tempPath, $fullPath, [NullString]::Value)
        }
        else {
            [System.IO.File]::Move($tempPath, $fullPath)
        }
        $tempPath = $null
        return $true
    }
    catch {
        Write-Log ("Failed to save settings to {0}: {1}" -f $Path, $_.Exception.Message) -Level WARN -Quiet
        return $false
    }
    finally {
        if ($tempPath -and (Test-Path -LiteralPath $tempPath)) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
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
# One mouse, one pending title-bar drag: the window and the press point.
$script:TitleBarDragPending    = $null

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
            $WM_NCHITTEST = 0x0084; $WM_NCLBUTTONDBLCLK = 0x00A3; $HTCAPTION = 2
            if ($msg -eq $WM_NCLBUTTONDBLCLK) {
                # Our HTCAPTION answer routes the double-click here as a
                # non-client message; nothing else toggles the window state.
                try {
                    if ($wParam.ToInt64() -ne $HTCAPTION) { return [IntPtr]::Zero }
                    $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                    if ($target -and (Switch-WindowMaximized -Window $target)) { $handled.Value = $true }
                } catch { $null = $_ }
                return [IntPtr]::Zero
            }
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

function Switch-WindowMaximized {
    <#
    .SYNOPSIS
        Toggles a window between Maximized and Normal. Returns $true when
        the state changed; $false for a window whose ResizeMode forbids
        maximizing.
    #>
    param([Parameter(Mandatory)][System.Windows.Window]$Window)
    if ($Window.ResizeMode -ne [System.Windows.ResizeMode]::CanResize -and
        $Window.ResizeMode -ne [System.Windows.ResizeMode]::CanResizeWithGrip) { return $false }
    if ($Window.WindowState -eq [System.Windows.WindowState]::Maximized) {
        $Window.WindowState = [System.Windows.WindowState]::Normal
    }
    else {
        $Window.WindowState = [System.Windows.WindowState]::Maximized
    }
    return $true
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
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            if ($e.ClickCount -ge 2) {
                $script:TitleBarDragPending = $null
                if (Switch-WindowMaximized -Window $s) { $e.Handled = $true }
                return
            }
            # DragMove runs the native move loop, which consumes the button-up;
            # WPF then filters the next button-down as a repeat and the second
            # click of a double-click is lost. The drag therefore starts on the
            # first mouse move past the system drag distance, never on the click.
            $script:TitleBarDragPending = @{ Window = $s; Start = $pos }
            $e.Handled = $true
        } catch { $null = $_ }
    })
    $Window.Add_PreviewMouseMove({
        param($s, $e)
        try {
            $pending = $script:TitleBarDragPending
            if (-not $pending -or -not [object]::ReferenceEquals($pending.Window, $s)) { return }
            if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) {
                $script:TitleBarDragPending = $null
                return
            }
            $pos = $e.GetPosition($s)
            if ([Math]::Abs($pos.X - $pending.Start.X) -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
                [Math]::Abs($pos.Y - $pending.Start.Y) -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }
            $script:TitleBarDragPending = $null
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $s.DragMove()
            $e.Handled = $true
        } catch { $script:TitleBarDragPending = $null }
    })
    $Window.Add_PreviewMouseLeftButtonUp({ param($s, $e) $script:TitleBarDragPending = $null })
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

        if ($null -ne $left -and $null -ne $top -and $null -ne $w -and $null -ne $h) {
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
        # Geometry is applied even when maximizing: it becomes the window's
        # normal bounds, so un-maximizing returns to the saved size instead
        # of the XAML defaults.
        if ($s.Maximized) {
            $Window.WindowState = [System.Windows.WindowState]::Maximized
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
        TitleBarInactiveText = $bc.ConvertFrom('#333333')
        DarkTitleBarInactive = $bc.ConvertFrom('#6B6B6B')
        DarkButtonBg      = $bc.ConvertFrom('#1E1E1E')
        DarkButtonBorder  = $bc.ConvertFrom('#8C8C8C')
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
        OnBlue            = $bc.ConvertFrom('#FFFFFF')
        AccentHoverLight  = $bc.ConvertFrom('#006CBE')
        AccentPressedLight = $bc.ConvertFrom('#005A9E')
        AccentHoverDark   = $bc.ConvertFrom('#5E6F80')
        AccentPressedDark = $bc.ConvertFrom('#53626F')
        NeutralHoverLight = $bc.ConvertFrom('#E6E6E6')
        NeutralPressedLight = $bc.ConvertFrom('#CFCFCF')
        NeutralHoverDark  = $bc.ConvertFrom('#3A3A3A')
        NeutralPressedDark = $bc.ConvertFrom('#4A4A4A')
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
        # The MahApps inactive brush (#808080) is 3.9:1 under white text. A fixed gray of 5.3:1 replaces it.
        $Window.NonActiveWindowTitleBrush = $ctx.Brushes.DarkTitleBarInactive
    } else {
        $Window.WindowTitleBrush          = $ctx.Brushes.TitleBarActive
        $Window.NonActiveWindowTitleBrush = $ctx.Brushes.TitleBarInactive
    }
    Register-SuiteTitleForeground -Window $Window -IsDark $IsDark
    Set-SuiteAccentButtonResource -Window $Window -IsDark $IsDark
}

function Set-SuiteAccentButtonResource {
    <#
    .SYNOPSIS
        Gives a window the explicit hover and pressed brushes of the accent button, and the template that uses them.
    .DESCRIPTION
        The MahApps accent button turns black (light theme) or white (dark theme) when pressed. A style that sets
        Template to {DynamicResource SuiteAccentButtonTemplate} keeps the fill blue in every state: hover is one
        darker step, pressed one more, white bold text stays at 4.5:1 or more. The brushes follow the theme and
        change when the theme changes, because the template reads them with DynamicResource.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Adds resources to one window only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [Nullable[bool]]$IsDark)
    if ($Window -isnot [System.Windows.FrameworkElement]) { return }
    $ctx = Get-SuiteThemeContext
    $dark = if ($null -ne $IsDark) { [bool]$IsDark } else { [bool](& $ctx.IsDarkGetter) }
    if (-not $Window.Resources.Contains('SuiteAccentButtonTemplate')) {
        $xaml = @'
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">
    <Border x:Name="Bd"
            Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
            BorderThickness="{TemplateBinding BorderThickness}" SnapsToDevicePixels="True">
        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                          VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                          Margin="{TemplateBinding Padding}" RecognizesAccessKey="True" SnapsToDevicePixels="True"/>
    </Border>
    <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter TargetName="Bd" Property="Background" Value="{DynamicResource SuiteAccentHover}"/>
        </Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True">
            <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.ThemeForeground}"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter TargetName="Bd" Property="Background" Value="{DynamicResource SuiteAccentPressed}"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
        </Trigger>
    </ControlTemplate.Triggers>
</ControlTemplate>
'@
        $Window.Resources['SuiteAccentButtonTemplate'] = [System.Windows.Markup.XamlReader]::Parse($xaml)
    }
    if (-not $Window.Resources.Contains('SuiteNeutralButtonTemplate')) {
        # The same template for the normal button: its own fill and border stay, hover and pressed move one step in the neutral family.
        $neutral = $xaml.Replace('SuiteAccentHover', 'SuiteNeutralHover').Replace('SuiteAccentPressed', 'SuiteNeutralPressed')
        $Window.Resources['SuiteNeutralButtonTemplate'] = [System.Windows.Markup.XamlReader]::Parse($neutral)
    }
    $Window.Resources['SuiteNeutralHover'] = $(if ($dark) { $ctx.Brushes.NeutralHoverDark } else { $ctx.Brushes.NeutralHoverLight })
    $Window.Resources['SuiteNeutralPressed'] = $(if ($dark) { $ctx.Brushes.NeutralPressedDark } else { $ctx.Brushes.NeutralPressedLight })
    $hover = if ($dark) { $ctx.Brushes.AccentHoverDark } else { $ctx.Brushes.AccentHoverLight }
    $pressed = if ($dark) { $ctx.Brushes.AccentPressedDark } else { $ctx.Brushes.AccentPressedLight }
    $Window.Resources['SuiteAccentHover'] = $hover
    $Window.Resources['SuiteAccentPressed'] = $pressed
}

function Set-SuiteTitleForeground {
    <#
    .SYNOPSIS
        Sets the caption text of a MetroWindow for its state: light theme, inactive: dark gray on the light blue bar;
        every other state keeps the library text (white on the active blue bar and on the dark bars).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only; no external state.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [Nullable[bool]]$IsDark)
    $ctx = Get-SuiteThemeContext
    $dark = if ($null -ne $IsDark) { [bool]$IsDark } else { [bool](& $ctx.IsDarkGetter) }
    if (-not $dark -and -not $Window.IsActive) {
        $Window.TitleForeground = $ctx.Brushes.TitleBarInactiveText
    } else {
        $Window.ClearValue([MahApps.Metro.Controls.MetroWindow]::TitleForegroundProperty)
    }
}

function Register-SuiteTitleForeground {
    <#
    .SYNOPSIS
        Gives a MetroWindow the caption text of its state now and whenever it is activated or deactivated.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only; no external state.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [Nullable[bool]]$IsDark)
    if (-not $Window.Resources.Contains('SuiteTitleForegroundHooked')) {
        $Window.Resources['SuiteTitleForegroundHooked'] = $true
        $Window.Add_Activated({ param($sender, $e) try { Set-SuiteTitleForeground -Window $sender } catch { $null = $_ } })
        $Window.Add_Deactivated({ param($sender, $e) try { Set-SuiteTitleForeground -Window $sender } catch { $null = $_ } })
    }
    Set-SuiteTitleForeground -Window $Window -IsDark $IsDark
}

# Light theme paints buttons blue. A plain MahApps Square button keeps the dark
# theme text brush there, so the text is black on blue. Every blue-background
# control gets white text (OnBlue) in light mode and its own brush back in dark mode.
function Set-ThemeForeground {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param($Control, [bool]$IsDark, $OnBlue)
    if ($null -eq $Control -or -not $Control.PSObject.Properties['Foreground']) { return }
    if ($IsDark) {
        if ($Control -is [System.Windows.DependencyObject]) { $Control.ClearValue([System.Windows.Controls.Control]::ForegroundProperty) }
        else { $Control.Foreground = $null }
    }
    else { $Control.Foreground = $OnBlue }
}

# Field titles are bold; field values stay in the normal body style. The site
# code and the SMS provider (site server) titles are 14 pt bold. The same in
# light and dark mode: a theme changes colors only, never font size, weight,
# style or decoration. The title of a box is the TextBlock or Label next to it:
# the left neighbour in a Grid row (or the cell above), or the previous child
# of a StackPanel or DockPanel.
$script:ConnectionBoxNames = @('txtSiteCode', 'txtSmsProvider', 'txtProvider', 'txtSiteServer', 'txtWizSiteCode', 'txtWizProvider')
$script:FieldTitleFontSize = 14.0

function Find-FieldTitle {
    param($Box)
    $parent = $null
    try { $parent = $Box.Parent } catch { $parent = $null }
    if ($null -eq $parent) { return $null }
    $isTitle = { param($c) ($c -is [System.Windows.Controls.TextBlock]) -or ($c -is [System.Windows.Controls.Label]) }
    if ($parent -is [System.Windows.Controls.Grid]) {
        $row = [System.Windows.Controls.Grid]::GetRow($Box); $col = [System.Windows.Controls.Grid]::GetColumn($Box)
        $cands = @()
        foreach ($c in $parent.Children) {
            if ($c -eq $Box -or -not (& $isTitle $c)) { continue }
            $r = [System.Windows.Controls.Grid]::GetRow($c); $k = [System.Windows.Controls.Grid]::GetColumn($c)
            if ($r -eq $row -and $k -lt $col) { $cands += [pscustomobject]@{ E = $c; D = ($col - $k) } }
            elseif ($k -eq $col -and $r -lt $row) { $cands += [pscustomobject]@{ E = $c; D = 100 + ($row - $r) } }
        }
        if ($cands.Count) { return ($cands | Sort-Object D | Select-Object -First 1).E }
        return $null
    }
    if ($parent -is [System.Windows.Controls.Panel]) {
        $prev = $null
        foreach ($c in $parent.Children) { if ($c -eq $Box) { break }; $prev = $c }
        if ($null -ne $prev -and (& $isTitle $prev)) { return $prev }
    }
    return $null
}

function Set-ConnectionTextTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window font properties only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'IsDark', Justification='Kept so every theme helper takes the same arguments; fonts do not depend on the theme.')]
    param($Window, [bool]$IsDark)
    if ($null -eq $Window -or -not $Window.PSObject.Methods['FindName']) { return }
    foreach ($name in $script:ConnectionBoxNames) {
        $box = $null
        try { $box = $Window.FindName($name) } catch { $box = $null }
        if ($null -eq $box) { continue }
        $title = Find-FieldTitle -Box $box
        if ($null -eq $title) { continue }
        $title.FontSize   = [Math]::Max([double]$title.FontSize, $script:FieldTitleFontSize)
        $title.FontWeight = [System.Windows.FontWeights]::Bold
    }
}

# Footer. One shared style for all footer text in every app, in light and dark:
# the same font family, size, weight, style and decoration. A theme changes
# colors only, through the MahApps text brush. The footer shows the status text
# on the left and the site connection on the right: "Site MCM [middle dot] cm01.contoso.com".
# The connection is read-only text. A click may open Options at the connection fields.
$script:FooterSiteName = 'txtFooterSite'

function Get-SuiteFooterStyle {
    [CmdletBinding()]
    param()
    [pscustomobject]@{
        FontFamily     = 'Segoe UI'
        FontSize       = 11.0
        FontWeight     = 'Normal'
        FontStyle      = 'Normal'
        TextDecoration = 'None'
        Foreground     = 'MahApps.Brushes.Text'
    }
}

function Format-SuiteFooterSite {
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$SiteCode, [string]$SiteServer)
    $code = ([string]$SiteCode).Trim()
    $server = ([string]$SiteServer).Trim()
    $sep = [string][char]0x00B7
    if (-not $code -and -not $server) { return 'Site not set' }
    if (-not $code) { return ('Site not set {0} {1}' -f $sep, $server) }
    if (-not $server) { return ('Site {0} {1} server not set' -f $code, $sep) }
    return ('Site {0} {1} {2}' -f $code, $sep, $server)
}

function Set-SuiteFooterFont {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window font properties only.')]
    param($Element)
    if ($null -eq $Element) { return }
    $s = Get-SuiteFooterStyle
    $Element.FontFamily = New-Object System.Windows.Media.FontFamily($s.FontFamily)
    $Element.FontSize   = [double]$s.FontSize
    $Element.FontWeight = [System.Windows.FontWeights]::($s.FontWeight)
    $Element.FontStyle  = [System.Windows.FontStyles]::($s.FontStyle)
    if ($Element -is [System.Windows.Controls.TextBlock]) {
        $Element.TextDecorations = $null
        $Element.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
    }
    $Element.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $s.Foreground)
}

# Puts the site connection text at the right end of the status bar and gives the
# status text and the connection text the same footer style. The status bar is
# the Border that holds the status TextBlock. Call it once after the window loads.
function Initialize-SuiteFooter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Window,
        [string]$StatusName = 'txtStatus',
        [scriptblock]$OnOpenOptions
    )
    $status = $Window.FindName($StatusName)
    if ($null -eq $status) { return $null }
    Set-SuiteFooterFont $status
    $existing = $Window.FindName($script:FooterSiteName)
    if ($null -ne $existing) { Set-SuiteFooterFont $existing; return $existing }
    $bar = $status.Parent
    $isBorder = ($bar -is [System.Windows.Controls.Border])
    $isDock = ($bar -is [System.Windows.Controls.DockPanel])
    if (-not $isBorder -and -not $isDock) {
        $kind = if ($bar) { $bar.GetType().Name } else { 'nothing' }
        throw ('Initialize-SuiteFooter: the status text must sit in a Border or a DockPanel, not in {0}.' -f $kind)
    }

    $site = New-Object System.Windows.Controls.TextBlock
    $site.Text = (Format-SuiteFooterSite -SiteCode '' -SiteServer '')
    $site.Margin = [System.Windows.Thickness]::new(12, 0, 0, 0)
    $site.VerticalAlignment = 'Center'
    $site.HorizontalAlignment = 'Right'
    $site.ToolTip = 'Site code and site server. Open Options to change them.'
    [System.Windows.Automation.AutomationProperties]::SetName($site, 'Site connection')
    if ($OnOpenOptions) {
        $handler = $OnOpenOptions
        $open = {
            try { & $handler } catch { Write-Verbose ('The footer click failed: {0}' -f $_.Exception.Message) }
        }.GetNewClosure()
        $site.Cursor = [System.Windows.Input.Cursors]::Hand
        $site.Focusable = $true
        $site.Add_MouseLeftButtonUp({ & $open }.GetNewClosure())
        $site.Add_KeyUp({ param($sender, $e) if ($e.Key -eq 'Enter' -or $e.Key -eq 'Space') { & $open } }.GetNewClosure())
    }

    if ($isDock) {
        [System.Windows.Controls.DockPanel]::SetDock($site, [System.Windows.Controls.Dock]::Right)
        $bar.Children.Insert($bar.Children.IndexOf($status), $site)
    }
    else {
        $grid = New-Object System.Windows.Controls.Grid
        $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::Auto
        [void]$grid.ColumnDefinitions.Add($c0); [void]$grid.ColumnDefinitions.Add($c1)
        $bar.Child = $null
        [void]$grid.Children.Add($status)
        [System.Windows.Controls.Grid]::SetColumn($site, 1)
        [void]$grid.Children.Add($site)
        $bar.Child = $grid
    }
    $site.Name = $script:FooterSiteName
    try { $Window.RegisterName($script:FooterSiteName, $site) }
    catch { Write-Verbose ('The footer name was not registered: {0}' -f $_.Exception.Message) }
    $status.VerticalAlignment = 'Center'
    Set-SuiteFooterFont $site
    return $site
}

function Update-SuiteFooter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window text only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [string]$SiteCode, [string]$SiteServer)
    $site = $Window.FindName($script:FooterSiteName)
    if ($null -eq $site) { return }
    $site.Text = Format-SuiteFooterSite -SiteCode $SiteCode -SiteServer $SiteServer
}

# Implicit styles that every suite window takes in both themes. A text box and a combo box show a Gray2 border: the
# MahApps light border (#CCCCCC) is 1.6:1 on white and a border needs 3:1 (Gray2 is 4.0:1 light, 3.8:1 dark). A column
# header in a disabled grid shows the gray text of the other disabled controls. The styles are built in code from the
# MahApps styles of the window, once per window. A window that has an implicit style of its own for the type keeps it.
# True when the dictionary, or a dictionary merged into it that does not come from MahApps or ControlzEx, holds the key.
function Test-SuiteOwnResource {
    param($Dictionary, $Key)
    if ($null -eq $Dictionary) { return $false }
    if (@($Dictionary.Keys) -contains $Key) { return $true }
    foreach ($m in $Dictionary.MergedDictionaries) {
        $src = [string]$m.Source
        if ($src -match 'MahApps\.Metro|ControlzEx') { continue }
        if (Test-SuiteOwnResource -Dictionary $m -Key $Key) { return $true }
    }
    return $false
}

function Add-SuiteImplicitStyle {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Adds styles to the resources of one window only.')]
    param($Window)
    if ($null -eq $Window -or $Window -isnot [System.Windows.FrameworkElement]) { return }
    $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    # A dynamic resource reference needs the XAML reader: a setter made in code would keep the markup extension as its value.
    $defs = @(
        @{ Type = [System.Windows.Controls.TextBox]; Base = 'MahApps.Styles.TextBox'; Xaml = "<Style $ns TargetType=`"TextBox`"><Setter Property=`"BorderBrush`" Value=`"{DynamicResource MahApps.Brushes.Gray2}`"/></Style>" }
        @{ Type = [System.Windows.Controls.ComboBox]; Base = 'MahApps.Styles.ComboBox'; Xaml = "<Style $ns TargetType=`"ComboBox`"><Setter Property=`"BorderBrush`" Value=`"{DynamicResource MahApps.Brushes.Gray2}`"/></Style>" }
        @{ Type = [System.Windows.Controls.Primitives.DataGridColumnHeader]; Base = 'MahApps.Styles.DataGridColumnHeader'; Xaml = "<Style $ns TargetType=`"DataGridColumnHeader`"><Style.Triggers><Trigger Property=`"IsEnabled`" Value=`"False`"><Setter Property=`"Foreground`" Value=`"{DynamicResource MahApps.Brushes.Gray.MouseOver}`"/></Trigger></Style.Triggers></Style>" }
    )
    foreach ($d in $defs) {
        # A window that has an implicit style of its own for the type, in its dictionary or in a merged one that does not
        # come from MahApps, keeps it.
        if (Test-SuiteOwnResource -Dictionary $Window.Resources -Key $d.Type) { continue }
        $base = $Window.TryFindResource($d.Base)
        if ($base -isnot [System.Windows.Style]) { continue }
        $style = [System.Windows.Markup.XamlReader]::Parse($d.Xaml)
        $style.BasedOn = $base
        $Window.Resources.Add($d.Type, $style)
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
            Set-ThemeForeground -Control $v.Button -IsDark $IsDark -OnBlue $b.OnBlue
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
            Set-ThemeForeground -Control $v.Button -IsDark $IsDark -OnBlue $b.OnBlue
        }
    }
    if ($ctx.OptionsButton) {
        $ctx.OptionsButton.Background      = $idleBg
        $ctx.OptionsButton.BorderBrush     = $border
        $ctx.OptionsButton.BorderThickness = $thickness
        Set-ThemeForeground -Control $ctx.OptionsButton -IsDark $IsDark -OnBlue $b.OnBlue
    }
    if ($ctx.LogLabel) {
        $ctx.LogLabel.Foreground = if ($IsDark) { $b.LogLabelDark } else { $b.LogLabelLight }
    }
    Add-SuiteImplicitStyle -Window $ctx.Window
    Set-ConnectionTextTheme -Window $ctx.Window -IsDark $IsDark
}

function Set-ButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param([Nullable[bool]]$IsDark)
    $ctx = Get-SuiteThemeContext
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    $b = $ctx.Brushes
    if ($IsDark) {
        foreach ($btn in @($ctx.WorkflowButtons)) { $btn.Background = $b.DarkButtonBg; $btn.BorderBrush = $b.DarkButtonBorder; Set-ThemeForeground -Control $btn -IsDark $true -OnBlue $b.OnBlue }
        foreach ($btn in @($ctx.OptionsButtons))  { $btn.Background = $b.DarkButtonBg; $btn.BorderBrush = $b.DarkButtonBorder; Set-ThemeForeground -Control $btn -IsDark $true -OnBlue $b.OnBlue }
        if ($ctx.LogLabel) { $ctx.LogLabel.Foreground = $b.LogLabelDark }
    }
    else {
        foreach ($btn in @($ctx.WorkflowButtons)) { $btn.Background = $b.LightWfBg;  $btn.BorderBrush = $b.LightWfBorder; Set-ThemeForeground -Control $btn -IsDark $false -OnBlue $b.OnBlue }
        foreach ($btn in @($ctx.OptionsButtons))  { $btn.Background = $b.LightOptBg; $btn.BorderBrush = $b.LightOptBorder; Set-ThemeForeground -Control $btn -IsDark $false -OnBlue $b.OnBlue }
        if ($ctx.LogLabel) { $ctx.LogLabel.Foreground = $b.LogLabelLight }
    }
    Add-SuiteImplicitStyle -Window $ctx.Window
    Set-ConnectionTextTheme -Window $ctx.Window -IsDark $IsDark
}

function Set-DialogTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Window]$Dialog,
        [Nullable[bool]]$IsDark
    )
    $ctx = Get-SuiteThemeContext
    if ($null -eq $IsDark) { $IsDark = [bool](& $ctx.IsDarkGetter) }
    Add-SuiteImplicitStyle -Window $Dialog
    if ($IsDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Dark.Steel')
        $Dialog.NonActiveWindowTitleBrush = $ctx.Brushes.DarkTitleBarInactive
    }
    else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Light.Blue')
        $Dialog.WindowTitleBrush          = $ctx.Brushes.TitleBarActive
        $Dialog.NonActiveWindowTitleBrush = $ctx.Brushes.TitleBarInactive
    }
    Register-SuiteTitleForeground -Window $Dialog -IsDark $IsDark
    Set-SuiteAccentButtonResource -Window $Dialog -IsDark $IsDark
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
            <!-- The one button style of the dialog, the same as ThemedButtonStyle of the apps: the neutral template, the theme
                 text color, no color literal. Every dialog button takes it or a style that is based on it. -->
            <Style x:Key="ThemedButtonStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Gray2}"/>
                <Setter Property="Template" Value="{DynamicResource SuiteNeutralButtonTemplate}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.ThemeForeground}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="HorizontalContentAlignment" Value="Center"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="Template" Value="{DynamicResource SuiteAccentButtonTemplate}"/>
                <Setter Property="Background" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.IdealForeground}"/>
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
            </Style>
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
            <Button x:Name="btnPrimary"   IsDefault="True" Style="{StaticResource DialogAccentButton}"/>
            <Button x:Name="btnSecondary" IsCancel="True" Visibility="Collapsed" Margin="0" Style="{StaticResource DialogButton}"/>
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
    # The caption text of the window state and the button templates of the dialog styles (safe before the theme is initialized).
    try { Register-SuiteTitleForeground -Window $dlg; Set-SuiteAccentButtonResource -Window $dlg } catch { $null = $_ }

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
            <!-- The one button style of the dialog, the same as ThemedButtonStyle of the apps: the neutral template, the theme
                 text color, no color literal. Every dialog button takes it or a style that is based on it. -->
            <Style x:Key="ThemedButtonStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Gray2}"/>
                <Setter Property="Template" Value="{DynamicResource SuiteNeutralButtonTemplate}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.ThemeForeground}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="HorizontalContentAlignment" Value="Center"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="Template" Value="{DynamicResource SuiteAccentButtonTemplate}"/>
                <Setter Property="Background" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.IdealForeground}"/>
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
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
    # The caption text of the window state and the button templates of the dialog styles (safe before the theme is initialized).
    try { Register-SuiteTitleForeground -Window $dlg; Set-SuiteAccentButtonResource -Window $dlg } catch { $null = $_ }
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
    # The new runspace reads PSModulePath from the process environment.
    [void](Repair-WindowsPowerShellModulePath)
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
    # A failed bootstrap (bad module path, import error) otherwise surfaces
    # much later as an unrelated "term not recognized" from the first
    # background operation; name the real cause here and hand back nothing
    # rather than a runspace whose every operation would fail.
    if ($initPS.HadErrors) {
        $messages = @($initPS.Streams.Error | ForEach-Object { $_.ToString() } | Select-Object -Unique)
        $detail = if ($messages.Count -gt 0) { $messages -join '; ' } else { 'Unknown initialization error' }
        Write-Log ("Background runspace initialization failed (module: {0}): {1}" -f $ModulePath, $detail) -Level ERROR
        $initPS.Dispose()
        try { $rs.Close() } catch { $null = $_ }
        try { $rs.Dispose() } catch { $null = $_ }
        throw "Background runspace initialization failed for '$ModulePath': $detail"
    }
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
        graveyard and reaped on a later pass once it has stopped.
        Returns the surviving graveyard; wrap the result in @() at the
        call site (an emptied graveyard returns nothing).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing.')]
    param(
        $PowerShell,
        $Timer,
        [array]$Graveyard = @(),
        $Handle,
        [switch]$Discard
    )
    # -Handle is the object of Start-SuiteBgWork. With -Discard the cancel flag is set (the worker stops before its next
    # read), no more item or completion callback runs, and the pipeline is stopped without blocking the UI thread.
    if ($Handle) {
        $Handle.Discarded = [bool]$Discard
        try { $Handle.State['Cancel'] = $true } catch { $null = $_ }
        try { $Handle.State['Queue'].Clear() } catch { $null = $_ }
        if (-not $PowerShell) { $PowerShell = $Handle.PowerShell }
        if (-not $Timer) { $Timer = $Handle.Timer }
    }
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
# Loading robustness (Q-47): retry, member chunks, background work with cancel
# =============================================================================

function Test-SuiteTransientFault {
    <#
    .SYNOPSIS
        True when an error looks like a network or provider hiccup that a retry can cure.
    .DESCRIPTION
        Timeouts, a lost connection, an unreachable RPC server or WinRM endpoint count.
        A bad query, a missing object, denied access and a cancel do not.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)
    $ex = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $ErrorRecord.Exception } else { $ErrorRecord }
    $text = ''
    for ($e = $ex; $e; $e = $e.InnerException) { $text += ' ' + $e.GetType().Name + ' ' + $e.Message }
    if ($text -match 'OperationCanceled|PipelineStopped') { return $false }
    return [bool]($text -match '(?i)timed? ?out|timeout|RPC server|0x800706B[AE]|0x80041032|0x800704C[0-9A-F]|temporarily|unavailable|connection (was )?(closed|reset|lost|forcibly|refused)|network (path|name|error)|WinRM cannot|cannot connect|transport|socket|name could not be resolved|server too busy|too many requests|The specified network|provider load failure|shutting down|quota violation|0x80041013|0x80041033|0x8004106C')
}

function Invoke-SuiteCMRead {
    <#
    .SYNOPSIS
        Runs a Configuration Manager read and retries it for transient faults only.
    .DESCRIPTION
        Up to MaxAttempts attempts, with DelaySeconds between them (2 s, then 6 s). Each failed attempt is logged.
        A fault that is not transient is not retried. When the attempts are used up, the error is thrown, or, when
        -Notes is given, one note is added to it and Default is returned, so that the rest of a load still shows.
        -Cancel is a hashtable with a Cancel key: when the key is true, the call stops with an OperationCanceledException.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Runs a read only; changes nothing.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$Description = 'Configuration Manager read',
        [int]$MaxAttempts = 3,
        [int[]]$DelaySeconds = @(2, 6),
        [scriptblock]$IsTransient = { param($er) Test-SuiteTransientFault -ErrorRecord $er },
        [System.Collections.IDictionary]$Cancel,
        [System.Collections.IList]$Notes,
        $Default = $null
    )
    for ($attempt = 1; ; $attempt++) {
        if ($Cancel -and $Cancel['Cancel']) { throw (New-Object System.OperationCanceledException("$Description was cancelled.")) }
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            $err = $_
            if ($err.Exception -is [System.OperationCanceledException]) { throw }
            $transient = [bool](& $IsTransient $err)
            if ($transient -and $attempt -lt $MaxAttempts) {
                $delay = if ($DelaySeconds.Count -gt 0) { $DelaySeconds[[Math]::Min($attempt - 1, $DelaySeconds.Count - 1)] } else { 0 }
                Write-Log ("{0}: attempt {1} of {2} failed ({3}); trying again in {4} s." -f $Description, $attempt, $MaxAttempts, $err.Exception.Message, $delay) -Level WARN
                $until = (Get-Date).AddSeconds($delay)
                while ((Get-Date) -lt $until) {
                    if ($Cancel -and $Cancel['Cancel']) { throw (New-Object System.OperationCanceledException("$Description was cancelled.")) }
                    Start-Sleep -Milliseconds 200
                }
                continue
            }
            $message = '{0} failed after {1} attempt(s): {2}' -f $Description, $attempt, $err.Exception.Message
            Write-Log $message -Level WARN
            if ($null -ne $Notes) { [void]$Notes.Add($message); return $Default }
            throw
        }
    }
}

function Get-SuiteCMWqlReader {
    # The default reader of the member helpers: one WQL query against the SMS Provider of the connected site.
    [CmdletBinding()]
    param([string]$ComputerName, [string]$Namespace, [int]$TimeoutSeconds = 120)
    if (-not $ComputerName -or -not $Namespace) {
        $info = Get-CMConnectionInfo
        if (-not $info) { throw 'Not connected to a site: give -ComputerName and -Namespace, or connect first.' }
        if (-not $ComputerName) { $ComputerName = $info.SMSProvider }
        if (-not $Namespace) { $Namespace = 'root\sms\site_' + $info.SiteCode }
    }
    # On the provider computer itself (localhost or .) the read goes without -ComputerName, so no remote logon is needed.
    if ($ComputerName -in 'localhost', '.') {
        return { param($Query) Get-CimInstance -Namespace $Namespace -Query $Query -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop }.GetNewClosure()
    }
    return { param($Query) Get-CimInstance -ComputerName $ComputerName -Namespace $Namespace -Query $Query -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop }.GetNewClosure()
}

function Test-SuiteCollectionId {
    param([string]$CollectionId)
    return [bool]($CollectionId -match '^[A-Za-z0-9]{8}$')
}

function Get-SuiteCollectionMemberPlan {
    <#
    .SYNOPSIS
        Reads only the ResourceIDs of a collection and cuts them into ResourceID ranges of ChunkSize members each.
    .DESCRIPTION
        One light read (one column) gives the exact boundaries, so each range returns ChunkSize rows whatever the
        spread of the IDs is. Returns Total and Ranges (From, To, Count).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [int]$ChunkSize = 2000,
        [scriptblock]$Reader
    )
    if (-not (Test-SuiteCollectionId $CollectionId)) { throw "Not a collection ID: '$CollectionId'." }
    if ($ChunkSize -lt 1) { throw 'ChunkSize must be at least 1.' }
    if (-not $Reader) { $Reader = Get-SuiteCMWqlReader }
    $rows = @(& $Reader ('SELECT ResourceID FROM SMS_CM_RES_COLL_{0}' -f $CollectionId))
    $ids = New-Object 'long[]' $rows.Count
    for ($i = 0; $i -lt $rows.Count; $i++) { $ids[$i] = [long]$rows[$i].ResourceID }
    [Array]::Sort($ids)
    $ranges = New-Object System.Collections.Generic.List[object]
    for ($start = 0; $start -lt $ids.Length; $start += $ChunkSize) {
        $end = [Math]::Min($start + $ChunkSize, $ids.Length) - 1
        $ranges.Add([pscustomobject]@{ From = $ids[$start]; To = $ids[$end]; Count = ($end - $start + 1) })
    }
    return [pscustomobject]@{ CollectionId = $CollectionId; Total = $ids.Length; Ranges = $ranges.ToArray() }
}

function Get-SuiteCollectionMemberChunk {
    <#
    .SYNOPSIS
        Reads the members of one ResourceID range of a collection, with only the columns that the grids show.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][long]$FromResourceId,
        [Parameter(Mandatory)][long]$ToResourceId,
        [scriptblock]$Reader
    )
    if (-not (Test-SuiteCollectionId $CollectionId)) { throw "Not a collection ID: '$CollectionId'." }
    if (-not $Reader) { $Reader = Get-SuiteCMWqlReader }
    $query = 'SELECT ResourceID, Name, Domain, IsClient, IsActive FROM SMS_CM_RES_COLL_{0} WHERE ResourceID >= {1} AND ResourceID <= {2}' -f $CollectionId, $FromResourceId, $ToResourceId
    foreach ($m in @(& $Reader $query)) {
        [pscustomobject]@{ Name = $m.Name; ResourceID = $m.ResourceID; Domain = $m.Domain; IsClient = $m.IsClient; IsActive = $m.IsActive }
    }
}

function Read-SuiteCollectionMember {
    <#
    .SYNOPSIS
        Worker for a member load. Runs in the background runspace and fills State.Queue with one array per chunk.
    .DESCRIPTION
        State is the hashtable of Start-SuiteBgWork (Queue, Cancel, Notes, Progress). Every read goes through
        Invoke-SuiteCMRead: a transient fault is retried; a chunk that still fails adds one note and the load goes on
        with the next chunk, so the grid keeps the partial data. The cancel flag is read before every chunk.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$State,
        [Parameter(Mandatory)][string]$CollectionId,
        [int]$ChunkSize = 2000,
        [scriptblock]$Reader,
        [int[]]$DelaySeconds = @(2, 6),
        [string]$ComputerName,
        [string]$Namespace
    )
    # A fresh background runspace has no connection record: the shell passes the provider and the namespace.
    if (-not $Reader) { $Reader = Get-SuiteCMWqlReader -ComputerName $ComputerName -Namespace $Namespace }
    $plan = Invoke-SuiteCMRead -Description "Member IDs of $CollectionId" -Cancel $State -DelaySeconds $DelaySeconds -ScriptBlock {
        Get-SuiteCollectionMemberPlan -CollectionId $CollectionId -ChunkSize $ChunkSize -Reader $Reader
    }
    $State['Total'] = $plan.Total
    $loaded = 0
    foreach ($range in $plan.Ranges) {
        if ($State['Cancel']) { throw (New-Object System.OperationCanceledException("The member load of $CollectionId was cancelled.")) }
        $rows = @(Invoke-SuiteCMRead -Description ("Members {0}-{1} of {2}" -f $range.From, $range.To, $CollectionId) -Cancel $State -DelaySeconds $DelaySeconds -Notes $State['Notes'] -Default @() -ScriptBlock {
                Get-SuiteCollectionMemberChunk -CollectionId $CollectionId -FromResourceId $range.From -ToResourceId $range.To -Reader $Reader
            })
        $loaded += $rows.Count
        $State['Loaded'] = $loaded
        $State['Progress'] = 'Loaded {0} of {1} members' -f $loaded, $plan.Total
        if ($rows.Count -gt 0) { $State['Queue'].Enqueue($rows) }
    }
}

function Start-SuiteBgWork {
    <#
    .SYNOPSIS
        Runs a script in the background runspace and hands its output to the UI thread in pieces.
    .DESCRIPTION
        ScriptBlock gets the shared State hashtable first, then ArgumentList. It runs in the runspace, so it cannot use
        variables of the caller (it is passed as text). State has Queue (items for OnItem), Notes, Progress, Cancel,
        Done and Error. A timer on the UI thread calls Receive-SuiteBgWork: OnItem runs for each queued item (at most
        MaxItemsPerTick per tick), OnDone runs once when the work has ended and the queue is empty. No Configuration
        Manager call runs on the UI thread. Stop it with Stop-SuiteBgWork -Handle $h -Discard.
        Pass -NoTimer to drive Receive-SuiteBgWork yourself (tests).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Starts ephemeral background plumbing only.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Runspace,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [scriptblock]$OnItem,
        [scriptblock]$OnDone,
        [int]$IntervalMs = 100,
        [int]$MaxItemsPerTick = 4,
        [switch]$NoTimer
    )
    $state = [hashtable]::Synchronized(@{
            Queue    = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
            Notes    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
            Progress = ''
            Cancel   = $false
            Done     = $false
            Error    = $null
            Total    = 0
            Loaded   = 0
        })
    $ps = [powershell]::Create()
    $ps.Runspace = $Runspace
    [void]$ps.AddScript({
            param($State, $ScriptText, $Arguments)
            try { & ([scriptblock]::Create($ScriptText)) $State @Arguments }
            catch { if ($_.Exception -isnot [System.OperationCanceledException]) { $State['Error'] = $_.Exception.Message } else { $State['Cancelled'] = $true } }
            finally { $State['Done'] = $true }
        }).AddArgument($state).AddArgument($ScriptBlock.ToString()).AddArgument($ArgumentList)
    $handle = [pscustomobject]@{
        State = $state; PowerShell = $ps; Async = $null; Timer = $null; OnItem = $OnItem; OnDone = $OnDone
        MaxItemsPerTick = $MaxItemsPerTick; Finished = $false; Discarded = $false
    }
    $handle.Async = $ps.BeginInvoke()
    if (-not $NoTimer) {
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds($IntervalMs)
        $timer.Tag = $handle
        $timer.Add_Tick({ Receive-SuiteBgWork -Handle $this.Tag })
        $handle.Timer = $timer
        $timer.Start()
    }
    return $handle
}

function Receive-SuiteBgWork {
    <#
    .SYNOPSIS
        One tick of a Start-SuiteBgWork handle: passes queued items to OnItem, and calls OnDone at the end.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Handle)
    if ($Handle.Finished -or $Handle.Discarded) { return $true }
    $state = $Handle.State
    $count = 0
    while ($count -lt $Handle.MaxItemsPerTick -and $state['Queue'].Count -gt 0) {
        $item = $state['Queue'].Dequeue()
        $count++
        if ($Handle.OnItem) { try { & $Handle.OnItem $item $state } catch { Write-Log ("Background work item failed: {0}" -f $_.Exception.Message) -Level WARN } }
    }
    if ($state['Done'] -and $state['Queue'].Count -eq 0) {
        $Handle.Finished = $true
        if ($Handle.Timer) { try { $Handle.Timer.Stop() } catch { $null = $_ } }
        try { [void]$Handle.PowerShell.EndInvoke($Handle.Async) } catch { $null = $_ }
        try { $Handle.PowerShell.Dispose() } catch { $null = $_ }
        if ($Handle.OnDone) { try { & $Handle.OnDone $state } catch { Write-Log ("Background work completion failed: {0}" -f $_.Exception.Message) -Level WARN } }
        return $true
    }
    return $false
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
            <!-- The one button style of the dialog, the same as ThemedButtonStyle of the apps: the neutral template, the theme
                 text color, no color literal. Every dialog button takes it or a style that is based on it. -->
            <Style x:Key="ThemedButtonStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Gray2}"/>
                <Setter Property="Template" Value="{DynamicResource SuiteNeutralButtonTemplate}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.ThemeForeground}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="HorizontalContentAlignment" Value="Center"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}">
                <Setter Property="Template" Value="{DynamicResource SuiteAccentButtonTemplate}"/>
                <Setter Property="Background" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.AccentBase}"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.IdealForeground}"/>
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
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

    # The highlighted collection is a candidate only. It becomes the result
    # when ShowDialog returns $true; the Cancel button, the close button and
    # Alt+F4 all leave DialogResult unset or $false.
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
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    $accepted = $dlg.ShowDialog()
    $picked = $script:PickerResult
    $script:PickerResult = $null
    if ($accepted -eq $true) { return $picked }
    return $null
}

# =============================================================================
# Suite update check and self-update
# =============================================================================
# One cache per machine user under %LOCALAPPDATA%\AppPackagerSuite, shared by
# the launcher and every tool, so a day's first launch of any tool performs
# the single release lookup. The installer is the signed per-user setup from
# the suite release; it is verified against checksums.txt and its
# Authenticode signature before it runs.

$script:SuiteUpdateRepo      = 'jasonulbright/app-packager-suite'
$script:SuiteUpdateUserAgent = 'AppPackagerSuite-UpdateCheck'
$script:SuiteUpdateSigner    = 'CN=Jason Ulbright'
$script:SuiteUpdateHours     = 24

function Get-SuiteInstallInfo {
    <#
    .SYNOPSIS
        Finds the suite install a tool or the launcher runs from.
    .DESCRIPTION
        The suite installer writes suite-manifest.json at the install root;
        each component sits in a folder below it. Path is the caller's own
        folder; the manifest is looked for there and one level up.
    .OUTPUTS
        [pscustomobject] Installed, Root, Version, ManifestPath.
    #>
    param([Parameter(Mandatory)][string]$Path)
    foreach ($dir in @($Path, (Split-Path -Parent $Path))) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        $manifest = Join-Path $dir 'suite-manifest.json'
        if (Test-Path -LiteralPath $manifest) {
            $version = ''
            try { $version = [string](Get-Content -LiteralPath $manifest -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop).version } catch { $null = $_ }
            return [pscustomobject]@{ Installed = $true; Root = $dir; Version = $version; ManifestPath = $manifest }
        }
    }
    return [pscustomobject]@{ Installed = $false; Root = ''; Version = ''; ManifestPath = '' }
}

function Get-SuiteUpdateCachePath {
    Join-Path (Join-Path $env:LOCALAPPDATA 'AppPackagerSuite') 'update-check.json'
}

function Read-SuiteUpdateCache {
    $path = Get-SuiteUpdateCachePath
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) }
    catch { return $null }
}

function Save-SuiteUpdateCache {
    param([AllowNull()][string]$LatestVersion, [AllowNull()][string]$ReleaseUrl)
    try {
        $path = Get-SuiteUpdateCachePath
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [pscustomobject]@{
            LastCheckUtc  = (Get-Date).ToUniversalTime().ToString('o')
            LatestVersion = $LatestVersion
            ReleaseUrl    = $ReleaseUrl
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $path -Encoding UTF8
    }
    catch { $null = $_ }
}

function Test-SuiteUpdateDue {
    param([AllowNull()][string]$LastCheckUtc, [int]$IntervalHours = $script:SuiteUpdateHours)
    if ([string]::IsNullOrWhiteSpace($LastCheckUtc)) { return $true }
    $last = [datetime]::MinValue
    if (-not [datetime]::TryParse($LastCheckUtc, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$last)) { return $true }
    return (((Get-Date).ToUniversalTime() - $last.ToUniversalTime()).TotalHours -ge $IntervalHours)
}

function ConvertFrom-SuiteReleaseTag {
    # 'v2026.10.11.0051' -> '2026.10.11.0051'; anything else -> ''.
    param([AllowNull()][string]$Tag)
    if ([string]::IsNullOrWhiteSpace($Tag)) { return '' }
    $t = $Tag.Trim()
    if ($t -match '^v?(\d{4}\.\d{2}\.\d{2}\.\d{4})$') { return $Matches[1] }
    return ''
}

function Compare-SuiteVersion {
    # -1, 0, 1 for A against B. Parts compare as numbers, so zero padding
    # and a dropped leading zero compare equal. An unparsable side counts
    # as the lowest.
    param([AllowNull()][string]$A, [AllowNull()][string]$B)
    $pa = @(); $pb = @()
    if ($A -match '^\d+(\.\d+)*$') { $pa = @($A -split '\.' | ForEach-Object { [long]$_ }) }
    if ($B -match '^\d+(\.\d+)*$') { $pb = @($B -split '\.' | ForEach-Object { [long]$_ }) }
    if ($pa.Count -eq 0 -and $pb.Count -eq 0) { return 0 }
    if ($pa.Count -eq 0) { return -1 }
    if ($pb.Count -eq 0) { return 1 }
    $n = [Math]::Max($pa.Count, $pb.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $x = $(if ($i -lt $pa.Count) { $pa[$i] } else { 0 })
        $y = $(if ($i -lt $pb.Count) { $pb[$i] } else { 0 })
        if ($x -lt $y) { return -1 }
        if ($x -gt $y) { return 1 }
    }
    return 0
}

function Invoke-SuiteHttpGet {
    <#
    .SYNOPSIS
        GET a URL to a string or a file. curl.exe first (SSL-inspecting
        proxies that break the .NET chain pass curl); Invoke-WebRequest
        when curl.exe is missing. Throws on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [AllowEmptyString()][string]$OutFile = '',
        [int]$TimeoutSeconds = 60
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $curl = Get-Command -Name curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        $curlArgs = @('-L', '--fail', '--silent', '--show-error', '--max-time', $TimeoutSeconds, '-A', $script:SuiteUpdateUserAgent)
        if ($OutFile) { $curlArgs += @('-o', $OutFile) }
        $curlArgs += $Url
        $out = & $curl.Source @curlArgs 2>&1
        if ($LASTEXITCODE -ne 0) { throw ('Download failed (curl exit {0}): {1}' -f $LASTEXITCODE, $Url) }
        if ($OutFile) { return $OutFile }
        return (($out | ForEach-Object { [string]$_ }) -join "`n")
    }
    if ($OutFile) {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -TimeoutSec $TimeoutSeconds -Headers @{ 'User-Agent' = $script:SuiteUpdateUserAgent } -ErrorAction Stop
        return $OutFile
    }
    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSeconds -Headers @{ 'User-Agent' = $script:SuiteUpdateUserAgent } -ErrorAction Stop
    return [string]$resp.Content
}

function Get-SuiteLatestRelease {
    <#
    .SYNOPSIS
        The latest suite release on GitHub: Version, Tag, Url.
    #>
    param([string]$Repo = $script:SuiteUpdateRepo)
    $json = Invoke-SuiteHttpGet -Url ('https://api.github.com/repos/{0}/releases/latest' -f $Repo) -TimeoutSeconds 20
    $release = $json | ConvertFrom-Json
    [pscustomobject]@{
        Tag     = [string]$release.tag_name
        Version = (ConvertFrom-SuiteReleaseTag -Tag ([string]$release.tag_name))
        Url     = [string]$release.html_url
    }
}

function Get-SuiteUpdateState {
    <#
    .SYNOPSIS
        Whether a newer suite release exists than the installed version.
    .DESCRIPTION
        Uses the shared cache when the last lookup is younger than the
        interval; otherwise looks the latest release up and refreshes the
        cache. A failed lookup never throws: the result carries Error and
        the cached answer when one exists.
    .PARAMETER InstalledVersion
        The suite version to compare against (suite-manifest.json version).
    .OUTPUTS
        [pscustomobject] Current, Latest, Available, ReleaseUrl, Checked
        (true when the lookup ran now), Error.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$InstalledVersion,
        [switch]$Force,
        [int]$IntervalHours = $script:SuiteUpdateHours
    )
    $state = [pscustomobject]@{ Current = $InstalledVersion; Latest = ''; Available = $false; ReleaseUrl = ''; Checked = $false; Error = '' }
    $cache = Read-SuiteUpdateCache
    $lastCheck = ''
    if ($cache) { $lastCheck = [string]$cache.LastCheckUtc }
    $due = $Force -or -not $cache -or (Test-SuiteUpdateDue -LastCheckUtc $lastCheck -IntervalHours $IntervalHours)
    if ($due) {
        try {
            $release = Get-SuiteLatestRelease
            if ($release.Version) {
                Save-SuiteUpdateCache -LatestVersion $release.Version -ReleaseUrl $release.Url
                $state.Latest = $release.Version
                $state.ReleaseUrl = $release.Url
                $state.Checked = $true
            }
            else { $state.Error = ('Release tag has no suite version: {0}' -f $release.Tag) }
        }
        catch { $state.Error = $_.Exception.Message }
    }
    if (-not $state.Latest -and $cache -and $cache.LatestVersion) {
        $state.Latest = [string]$cache.LatestVersion
        $state.ReleaseUrl = [string]$cache.ReleaseUrl
    }
    if ($state.Latest -and $InstalledVersion) {
        $state.Available = ((Compare-SuiteVersion -A $InstalledVersion -B $state.Latest) -lt 0)
    }
    return $state
}

function Get-SuiteUpdateDownload {
    <#
    .SYNOPSIS
        Downloads SuiteSetup-<Version>.exe and checksums.txt from the suite
        release and verifies the hash and the Authenticode signature.
    .OUTPUTS
        The installer path. Throws when the download, the hash or the
        signature fails; a failed file is deleted.
    #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [string]$Repo = $script:SuiteUpdateRepo,
        [string]$Folder = (Join-Path ([IO.Path]::GetTempPath()) ('suite-update-' + [guid]::NewGuid().ToString('N')))
    )
    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $name = 'SuiteSetup-{0}.exe' -f $Version
    $base = 'https://github.com/{0}/releases/download/v{1}/' -f $Repo, $Version
    $exe = Join-Path $Folder $name
    $sums = Join-Path $Folder 'checksums.txt'
    try {
        $null = Invoke-SuiteHttpGet -Url ($base + $name) -OutFile $exe -TimeoutSeconds 300
        $null = Invoke-SuiteHttpGet -Url ($base + 'checksums.txt') -OutFile $sums -TimeoutSeconds 60
        $expected = ''
        foreach ($line in (Get-Content -LiteralPath $sums)) {
            if ($line -match ('^([0-9a-fA-F]{64})\s+\*?' + [regex]::Escape($name) + '\s*$')) { $expected = $Matches[1].ToLowerInvariant(); break }
        }
        if (-not $expected) { throw ('checksums.txt has no entry for {0}' -f $name) }
        $actual = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actual -ne $expected) { throw ('SHA256 mismatch for {0}' -f $name) }
        $sig = Get-AuthenticodeSignature -LiteralPath $exe
        if ($sig.Status -ne 'Valid') { throw ('Signature of {0} is {1}' -f $name, $sig.Status) }
        $subject = ''
        if ($sig.SignerCertificate) { $subject = [string]$sig.SignerCertificate.Subject }
        if (-not $subject.StartsWith($script:SuiteUpdateSigner + ',') -and $subject -ne $script:SuiteUpdateSigner) {
            throw ('Signer of {0} is {1}, not {2}' -f $name, $subject, $script:SuiteUpdateSigner)
        }
        return $exe
    }
    catch {
        Remove-Item -LiteralPath $exe -Force -ErrorAction SilentlyContinue
        throw
    }
}

function ConvertTo-SuiteNativeArgument {
    # Start-Process joins -ArgumentList elements with single spaces and adds no
    # quotes, so a path holding a space reaches the child as several arguments.
    # This applies the CommandLineToArgvW rules: backslashes double only in
    # front of a quote, and a trailing run doubles before the closing quote.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    $slash = [char]92
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]'\') { $backslashes++; continue }
        if ($ch -eq [char]'"') { [void]$builder.Append($slash, ($backslashes * 2 + 1)) }
        else { [void]$builder.Append($slash, $backslashes) }
        $backslashes = 0
        [void]$builder.Append($ch)
    }
    [void]$builder.Append($slash, ($backslashes * 2))
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Start-SuiteUpdateInstaller {
    <#
    .SYNOPSIS
        Runs the verified installer silently after the calling process has
        exited, then relaunches the given script. The caller closes its
        window right after this returns.
    .DESCRIPTION
        The installer replaces the folder the caller runs from, so a
        detached child waits on the caller's process id (bounded), runs the
        setup with /S, waits for it, and starts the relaunch script.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [int]$WaitForProcessId = $PID,
        [AllowEmptyString()][string]$Relaunch = ''
    )
    if (-not (Test-Path -LiteralPath $InstallerPath)) { throw ('Installer not found: {0}' -f $InstallerPath) }
    $quotedInstaller = $InstallerPath.Replace("'", "''")
    $lines = @(
        ('Wait-Process -Id {0} -Timeout 120 -ErrorAction SilentlyContinue' -f $WaitForProcessId)
        ('$p = Start-Process -FilePath ' + "'" + $quotedInstaller + "'" + ' -ArgumentList ' + "'/S'" + ' -PassThru -Wait')
    )
    if ($Relaunch) {
        $quotedRelaunch = (ConvertTo-SuiteNativeArgument -Value $Relaunch).Replace("'", "''")
        $lines += ('if ($p.ExitCode -eq 0) { Start-Process -FilePath ' + "'powershell.exe'" + ' -ArgumentList ' + "'-NoProfile','-ExecutionPolicy','Bypass','-File','" + $quotedRelaunch + "'" + ' }')
    }
    $bootstrap = Join-Path ([IO.Path]::GetTempPath()) ('suite-update-' + [guid]::NewGuid().ToString('N') + '.ps1')
    Set-Content -LiteralPath $bootstrap -Value ($lines -join "`r`n") -Encoding UTF8
    Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', (ConvertTo-SuiteNativeArgument -Value $bootstrap) | Out-Null
    return $bootstrap
}

function Start-SuiteUpdateNotice {
    <#
    .SYNOPSIS
        Runs the suite update check off the UI thread and hands the result
        to OnResult on the UI thread. One call per tool at window load.
    .DESCRIPTION
        Nothing happens for a plain checkout (no suite-manifest.json). The
        lookup runs in its own runspace with this module imported; a
        DispatcherTimer polls for the result and invokes OnResult with the
        Get-SuiteUpdateState object. Every failure is swallowed: a missing
        network must never touch a tool's launch.
    .PARAMETER Path
        The caller's folder ($PSScriptRoot); the manifest is looked for
        there and one level up.
    .PARAMETER OnResult
        Scriptblock receiving the state object. Called only when the check
        ran or a cached answer exists.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$OnResult,
        [switch]$Force
    )
    try {
        $install = Get-SuiteInstallInfo -Path $Path
        if (-not $install.Installed) { return $null }
        $state = [hashtable]::Synchronized(@{ Done = $false; Result = $null })
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        $modulePath = Join-Path $PSScriptRoot 'SuiteCommon.psd1'
        [void]$ps.AddScript({
            param($State, $ModulePath, $Installed, $ForceCheck)
            try {
                Import-Module $ModulePath -Force -DisableNameChecking
                $State.Result = Get-SuiteUpdateState -InstalledVersion $Installed -Force:$ForceCheck
            }
            catch { $State.Result = $null }
            finally { $State.Done = $true }
        }).AddArgument($state).AddArgument($modulePath).AddArgument([string]$install.Version).AddArgument([bool]$Force)
        $handle = $ps.BeginInvoke()
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(500)
        $timer.Add_Tick({
            if (-not $state.Done) { return }
            $timer.Stop()
            try { if ($state.Result) { & $OnResult $state.Result } } catch { $null = $_ }
            try { $ps.EndInvoke($handle) } catch { $null = $_ }
            try { $ps.Dispose() } catch { $null = $_ }
            try { $rs.Close() } catch { $null = $_ }
        }.GetNewClosure())
        $timer.Start()
        return $timer
    }
    catch { return $null }
}

#region Window test helpers (measurement only; no assertions)
# These functions are for window tests that run in an STA runspace with a shown window.
# The state report needs the real mouse, so it belongs to a visible run (SUITE_VISIBLE_WINDOWS=1).

function Get-SuiteContrastRatio {
    <#
    .SYNOPSIS
        WCAG contrast ratio of two colors (1 to 21).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Windows.Media.Color]$Foreground,
        [Parameter(Mandatory)][System.Windows.Media.Color]$Background
    )
    $lum = {
        param($c)
        $f = { param($v) $v = $v / 255.0; if ($v -le 0.04045) { $v / 12.92 } else { [math]::Pow(($v + 0.055) / 1.055, 2.4) } }
        0.2126 * (& $f $c.R) + 0.7152 * (& $f $c.G) + 0.0722 * (& $f $c.B)
    }
    $l1 = & $lum $Foreground
    $l2 = & $lum $Background
    if ($l1 -lt $l2) { $t = $l1; $l1 = $l2; $l2 = $t }
    return (($l1 + 0.05) / ($l2 + 0.05))
}

function Get-SuiteVisualDescendant {
    <#
    .SYNOPSIS
        All visual descendants of an element, depth first.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Element)
    $list = New-Object System.Collections.Generic.List[object]
    $walk = {
        param($e)
        $n = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($e)
        for ($i = 0; $i -lt $n; $i++) {
            $c = [System.Windows.Media.VisualTreeHelper]::GetChild($e, $i)
            $list.Add($c)
            & $walk $c
        }
    }
    & $walk $Element
    return $list.ToArray()
}

function Test-SuiteVisibleInTree {
    <#
    .SYNOPSIS
        True when an element and all of its ancestors are visible.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Element)
    $e = $Element
    while ($e) {
        if ($e -is [System.Windows.UIElement] -and $e.Visibility -ne [System.Windows.Visibility]::Visible) { return $false }
        $e = [System.Windows.Media.VisualTreeHelper]::GetParent($e)
    }
    return $true
}

function Get-SuiteButtonRow {
    <#
    .SYNOPSIS
        One row per visible text button of a window: name, text, font weight, border thickness and border/foreground colors.
    .DESCRIPTION
        Window command buttons (minimize, maximize, close), toggle buttons and the template buttons of other controls are skipped.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [string]$Label = '')
    $hex = { param($b) if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Color.ToString() } else { '' } }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($d in (Get-SuiteVisualDescendant -Element $Window)) {
        if ($d -isnot [System.Windows.Controls.Primitives.ButtonBase] -or $d -is [System.Windows.Controls.Primitives.ToggleButton]) { continue }
        if ($d.GetType().Name -ne 'Button') { continue }
        if ($d.TemplatedParent -and $d.TemplatedParent.GetType().Name -ne 'Button') { continue }
        if (-not (Test-SuiteVisibleInTree -Element $d)) { continue }
        $content = if ($d.Content -is [string]) { $d.Content } else { $d.Name }
        if (-not $content) { continue }
        $rows.Add([pscustomobject]@{
                Window          = $Label
                Name            = $d.Name
                Text            = [string]$content
                Weight          = $d.FontWeight.ToOpenTypeWeight()
                BorderThickness = [math]::Max([math]::Max($d.BorderThickness.Left, $d.BorderThickness.Right), [math]::Max($d.BorderThickness.Top, $d.BorderThickness.Bottom))
                BorderColor     = (& $hex $d.BorderBrush)
                ForegroundColor = (& $hex $d.Foreground)
            })
    }
    return $rows.ToArray()
}

function Get-SuiteWindowChrome {
    <#
    .SYNOPSIS
        Title bar brushes, title text contrast and button rows of a shown MetroWindow.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [string]$Label = '')
    $hex = { param($b) if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Color.ToString() } else { '' } }
    $act = $Window.WindowTitleBrush
    # The bar in effect: the active brush while the window is active, the inactive brush otherwise. The caption text is measured on it.
    $now = if ($Window.IsActive) { $act } else { $Window.NonActiveWindowTitleBrush }
    $text = if ($Window.TitleForeground -is [System.Windows.Media.SolidColorBrush]) { $Window.TitleForeground.Color } else { [System.Windows.Media.Colors]::White }
    $ratio = 0.0
    if ($now -is [System.Windows.Media.SolidColorBrush]) { $ratio = [math]::Round((Get-SuiteContrastRatio -Foreground $text -Background $now.Color), 2) }
    $themeText = ''
    try { $themeText = & $hex $Window.FindResource('MahApps.Brushes.ThemeForeground') } catch { $null = $_ }
    return [pscustomobject]@{
        Title      = [string]$Window.Title
        Label      = $Label
        Active     = (& $hex $act)
        Inactive   = (& $hex $Window.NonActiveWindowTitleBrush)
        IsActive   = [bool]$Window.IsActive
        TitleBrushNow = (& $hex $now)
        TitleText  = $text.ToString()
        TitleRatio = $ratio
        ThemeText  = $themeText
        Buttons    = @(Get-SuiteButtonRow -Window $Window -Label $Label)
    }
}

function Get-SuiteRelativeLuminance {
    <#
    .SYNOPSIS
        WCAG relative luminance of a color (0 black to 1 white).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Media.Color]$Color)
    $f = { param($v) $v = $v / 255.0; if ($v -le 0.04045) { $v / 12.92 } else { [math]::Pow(($v + 0.055) / 1.055, 2.4) } }
    return (0.2126 * (& $f $Color.R) + 0.7152 * (& $f $Color.G) + 0.0722 * (& $f $Color.B))
}

function Get-SuiteWindowBody {
    <#
    .SYNOPSIS
        Body colors of a shown window: surface, text, and the input controls (text boxes, combo boxes, lists, grids).
    .DESCRIPTION
        BackLuminance tells a light surface from a dark one. TextRatio is the contrast of the window text on the surface.
        Each input row has the luminance of its background, the contrast of its text on that background and the contrast of
        its border on the surface. A dark dialog body in the light theme (or the other way round) shows in BackLuminance.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [string]$Label = '')
    $solid = { param($b) if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Color } else { $null } }
    $over = {
        param($c, $bg)
        $a = $c.A / 255.0
        [System.Windows.Media.Color]::FromRgb([byte]($c.R * $a + $bg.R * (1 - $a)), [byte]($c.G * $a + $bg.G * (1 - $a)), [byte]($c.B * $a + $bg.B * (1 - $a)))
    }
    $back = & $solid $Window.Background
    if (-not $back) { $back = [System.Windows.Media.Colors]::Transparent }
    $text = & $solid $Window.Foreground
    $rows = New-Object System.Collections.Generic.List[object]
    $kinds = 'TextBox', 'PasswordBox', 'ComboBox', 'ListBox', 'ListView', 'DataGrid', 'TreeView', 'DatePicker'
    foreach ($d in (Get-SuiteVisualDescendant -Element $Window)) {
        $kind = $d.GetType().Name
        if ($kind -notin $kinds -or -not (Test-SuiteVisibleInTree -Element $d) -or $d.ActualWidth -lt 10) { continue }
        $bg = & $solid $d.Background
        $fg = & $solid $d.Foreground
        $bd = & $solid $d.BorderBrush
        $bgOpaque = if ($bg -and $back) { & $over $bg $back } else { $back }
        $row = [ordered]@{ Kind = $kind; Name = [string]$d.Name; BackLuminance = $null; TextRatio = $null; BorderRatio = $null }
        if ($bgOpaque) { $row.BackLuminance = [math]::Round((Get-SuiteRelativeLuminance -Color $bgOpaque), 3) }
        if ($bgOpaque -and $fg) { $row.TextRatio = [math]::Round((Get-SuiteContrastRatio -Foreground (& $over $fg $bgOpaque) -Background $bgOpaque), 2) }
        if ($bd -and $back -and $d.BorderThickness.Left -gt 0) { $row.BorderRatio = [math]::Round((Get-SuiteContrastRatio -Foreground (& $over $bd $back) -Background $back), 2) }
        $rows.Add([pscustomobject]$row)
    }
    return [pscustomobject]@{
        Title         = [string]$Window.Title
        Label         = $Label
        BackLuminance = $(if ($back.A -gt 0) { [math]::Round((Get-SuiteRelativeLuminance -Color $back), 3) } else { $null })
        TextRatio     = $(if ($back.A -gt 0 -and $text) { [math]::Round((Get-SuiteContrastRatio -Foreground $text -Background $back), 2) } else { $null })
        Inputs        = $rows.ToArray()
    }
}
function Close-SuiteOpenWindow {
    <#
    .SYNOPSIS
        Closes every WPF window that is still open in this runspace; returns the titles of the windows it closed.
    .DESCRIPTION
        For window tests: call it in a finally block of the test (or after each STA call), so that no window stays on the
        desktop of the lab computer. It finds the windows through the presentation sources, so it needs no Application object.
        Run it on the thread that owns the windows (the STA thread).
    #>
    [CmdletBinding()]
    param()
    $closed = New-Object System.Collections.Generic.List[string]
    foreach ($source in @([System.Windows.PresentationSource]::CurrentSources)) {
        $w = $source.RootVisual
        if ($w -is [System.Windows.Window]) {
            $closed.Add([string]$w.Title)
            try { $w.Close() } catch { $null = $_ }
        }
    }
    return $closed.ToArray()
}
function Get-SuiteTitleBarContrast {
    <#
    .SYNOPSIS
        Contrast of the caption text on the active and on the inactive title bar, each with the text of its own state.
    .DESCRIPTION
        Active: white (library text) on the active bar. Inactive: in the light theme the dark caption text
        (TitleBarInactiveText) on the light blue bar, in the dark theme white on the gray bar. A window that is
        active or inactive at the time of the call gives the same answer, because the text of each state is taken
        from the palette, not from the window.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Window, [Nullable[bool]]$IsDark)
    $ctx = Get-SuiteThemeContext
    $dark = if ($null -ne $IsDark) { [bool]$IsDark } else { [bool](& $ctx.IsDarkGetter) }
    $white = [System.Windows.Media.Colors]::White
    $inactiveText = if ($dark) { $white } else { $ctx.Brushes.TitleBarInactiveText.Color }
    $result = [ordered]@{ Active = $null; Inactive = $null }
    $states = @(@{ Name = 'Active'; Brush = $Window.WindowTitleBrush; Text = $white }, @{ Name = 'Inactive'; Brush = $Window.NonActiveWindowTitleBrush; Text = $inactiveText })
    foreach ($s in $states) {
        if ($s.Brush -isnot [System.Windows.Media.SolidColorBrush]) { continue }
        $result[$s.Name] = [math]::Round((Get-SuiteContrastRatio -Foreground $s.Text -Background $s.Brush.Color), 2)
    }
    return [pscustomobject]$result
}
function Get-SuiteButtonStateReport {
    <#
    .SYNOPSIS
        Measures the normal, focus, hover and pressed states of the buttons of a shown window with the real mouse.
    .DESCRIPTION
        Samples the visible, enabled buttons (one per distinct style, at most -Max). For each state it reads the colors
        that the template shows: Text is the contrast of the button text on the button fill (the backdrop where the fill is
        transparent) and Edge is the larger contrast of the border or fill on the backdrop. The mouse is parked at the
        bottom right corner of the screen, and the mouse button goes up while the pointer is off the button, so the pointer does not trigger
        the button. Hover, Pressed and Focused tell whether the state was reached. Needs the visible labadmin session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Window,
        [bool]$Dark = $false,
        [int]$Max = 6
    )
    if (-not ('SuiteStateUi' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class SuiteStateUi { [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e); }'
    }
    $w = $Window
    $pump = {
        param($win, $n)
        for ($i = 0; $i -lt $n; $i++) {
            $win.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
            Start-Sleep -Milliseconds 40
        }
    }
    $solid = { param($b) if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Color } else { $null } }
    $over = {
        param($c, $bg)
        $a = $c.A / 255.0
        [System.Windows.Media.Color]::FromRgb([byte]($c.R * $a + $bg.R * (1 - $a)), [byte]($c.G * $a + $bg.G * (1 - $a)), [byte]($c.B * $a + $bg.B * (1 - $a)))
    }
    $winBg = & $solid $w.Background
    if (-not $winBg) { $winBg = $(if ($Dark) { [System.Windows.Media.Colors]::Black } else { [System.Windows.Media.Colors]::White }) }
    $backdropOf = {
        param($el)
        $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($el)
        while ($parent) {
            $b = $null
            if ($parent -is [System.Windows.Controls.Border]) { $b = $parent.Background }
            elseif ($parent -is [System.Windows.Controls.Panel]) { $b = $parent.Background }
            elseif ($parent -is [System.Windows.Controls.Control]) { $b = $parent.Background }
            $c = & $solid $b
            if ($c -and $c.A -eq 255) { return $c }
            $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($parent)
        }
        return $winBg
    }
    $readState = {
        param($btn, $backdrop)
        $cp = $null; $fill = $null; $border = $null
        foreach ($d in (Get-SuiteVisualDescendant -Element $btn)) {
            if (-not $cp -and $d -is [System.Windows.Controls.ContentPresenter]) { $cp = $d }
            $bgBrush = $null
            if ($d -is [System.Windows.Controls.Border]) { $bgBrush = $d.Background } elseif ($d -is [System.Windows.Controls.Panel]) { $bgBrush = $d.Background }
            if (-not $fill -and $bgBrush) { $c = & $solid $bgBrush; if ($c -and $c.A -gt 0) { $fill = $c } }
            if (-not $border -and $d -is [System.Windows.Controls.Border] -and $d.BorderThickness.Left -gt 0) { $c = & $solid $d.BorderBrush; if ($c -and $c.A -gt 0) { $border = $c } }
        }
        if (-not $fill) { $c = & $solid $btn.Background; if ($c -and $c.A -gt 0) { $fill = $c } }
        if (-not $border -and $btn.BorderThickness.Left -gt 0) { $c = & $solid $btn.BorderBrush; if ($c -and $c.A -gt 0) { $border = $c } }
        $fillOpaque = if ($fill) { & $over $fill $backdrop } else { $backdrop }
        $text = & $solid ([System.Windows.Documents.TextElement]::GetForeground($(if ($cp) { $cp } else { $btn })))
        if (-not $text) { $text = & $solid $btn.Foreground }
        $textOpaque = & $over $text $fillOpaque
        $edge = 0.0
        if ($border) { $edge = [math]::Max($edge, (Get-SuiteContrastRatio -Foreground (& $over $border $backdrop) -Background $backdrop)) }
        if ($fill) { $edge = [math]::Max($edge, (Get-SuiteContrastRatio -Foreground $fillOpaque -Background $backdrop)) }
        [pscustomobject]@{ Text = [math]::Round((Get-SuiteContrastRatio -Foreground $textOpaque -Background $fillOpaque), 2); Edge = [math]::Round($edge, 2); Fill = $fillOpaque.ToString(); FillLuminance = [math]::Round((Get-SuiteRelativeLuminance -Color $fillOpaque), 4); Blue = (($fillOpaque.B - $fillOpaque.R) -ge 20) }
    }
    $seen = @{}
    $sample = New-Object System.Collections.Generic.List[object]
    foreach ($d in (Get-SuiteVisualDescendant -Element $w)) {
        if ($d.GetType().Name -ne 'Button' -or -not $d.IsEnabled -or -not (Test-SuiteVisibleInTree -Element $d) -or $d.ActualWidth -lt 20) { continue }
        if ($d.TemplatedParent) { continue }
        $k = if ($d.Style) { [string]$d.Style.GetHashCode() } else { 'default' }
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        $sample.Add($d)
        if ($sample.Count -ge $Max) { break }
    }
    $px = [int]([System.Windows.SystemParameters]::VirtualScreenWidth - 5)
    $py = [int]([System.Windows.SystemParameters]::VirtualScreenHeight - 5)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($btn in $sample) {
        $name = if ($btn.Name) { $btn.Name } else { [string]$btn.Content }
        $backdrop = & $backdropOf $btn
        $pt = $btn.PointToScreen((New-Object System.Windows.Point ($btn.ActualWidth / 2), ($btn.ActualHeight / 2)))
        $states = [ordered]@{}
        [void][SuiteStateUi]::SetCursorPos($px, $py); Start-Sleep -Milliseconds 400; & $pump $w 2
        $states['normal'] = & $readState $btn $backdrop
        [void]$w.Activate(); [void]$btn.Focus(); Start-Sleep -Milliseconds 500; & $pump $w 2
        $focused = $btn.IsKeyboardFocused
        $states['focus'] = & $readState $btn $backdrop
        [void][SuiteStateUi]::SetCursorPos([int]$pt.X, [int]$pt.Y); Start-Sleep -Milliseconds 600; & $pump $w 2
        $hover = $btn.IsMouseOver
        $states['hover'] = & $readState $btn $backdrop
        $pressed = $false
        try {
            [SuiteStateUi]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 500; & $pump $w 2
            $pressed = $btn.IsPressed
            $states['pressed'] = & $readState $btn $backdrop
        }
        finally {
            # Leave the button before the mouse button goes up, so that no click happens. The mouse button always goes up.
            [void][SuiteStateUi]::SetCursorPos($px, $py); Start-Sleep -Milliseconds 300
            [SuiteStateUi]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 300
        }
        $rows.Add([pscustomobject]@{ Name = $name; Hover = $hover; Pressed = $pressed; Focused = $focused; States = $states })
    }
    return $rows.ToArray()
}
#endregion
