<#
.SYNOPSIS
    Shared module for AppPackager packager scripts.

.DESCRIPTION
    Import this module at the top of every packager script to get:
      - TLS 1.2 enforcement
      - Structured logging (Write-Log, Initialize-Logging)
      - Download with retry (Invoke-DownloadWithRetry)
      - Admin check (Test-IsAdmin)
      - ConfigMgr site connection (Connect-CMSite)
      - Folder initialization (Initialize-Folder)
      - Network share access test (Test-NetworkShareAccess)
      - Content wrapper generation (Write-ContentWrappers, New-MsiWrapperContent)
      - MECM application creation (New-MECMApplicationFromManifest)
      - CM revision history cleanup (Remove-CMApplicationRevisionHistoryByCIId)

.EXAMPLE
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
    Initialize-Logging -LogPath $LogPath

    Write-Log "Starting packager..."
    Write-Log "Something went wrong" -Level ERROR
    Invoke-DownloadWithRetry -Url $url -OutFile $file
#>

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

$script:__AppPackagerLogPath = $null
$script:__AppPackagerVerbose = $false

function Initialize-Logging {
    param(
        [string]$LogPath,
        [switch]$VerboseLogging
    )

    $script:__AppPackagerLogPath = $LogPath

    # Verbose diagnostics: explicit switch wins; otherwise the
    # APP_PACKAGER_VERBOSE env var enables it (set by the GUI host or the
    # operator's shell) unless it holds an explicit "off" value.
    $envVerbose = $env:APP_PACKAGER_VERBOSE
    $script:__AppPackagerVerbose = [bool]$VerboseLogging -or
        (-not [string]::IsNullOrWhiteSpace($envVerbose) -and $envVerbose -notin @('0', 'false', 'no'))

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }

    if ($script:__AppPackagerVerbose) {
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
        DEBUG -> log file always; Write-Host (stdout) only when verbose
                 logging is enabled (Initialize-Logging -VerboseLogging or
                 APP_PACKAGER_VERBOSE env var).

        -Quiet suppresses all console output but still writes to the log file.
    #>
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

    $suppressConsole = $Quiet -or ($Level -eq 'DEBUG' -and -not $script:__AppPackagerVerbose)

    if (-not $suppressConsole) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__AppPackagerLogPath) {
        Add-Content -LiteralPath $script:__AppPackagerLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

function Write-LogErrorRecord {
    <#
    .SYNOPSIS
        Logs full diagnostics for an ErrorRecord: exception chain, error id,
        failing file/line/statement, and the script stack trace.

    .DESCRIPTION
        Designed for catch blocks. The InvocationInfo position and
        ScriptStackTrace identify the exact line of code that threw, which a
        bare $_.Exception.Message ("SCRIPT FAILED: Key cannot be null...")
        never reveals.

    .PARAMETER Level
        Log level for the diagnostic lines. Default ERROR. Pass DEBUG when the
        caller already emits its own single ERROR summary line and the detail
        should only land in the log file (unless verbose logging is enabled).
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
# Download with retry
# ---------------------------------------------------------------------------

function Invoke-DownloadWithRetry {
    <#
    .SYNOPSIS
        Downloads a file via Invoke-WebRequest with curl.exe fallback and retry logic.

    .DESCRIPTION
        Attempts to download the file using Invoke-WebRequest first.
        If Invoke-WebRequest fails, curl.exe is used as a fallback.

        Each attempt consists of:
        1. Invoke-WebRequest
        2. curl.exe if Invoke-WebRequest fails

        Retries the complete download operation according to RetryCount.

        Throws on final failure.

        Does NOT wrap scraping/variable-capture calls or URL-resolution calls.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [string[]]$ExtraCurlArgs = @(),

        [int]$RetryCount = 1,

        [int]$RetryDelaySec = 5,

        [switch]$Quiet
    )

    $maxAttempts = 1 + $RetryCount

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {

        if ($attempt -gt 1) {
            Write-Log (
                "Retrying download (attempt {0} of {1}) after {2}s delay..." -f
                $attempt,
                $maxAttempts,
                $RetryDelaySec
            ) -Level WARN -Quiet:$Quiet

            Start-Sleep -Seconds $RetryDelaySec
        }

        # ------------------------------------------------------------
        # First try: Invoke-WebRequest
        # ------------------------------------------------------------
        try {
            Write-Log "Downloading with Invoke-WebRequest: $Url" -Quiet:$Quiet

            Invoke-WebRequest `
                -Uri $Url `
                -OutFile $OutFile `
                -ErrorAction Stop

            if (Test-Path -LiteralPath $OutFile) {
                Write-Log "Download successful using Invoke-WebRequest: $OutFile" -Quiet:$Quiet
                return
            }

            throw "Invoke-WebRequest completed but output file was not created."
        }
        catch {
            $iwrError = $_.Exception.Message

            Write-Log (
                "Invoke-WebRequest failed: {0}" -f $iwrError
            ) -Level WARN -Quiet:$Quiet
        }

        # ------------------------------------------------------------
        # Fallback: curl.exe
        # ------------------------------------------------------------
        try {
            Write-Log "Falling back to curl.exe..." -Level WARN -Quiet:$Quiet

            $allArgs = @(
                '-L',
                '--fail',
                '--silent',
                '--show-error'
            ) + $ExtraCurlArgs + @(
                '-o',
                $OutFile,
                $Url
            )

            & curl.exe @allArgs 2>$null
            $exitCode = $LASTEXITCODE

            if ($exitCode -eq 0 -and (Test-Path -LiteralPath $OutFile)) {
                Write-Log "Download successful using curl.exe: $OutFile" -Quiet:$Quiet
                return
            }

            throw "curl.exe failed with exit code $exitCode."
        }
        catch {
            $curlError = $_.Exception.Message

            Write-Log (
                "curl.exe failed: {0}" -f $curlError
            ) -Level WARN -Quiet:$Quiet
        }

        # ------------------------------------------------------------
        # Both methods failed
        # ------------------------------------------------------------
        if ($attempt -lt $maxAttempts) {
            Write-Log (
                "Download attempt {0} failed using both Invoke-WebRequest and curl.exe. Will retry." -f
                $attempt
            ) -Level WARN -Quiet:$Quiet
        }
    }

    $msg = "Download failed after $maxAttempts attempt(s) using both Invoke-WebRequest and curl.exe: $Url"

    Write-Log $msg -Level ERROR -Quiet:$Quiet

    throw $msg
}

# ---------------------------------------------------------------------------
# TLS 1.2 enforcement
# ---------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------------------
# Environment & pre-flight checks
# ---------------------------------------------------------------------------

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Log "Admin check failed: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Get-AppPackagerRootPreferences {
    $prefsPath = Join-Path $PSScriptRoot '..\AppPackager.preferences.json'
    if (-not (Test-Path -LiteralPath $prefsPath)) { return $null }

    try {
        $raw = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        return $null
    }
}

function Resolve-ConfigurationManagerModulePath {
    $candidates = @()

    if ($env:SMS_ADMIN_UI_PATH) {
        $candidates += (Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1')
        $candidates += (Join-Path (Split-Path -Parent $env:SMS_ADMIN_UI_PATH) 'ConfigurationManager.psd1')
    }

    $prefs = Get-AppPackagerRootPreferences
    if ($prefs -and $prefs.DetectedTools -and $prefs.DetectedTools.ConfigMgrConsole -and $prefs.DetectedTools.ConfigMgrConsole.ModulePath) {
        $candidates += [string]$prefs.DetectedTools.ConfigMgrConsole.ModulePath
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

function Resolve-CMProviderMachineName {
    param([string]$ProviderMachineName)

    if (-not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
        return $ProviderMachineName.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_CM_PROVIDER)) {
        return $env:APP_PACKAGER_CM_PROVIDER.Trim()
    }

    $prefs = Get-AppPackagerRootPreferences
    if ($prefs) {
        foreach ($propName in @('ProviderMachineName','ServerFQDN','ProviderServer')) {
            $prop = $prefs.PSObject.Properties[$propName]
            if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                return ([string]$prop.Value).Trim()
            }
        }
        if ($prefs.MECM) {
            foreach ($propName in @('ProviderMachineName','ServerFQDN','ProviderServer')) {
                $prop = $prefs.MECM.PSObject.Properties[$propName]
                if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                    return ([string]$prop.Value).Trim()
                }
            }
        }
    }

    return $null
}

function Test-CMSiteCodeMatchesProvider {
    <#
    .SYNOPSIS
        Asks the SMS Provider which site code(s) it actually serves.

    .DESCRIPTION
        Queries root\sms:SMS_ProviderLocation on the provider machine. A
        CMSite PSDrive whose NAME does not match a real site code on its Root
        server is the classic cause of CM cmdlets failing with
        "Key cannot be null. Parameter name: key". Returns $null when the
        query itself fails (offline provider, no WMI rights).
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$ProviderMachineName
    )

    try {
        $locations = @(Get-WmiObject -ComputerName $ProviderMachineName -Namespace 'root\sms' -Class 'SMS_ProviderLocation' -ErrorAction Stop)
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

function Connect-CMSite {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = $null
    )

    try {
        if (-not (Get-Module -Name ConfigurationManager -ErrorAction SilentlyContinue)) {
            $cmModulePath = Resolve-ConfigurationManagerModulePath
            $moduleSource = if ($cmModulePath) { $cmModulePath } else { 'PSModulePath lookup' }
            Write-Log ("ConfigurationManager module  : importing from {0}" -f $moduleSource) -Level DEBUG
            if ($cmModulePath -and (Test-Path -LiteralPath $cmModulePath)) {
                Import-Module $cmModulePath -ErrorAction Stop
            }
            else {
                Import-Module ConfigurationManager -ErrorAction Stop
            }
        }
        else {
            Write-Log "ConfigurationManager module  : already loaded" -Level DEBUG
        }

        $cmDrives = @(Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue)
        $driveList = if ($cmDrives.Count -gt 0) {
            ($cmDrives | ForEach-Object { "{0} -> {1}" -f $_.Name, $_.Root }) -join '; '
        } else { '(none)' }
        Write-Log ("CMSite drives in session     : {0}" -f $driveList) -Level DEBUG

        $siteDrive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue
        $staleRoot = $null
        $connected = $false

        if ($siteDrive) {
            try {
                Set-Location "$($SiteCode):\" -ErrorAction Stop
                $connected = $true
            }
            catch {
                # Entering an existing drive can fail when its provider
                # connection is gone (e.g. provider restart). Tear it down
                # and rebuild from the provider instead of giving up.
                Write-Log ("Set-Location {0}: failed on existing drive ({1}); recreating it." -f $SiteCode, $_.Exception.Message)
                $staleRoot = [string]$siteDrive.Root
                Remove-PSDrive -Name $SiteCode -Force -ErrorAction SilentlyContinue
            }
        }

        if (-not $connected) {
            $provider = Resolve-CMProviderMachineName -ProviderMachineName $ProviderMachineName
            if ([string]::IsNullOrWhiteSpace($provider)) { $provider = $staleRoot }
            if ([string]::IsNullOrWhiteSpace($provider)) {
                throw "Configuration Manager PSDrive '$SiteCode' is not available and no provider machine name is configured. Set Provider Machine in MECM Preferences, copy the ProviderMachineName value from the AdminUI connect script, or set APP_PACKAGER_CM_PROVIDER."
            }

            Write-Log "Connecting to CM provider     : $provider"
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $provider -ErrorAction Stop | Out-Null
            Set-Location "$($SiteCode):\" -ErrorAction Stop
        }

        Write-Log "Connected to CM site: $SiteCode"

        # Verbose-only sanity check: a drive named for the wrong site code
        # connects fine but every subsequent CM cmdlet dies with
        # "Key cannot be null. Parameter name: key". Surface that here so the
        # log explains the failure instead of the cmdlet's opaque message.
        if ($script:__AppPackagerVerbose) {
            $drive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue
            if ($drive -and -not [string]::IsNullOrWhiteSpace([string]$drive.Root)) {
                $check = Test-CMSiteCodeMatchesProvider -SiteCode $SiteCode -ProviderMachineName ([string]$drive.Root)
                if ($check) {
                    if ($check.Match) {
                        Write-Log ("Provider confirms site code  : {0} on {1}" -f $SiteCode, $drive.Root) -Level DEBUG
                    }
                    else {
                        Write-Log ("Site code mismatch: drive is named '{0}' but provider {1} serves site code(s) {2}. CM cmdlets typically fail with 'Key cannot be null. Parameter name: key' in this state - correct the SiteCode in MECM Preferences / -SiteCode." -f $SiteCode, $drive.Root, ($check.SiteCodes -join ', ')) -Level WARN
                    }
                }
            }
        }

        return $true
    }
    catch {
        Write-LogErrorRecord -ErrorRecord $_ -Context ("Connect-CMSite SiteCode={0}" -f $SiteCode) -Level DEBUG
        Write-Log "Failed to connect to CM site: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Initialize-Folder {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Log "Network path does not exist or is inaccessible: $Path" -Level ERROR
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Log "Network share is not writable: $Path ($($_.Exception.Message))" -Level ERROR
            return $false
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# MECM helpers
# ---------------------------------------------------------------------------

function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)

    $installer = $null
    $db = $null
    $view = $null
    $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}

        foreach ($p in $wanted) {
            $sql  = "SELECT `Value` FROM `Property` WHERE `Property`='$p'"
            $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)

            if ($null -ne $record) {
                $val = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
                $map[$p] = $val
            }
            else {
                $map[$p] = $null
            }
        }

        return $map
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------------
# ARP (Add/Remove Programs) registry discovery
# ---------------------------------------------------------------------------

function Find-UninstallEntry {
    <#
    .SYNOPSIS
        Searches the ARP uninstall registry keys for a product by DisplayName.

    .DESCRIPTION
        Searches both native and WOW6432Node uninstall registry paths for entries
        matching the given DisplayName pattern. Returns the registry key path
        (relative, ready for New-CMDetectionClauseRegistryKeyValue), DisplayVersion,
        Publisher, and uninstall strings.

        Supports retry/polling for installers that register asynchronously.
    #>
    param(
        [Parameter(Mandatory)][string]$DisplayNamePattern,

        [string]$ExpectedVersion,

        [int]$MaxRetries = 1,

        [int]$RetryDelaySec = 0
    )

    $uninstallRoots = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $true },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $false }
    )

    $found = $null

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Registry poll attempt {0}/{1} - sleeping {2}s..." -f $attempt, $MaxRetries, $RetryDelaySec) -Level WARN
            Start-Sleep -Seconds $RetryDelaySec
        }

        $candidates = @()

        foreach ($root in $uninstallRoots) {
            $keys = Get-ChildItem -Path $root.Path -ErrorAction SilentlyContinue
            if (-not $keys) { continue }

            foreach ($k in $keys) {
                $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $dn = $props.DisplayName
                if ([string]::IsNullOrWhiteSpace($dn)) { continue }

                if ($dn -like $DisplayNamePattern) {
                    $regRelative = ($root.Path -replace '^HKLM:\\', '') + '\' + $k.PSChildName

                    $candidates += [pscustomobject]@{
                        RegistryKeyRelative  = $regRelative
                        DisplayName          = $dn
                        DisplayVersion       = $props.DisplayVersion
                        Publisher            = $props.Publisher
                        UninstallString      = $props.UninstallString
                        QuietUninstallString = $props.QuietUninstallString
                        Is64Bit              = $root.Is64Bit
                    }
                }
            }
        }

        if ($candidates.Count -gt 0) {
            if ($ExpectedVersion) {
                $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
                if ($match) { $found = $match; break }
            }

            $found = $candidates | Select-Object -First 1
            break
        }
    }

    return $found
}

# ---------------------------------------------------------------------------
# Stage manifest
# ---------------------------------------------------------------------------

function Get-StageFileHashes {
    <#
    .SYNOPSIS
        Computes SHA256 + size metadata for files under a stage folder.

    .DESCRIPTION
        Returns an ordered list of RelativePath/Sha256/Size records. Relative
        paths use backslashes so manifest data is stable across callers. The
        stage manifest itself is excluded by default because it is the recorder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [string[]]$Exclude = @('stage-manifest.json')
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        throw "Stage root not found: $Root"
    }

    $rootItem = Get-Item -LiteralPath $Root -ErrorAction Stop
    if (-not $rootItem.PSIsContainer) {
        throw "Stage root is not a directory: $Root"
    }

    $rootFull = $rootItem.FullName.TrimEnd('\', '/')
    $excludeSet = @{}
    foreach ($item in @($Exclude)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $normalized = ([string]$item).TrimStart('\', '/') -replace '/', '\'
        $excludeSet[$normalized.ToLowerInvariant()] = $true
    }

    $records = New-Object System.Collections.Generic.List[object]
    $files = @(Get-ChildItem -LiteralPath $rootFull -File -Recurse -Force -ErrorAction Stop | Sort-Object FullName)
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\', '/') -replace '/', '\'
        if ($excludeSet.ContainsKey($relative.ToLowerInvariant())) { continue }

        $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop
        $records.Add([pscustomobject]@{
            RelativePath = $relative
            Sha256       = $hash.Hash.ToUpperInvariant()
            Size         = [Int64]$file.Length
        })
    }

    return @($records.ToArray())
}

function Compare-StageFileHashes {
    <#
    .SYNOPSIS
        Compares a folder tree to expected stage-manifest file hashes.

    .DESCRIPTION
        Returns a result object instead of throwing so callers can decide
        whether to hard-fail or warn. Missing expected hash lists are treated as
        a soft-landing pass for pre-1.0.7 manifests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [AllowNull()][object[]]$Expected,
        [string[]]$Exclude = @('stage-manifest.json'),
        [switch]$AllowExtra
    )

    if ($null -eq $Expected) {
        return [pscustomobject]@{
            Pass          = $true
            Skipped       = $true
            Reason        = 'Stage manifest does not contain FileHashes.'
            Missing       = @()
            Mismatches    = @()
            Extra         = @()
            ExpectedCount = 0
            ActualCount   = 0
            Root          = $Root
        }
    }

    if (-not (Test-Path -LiteralPath $Root)) {
        return [pscustomobject]@{
            Pass          = $false
            Skipped       = $false
            Reason        = "Stage root not found: $Root"
            Missing       = @($Expected)
            Mismatches    = @()
            Extra         = @()
            ExpectedCount = @($Expected).Count
            ActualCount   = 0
            Root          = $Root
        }
    }

    $expectedMap = @{}
    foreach ($entry in @($Expected)) {
        if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.RelativePath)) { continue }
        $relative = ([string]$entry.RelativePath).TrimStart('\', '/') -replace '/', '\'
        $expectedMap[$relative.ToLowerInvariant()] = [pscustomobject]@{
            RelativePath = $relative
            Sha256       = ([string]$entry.Sha256).ToUpperInvariant()
            Size         = [Int64]$entry.Size
        }
    }

    $actual = @(Get-StageFileHashes -Root $Root -Exclude $Exclude)
    $actualMap = @{}
    foreach ($entry in $actual) {
        $actualMap[[string]$entry.RelativePath.ToLowerInvariant()] = $entry
    }

    $missing = New-Object System.Collections.Generic.List[object]
    $mismatches = New-Object System.Collections.Generic.List[object]
    foreach ($key in $expectedMap.Keys) {
        $expectedEntry = $expectedMap[$key]
        if (-not $actualMap.ContainsKey($key)) {
            $missing.Add($expectedEntry)
            continue
        }

        $actualEntry = $actualMap[$key]
        if ([Int64]$actualEntry.Size -ne [Int64]$expectedEntry.Size -or
            ([string]$actualEntry.Sha256).ToUpperInvariant() -ne ([string]$expectedEntry.Sha256).ToUpperInvariant()) {
            $mismatches.Add([pscustomobject]@{
                RelativePath   = $expectedEntry.RelativePath
                ExpectedSha256 = $expectedEntry.Sha256
                ActualSha256   = $actualEntry.Sha256
                ExpectedSize   = [Int64]$expectedEntry.Size
                ActualSize     = [Int64]$actualEntry.Size
            })
        }
    }

    $extra = New-Object System.Collections.Generic.List[object]
    if (-not $AllowExtra) {
        foreach ($key in $actualMap.Keys) {
            if (-not $expectedMap.ContainsKey($key)) {
                $extra.Add($actualMap[$key])
            }
        }
    }

    $missingArray = @($missing.ToArray())
    $mismatchArray = @($mismatches.ToArray())
    $extraArray = @($extra.ToArray())

    return [pscustomobject]@{
        Pass          = ($missingArray.Count -eq 0 -and $mismatchArray.Count -eq 0 -and $extraArray.Count -eq 0)
        Skipped       = $false
        Reason        = ''
        Missing       = $missingArray
        Mismatches    = $mismatchArray
        Extra         = $extraArray
        ExpectedCount = $expectedMap.Count
        ActualCount   = $actual.Count
        Root          = $Root
    }
}

function Format-StageFileHashComparison {
    param([Parameter(Mandatory)]$Comparison)

    if ($Comparison.Pass) { return 'Stage/package file integrity verified.' }
    if ($Comparison.Skipped) { return [string]$Comparison.Reason }

    $parts = New-Object System.Collections.Generic.List[string]
    if ($Comparison.Missing.Count -gt 0) {
        $sample = @($Comparison.Missing | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("missing {0}: {1}" -f $Comparison.Missing.Count, $sample))
    }
    if ($Comparison.Mismatches.Count -gt 0) {
        $sample = @($Comparison.Mismatches | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("mismatched {0}: {1}" -f $Comparison.Mismatches.Count, $sample))
    }
    if ($Comparison.Extra.Count -gt 0) {
        $sample = @($Comparison.Extra | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("extra {0}: {1}" -f $Comparison.Extra.Count, $sample))
    }
    if ($parts.Count -eq 0 -and $Comparison.Reason) { $parts.Add([string]$Comparison.Reason) }
    return ($parts.ToArray() -join '; ')
}

function Write-StageManifest {
    <#
    .SYNOPSIS
        Writes a stage-manifest.json file.

    .DESCRIPTION
        Serializes ManifestData to JSON with schema metadata.

        Schema v2 adds optional fields for PSADT/deployment tool integration:
          InstallerType     "MSI" or "EXE"
          InstallArgs       Silent install arguments
          UninstallArgs     Silent uninstall arguments
          UninstallCommand  Full uninstall command (for EXE products)
          ProductCode       MSI ProductCode GUID (for MSI products)
          RunningProcess    Array of process names to close before install

        Schema v3 adds FileHashes, an ordered list of every staged payload and
        wrapper file with RelativePath, SHA256, and Size. stage-manifest.json is
        excluded from its own hash list.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$ManifestData
    )

    $stageRoot = Split-Path -Path $Path -Parent
    $manifestName = Split-Path -Path $Path -Leaf
    $fileHashes = Get-StageFileHashes -Root $stageRoot -Exclude @($manifestName)

    $ManifestData['SchemaVersion'] = 3
    $ManifestData['StagedAt'] = (Get-Date -Format 'o')
    $ManifestData['FileHashes'] = @($fileHashes)

    $json = $ManifestData | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8 -ErrorAction Stop
    Write-Log "Wrote stage manifest         : $Path"
    Write-Log ("Recorded file hashes         : {0} file(s)" -f @($fileHashes).Count)

    $verification = Compare-StageFileHashes -Root $stageRoot -Expected $fileHashes -Exclude @($manifestName)
    if (-not $verification.Pass) {
        throw ("Stage integrity verification failed: {0}" -f (Format-StageFileHashComparison -Comparison $verification))
    }
    Write-Log ("Stage integrity verified     : {0} file(s)" -f $verification.ExpectedCount)
}

function Read-StageManifest {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Stage manifest not found: $Path"
    }

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
    $manifest = $json | ConvertFrom-Json

    if (-not $manifest.SchemaVersion) {
        throw "Invalid stage manifest (missing SchemaVersion): $Path"
    }

    $hasFileHashes = ($manifest.PSObject.Properties.Name -contains 'FileHashes') -and $null -ne $manifest.FileHashes
    if (-not $hasFileHashes) {
        if ([int]$manifest.SchemaVersion -ge 3) {
            throw "Invalid stage manifest (SchemaVersion $($manifest.SchemaVersion) missing FileHashes): $Path"
        }
        Write-Log "Stage manifest has no file hashes; byte-level integrity verification skipped for this pre-1.0.7 manifest." -Level WARN
    }

    Write-Log "Read stage manifest          : $Path"
    return $manifest
}

function Update-StageManifest {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Destination,
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Stage manifest not found: $Path"
    }

    # Read existing manifest
    $manifest = Get-Content -LiteralPath $Path -Raw |
        ConvertFrom-Json

    # Update RelativePath for every FileHashes entry
    foreach ($fileHash in $manifest.FileHashes) {
        $fileHash.RelativePath = Join-Path $RelativePath $fileHash.RelativePath
    }

    # Write manifest back to disk
    $manifest |
        ConvertTo-Json -Depth 10 |
        Set-Content -Path $Destination -Encoding UTF8
}

# ---------------------------------------------------------------------------
# MECM helpers (continued)
# ---------------------------------------------------------------------------

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][UInt32]$CI_ID,
        [UInt32]$KeepLatest = 1
    )

    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }

    $revs = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

# ---------------------------------------------------------------------------
# Network path helpers
# ---------------------------------------------------------------------------

function Get-NetworkAppRoot {
    <#
    .SYNOPSIS
        Creates and returns the network content root for an application.

    .DESCRIPTION
        Builds the path based on the Application Share Pattern parameter,
        creating each level if it does not exist. Returns the final path.
    #>
        param(
        [Parameter(Mandatory)]
        [string]$FileServerPath,

        [Parameter(Mandatory)]
        [string]$PathPattern,

        [Parameter(Mandatory)]
        [hashtable]$Variables
    )

    $relativePath = $PathPattern

    foreach ($key in $Variables.Keys) {
        $value = $Variables[$key]

        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $relativePath = $relativePath -replace "_?\{$key\}?", ""
        }
        else {
            $relativePath = $relativePath.Replace("{$key}", ([string]$value -replace " ", "-"))
        }
    }

    Initialize-Folder -Path (Join-Path $FileServerPath $relativePath)

    return Join-Path $FileServerPath $relativePath
}

# ---------------------------------------------------------------------------
# Content wrapper generation
# ---------------------------------------------------------------------------

function Write-ContentWrappers {
    <#
    .SYNOPSIS
        Creates install/uninstall .bat and .ps1 wrapper files in a content folder.

    .DESCRIPTION
        Writes four files to OutputPath: install.bat, install.ps1, uninstall.bat,
        uninstall.ps1. The .bat files are thin shims that call the corresponding
        .ps1. The .ps1 content is passed as strings by the caller.

        Overwrites existing wrapper files so regenerated staged content remains
        deterministic. All files are written with -Encoding ASCII to avoid BOM
        issues.

    .PARAMETER InstallBatExitCode
        Exit code expression for install.bat. Default: '%ERRORLEVEL%'.
        Use '3010' for products that always require reboot (e.g. VMware Tools).

    .PARAMETER UninstallBatExitCode
        Exit code expression for uninstall.bat. Default: '%ERRORLEVEL%'.
    #>
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$InstallPs1Content,
        [Parameter(Mandatory)][string]$UninstallPs1Content,
        [string]$InstallBatExitCode   = '%ERRORLEVEL%',
        [string]$UninstallBatExitCode = '%ERRORLEVEL%'
    )

    $installBatPath   = Join-Path $OutputPath "install.bat"
    $installPs1Path   = Join-Path $OutputPath "install.ps1"
    $uninstallBatPath = Join-Path $OutputPath "uninstall.bat"
    $uninstallPs1Path = Join-Path $OutputPath "uninstall.ps1"

    # .bat wrapper template: @echo off, call PowerShell, propagate exit code
    # When exit code override is set (e.g. 3010), only apply on success --
    # real failures must propagate so ConfigMgr can detect them.
    $installBat = if ($InstallBatExitCode -eq '%ERRORLEVEL%') {
        (@(
            '@echo off',
            'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"',
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    } else {
        (@(
            '@echo off',
            'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"',
            ('if %ERRORLEVEL% EQU 0 exit /b {0}' -f $InstallBatExitCode),
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    }

    $uninstallBat = if ($UninstallBatExitCode -eq '%ERRORLEVEL%') {
        (@(
            '@echo off',
            'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"',
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    } else {
        (@(
            '@echo off',
            'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"',
            ('if %ERRORLEVEL% EQU 0 exit /b {0}' -f $UninstallBatExitCode),
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    }

    $files = @(
        @{ Path = $installBatPath;   Content = $installBat;          Label = 'install.bat' },
        @{ Path = $installPs1Path;   Content = $InstallPs1Content;   Label = 'install.ps1' },
        @{ Path = $uninstallBatPath; Content = $uninstallBat;        Label = 'uninstall.bat' },
        @{ Path = $uninstallPs1Path; Content = $UninstallPs1Content; Label = 'uninstall.ps1' }
    )

    foreach ($f in $files) {
        Set-Content -LiteralPath $f.Path -Value $f.Content -Encoding ASCII -Force -ErrorAction Stop
        Write-Log "Wrote wrapper                : $($f.Label)"
    }
}

function New-MsiWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an MSI product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList (avoiding quoting issues). Returns a hashtable
        with Install and Uninstall keys.

        -ExtraInstallArgs: optional MSI properties (e.g. "APITOKEN=xxx",
        'ASSIGNMENTOPTIONS="--grant-easy-access"') appended to the install
        command line. Each entry becomes one Start-Process argument.

        -PostInstallKillProcesses: optional array of process names (no .exe)
        that the install.ps1 should Stop-Process after msiexec returns. Used
        to kill the installer-spawned GUI that runs under the SYSTEM context
        (TeamViewer, TeamViewer Host, anything else that helpfully launches
        a post-install splash).
    #>
    param(
        [Parameter(Mandatory)][string]$MsiFileName,
        [string[]]$ExtraInstallArgs = @(),
        [string[]]$PostInstallKillProcesses = @()
    )

    $installLines = @(
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$args = @(''/i'', "`"$msiPath`"", ''/qn'', ''/norestart'')'
    )
    foreach ($extra in $ExtraInstallArgs) {
        $escaped = $extra -replace "'", "''"
        $installLines += ('$args += ''{0}''' -f $escaped)
    }
    $installLines += '$proc = Start-Process msiexec.exe -ArgumentList $args -Wait -PassThru -NoNewWindow'
    $installLines += '$exit = $proc.ExitCode'
    if ($PostInstallKillProcesses.Count -gt 0) {
        $procList = ($PostInstallKillProcesses | ForEach-Object { "'$_'" }) -join ', '
        $installLines += ('foreach ($pn in @({0})) {{ Get-Process -Name $pn -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }}' -f $procList)
    }
    $installLines += 'exit $exit'
    $install = $installLines -join "`r`n"

    $uninstall = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

function New-MsixWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an MSIX/APPX
        package, using the Script deployment-type pattern (install.bat +
        install.ps1) so MECM treats it the same way as MSI / EXE packagers.

    .DESCRIPTION
        MECM also supports a native MSIX deployment type via
        Add-CMWindowsAppxDeploymentType, but the house rule is Script
        deployment for everything we can shoehorn that way (single code
        path, uniform logging, consistent detection authoring). These
        wrappers use Add-AppxProvisionedPackage for deployment-wide
        (per-device) installs. Per-user MSIX can layer on top of this
        pattern but isn't the default.

        -MsixFileName  : name of the .msix / .appx / .msixbundle in the
                         package content folder.
        -Provisioned   : when true (default), installs via
                         Add-AppxProvisionedPackage -Online so all users
                         (including new accounts) get the app. Uninstall
                         uses Remove-AppxProvisionedPackage.
                         When false, installs per-user via Add-AppxPackage.
        -SignatureSha1 : optional SHA1 thumbprint of the expected signing
                         certificate. When set, install.ps1 verifies the
                         MSIX is signed by that cert before invoking the
                         AppX cmdlet (catches tampered / wrong-package
                         downloads before deployment).
    #>
    param(
        [Parameter(Mandatory)][string]$MsixFileName,
        [bool]$Provisioned = $true,
        [string]$SignatureSha1 = ''
    )

    $sigCheckLines = @()
    if (-not [string]::IsNullOrWhiteSpace($SignatureSha1)) {
        $sigCheckLines = @(
            ('$expected = ''{0}''' -f ($SignatureSha1 -replace "'", "''")),
            '$sig = Get-AuthenticodeSignature -LiteralPath $msixPath',
            'if ($sig.Status -ne ''Valid'') { Write-Error "MSIX signature not valid: $($sig.Status)"; exit 2 }',
            'if ($sig.SignerCertificate.Thumbprint -ne $expected) { Write-Error "MSIX signed by unexpected cert (got $($sig.SignerCertificate.Thumbprint), expected $expected)"; exit 3 }'
        )
    }

    if ($Provisioned) {
        $installBody = @(
            'Add-AppxProvisionedPackage -Online -PackagePath $msixPath -SkipLicense | Out-Null'
        )
        $uninstallBody = @(
            ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
            '# Provisioned-package removal needs the PackageName, which is stored',
            '# inside the MSIX AppxManifest. Import-Metadata pattern: read the',
            '# manifest during install and stash the name; for the uninstall path',
            '# re-read the manifest on the client.',
            '$zip = [System.IO.Compression.ZipFile]::OpenRead($msixPath)',
            'try {',
            '    $entry = $zip.GetEntry(''AppxManifest.xml'')',
            '    if (-not $entry) { Write-Error ''AppxManifest.xml not found in MSIX''; exit 4 }',
            '    $reader = New-Object System.IO.StreamReader($entry.Open())',
            '    try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }',
            '} finally { $zip.Dispose() }',
            '$identity = $manifest.Package.Identity',
            '$pkgFullName = ''{0}_{1}_{2}__{3}'' -f $identity.Name, $identity.Version, $identity.ProcessorArchitecture, $identity.PublisherId',
            '$prov = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $pkgFullName -or $_.DisplayName -eq $identity.Name }',
            'if ($prov) { Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName | Out-Null }',
            'Get-AppxPackage -AllUsers -Name $identity.Name | ForEach-Object { Remove-AppxPackage -Package $_.PackageFullName -AllUsers }',
            'exit 0'
        )
    } else {
        $installBody = @(
            'Add-AppxPackage -Path $msixPath'
        )
        $uninstallBody = @(
            ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
            '$zip = [System.IO.Compression.ZipFile]::OpenRead($msixPath)',
            'try {',
            '    $entry = $zip.GetEntry(''AppxManifest.xml'')',
            '    if (-not $entry) { Write-Error ''AppxManifest.xml not found in MSIX''; exit 4 }',
            '    $reader = New-Object System.IO.StreamReader($entry.Open())',
            '    try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }',
            '} finally { $zip.Dispose() }',
            '$name = $manifest.Package.Identity.Name',
            'Get-AppxPackage -Name $name | Remove-AppxPackage',
            'exit 0'
        )
    }

    $installLines = @(
        'Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem',
        ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
        'if (-not (Test-Path -LiteralPath $msixPath)) { Write-Error "MSIX not found: $msixPath"; exit 1 }'
    ) + $sigCheckLines + $installBody + @('exit 0')

    return @{
        Install   = ($installLines -join "`r`n")
        Uninstall = ($uninstallBody -join "`r`n")
    }
}

function New-ExeWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an EXE product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList for the installer EXE. Returns a hashtable
        with Install and Uninstall keys.

        For products where uninstall uses a different command (e.g. registry
        lookup, msiexec), the caller should build uninstall content directly
        and pass it to Write-ContentWrappers.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$InstallArgs,
        [Parameter(Mandatory)][string]$UninstallCommand,
        [string]$UninstallArgs = ''
    )

    $install = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $InstallArgs),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    if ($UninstallArgs -ne '') {
        $uninstall = (
            ('$proc = Start-Process -FilePath ''{0}'' -ArgumentList @({1}) -Wait -PassThru -NoNewWindow' -f $UninstallCommand, $UninstallArgs),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }
    else {
        $uninstall = (
            ('$proc = Start-Process -FilePath ''{0}'' -Wait -PassThru -NoNewWindow' -f $UninstallCommand),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

# ---------------------------------------------------------------------------
# MECM application creation from manifest
# ---------------------------------------------------------------------------

function New-SingleDetectionClause {
    <#
    .SYNOPSIS
        Builds a single CM detection clause object from a manifest detection block.
    .DESCRIPTION
        Internal helper for New-MECMApplicationFromManifest. Supports
        RegistryKeyValue, RegistryKey, and File detection types.
        Must be called while the current location is a filesystem drive
        (not the CM PSDrive).
    #>
    param([Parameter(Mandatory)][pscustomobject]$Det)

    $type = if ($Det.Type) { $Det.Type } else { 'RegistryKeyValue' }

    switch ($type) {
        'RegistryKeyValue' {
            $operator = if ($Det.Operator) { $Det.Operator } else { 'IsEquals' }
            $expected = if ($Det.ExpectedValue) { $Det.ExpectedValue } else { $Det.DisplayVersion }
            $valName  = if ($Det.ValueName) { $Det.ValueName } else { 'DisplayVersion' }
            $propType = if ($Det.PropertyType) { $Det.PropertyType } else { 'String' }

            $p = @{
                Hive               = 'LocalMachine'
                KeyName            = $Det.RegistryKeyRelative
                ValueName          = $valName
                PropertyType       = $propType
                Value              = $true
                ExpressionOperator = $operator
                ExpectedValue      = $expected
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKeyValue @p)
        }
        'RegistryKey' {
            $p = @{
                Hive      = 'LocalMachine'
                KeyName   = $Det.RegistryKeyRelative
                Existence = $true
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKey @p)
        }
        'File' {
            if ($Det.PropertyType -eq 'Existence') {
                $p = @{
                    Path      = $Det.FilePath
                    FileName  = $Det.FileName
                    Existence = $true
                }
            }
            else {
                $op = if ($Det.Operator) { $Det.Operator } else { 'GreaterEquals' }
                $p = @{
                    Path               = $Det.FilePath
                    FileName           = $Det.FileName
                    PropertyType       = $Det.PropertyType
                    Value              = $true
                    ExpressionOperator = $op
                    ExpectedValue      = $Det.ExpectedValue
                }
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseFile @p)
        }
        default { throw "Unsupported detection clause type: $type" }
    }
}

function Test-MECMApplicationHasDeploymentType {
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$DeploymentTypeName
    )

    if (-not (Get-Command -Name Get-CMDeploymentType -ErrorAction SilentlyContinue)) {
        throw "Get-CMDeploymentType is not available; cannot validate existing application '$ApplicationName'."
    }

    try {
        $deploymentTypes = @(Get-CMDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName -ErrorAction SilentlyContinue)
    }
    catch {
        $deploymentTypes = @()
    }

    if ($deploymentTypes.Count -eq 0) {
        try {
            $deploymentTypes = @(Get-CMDeploymentType -ApplicationName $ApplicationName -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.LocalizedDisplayName -eq $DeploymentTypeName -or
                    $_.DeploymentTypeName -eq $DeploymentTypeName -or
                    $_.Name -eq $DeploymentTypeName
                })
        }
        catch {
            $deploymentTypes = @()
        }
    }

    return ($deploymentTypes.Count -gt 0)
}


function New-MECMApplicationFromManifest {
    <#
    .SYNOPSIS
        Creates an MECM application with Script deployment type from a stage manifest.

    .DESCRIPTION
        Reads a stage manifest object and creates a CM Application with a single
        Script deployment type. Supports all detection methods:

          RegistryKeyValue  Single registry value comparison (IsEquals or GreaterEquals)
          RegistryKey       Registry key existence check
          File              File existence or version comparison
          Script            PowerShell script-based detection
          Compound          Multiple clauses joined by AND or OR

        Handles CM site connection, duplicate app check, New-CMApplication with
        -AutoInstall $true, detection clause creation, Add-CMScriptDeploymentType
        with gold standard parameters, optional PostExecutionBehavior, and
        revision history cleanup.

        Backward compatible: manifests without a Detection.Type field default
        to RegistryKeyValue; DisplayVersion is accepted as an alias for
        ExpectedValue.

    .OUTPUTS
        [UInt32] CI_ID of the created or already-complete application.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$AppNamePattern,
        [Parameter(Mandatory)][string]$SiteCode,
        [AllowEmptyString()][string]$Comment = '',
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [string]$PSAppDeployToolkitPath = $null,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30
    )

    $orig = Get-Location

    if($PSAppDeployToolkitPath -and [string]::IsNullOrWhiteSpace($PSAppDeployToolkitPath) -eq $false) {
        $contentVerification = Compare-StageFileHashes -Root $NetworkContentPath -Expected $Manifest.FileHashes -AllowExtra
    } else {
        $contentVerification = Compare-StageFileHashes -Root $NetworkContentPath -Expected $Manifest.FileHashes
    }
    
    if ($contentVerification.Skipped) {
        Write-Log ("Package integrity verification : skipped ({0})" -f $contentVerification.Reason) -Level WARN
    }
    elseif (-not $contentVerification.Pass) {
        throw ("Package integrity verification failed: {0}" -f (Format-StageFileHashComparison -Comparison $contentVerification))
    }
    else {
        Write-Log ("Package integrity verified     : {0} file(s)" -f $contentVerification.ExpectedCount)
    }

    # Tracks the operation in flight so the catch block can name the step
    # that failed. CM cmdlet errors (e.g. "Key cannot be null. Parameter
    # name: key") rarely identify their own call site.
    $step = 'initialization'

    try {
        $step = "Connect-CMSite (SiteCode=$SiteCode)"
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "CM site connection failed."
        }

        $appName = $Manifest.AppName
        $cmAppName = $AppNamePattern

        $variables = @{ Publisher = $Manifest.Publisher; AppName = $Manifest.AppName; SoftwareVersion = $Manifest.SoftwareVersion; Language = $Manifest.Language; Architecture = $Manifest.Architecture }

        foreach ($key in $Variables.Keys) {
            $value = $Variables[$key]

            if ([string]::IsNullOrWhiteSpace([string]$value)) {
                $cmAppName = $cmAppName -replace "_?\{$key\}_?", ""
            } else {
                $cmAppName = $cmAppName.Replace("{$key}", ([string]$value -replace " ", "-"))
            }
        }

        if ([string]::IsNullOrWhiteSpace([string]$cmAppName)) {
            throw "Stage manifest AppName is null or empty; cannot create an MECM application. Re-run the Stage phase and verify the manifest."
        }

        Write-Log ("Manifest fields              : AppName='{0}' Publisher='{1}' SoftwareVersion='{2}' DetectionType='{3}'" -f $appName, $Manifest.Publisher, $Manifest.SoftwareVersion, $Manifest.Detection.Type) -Level DEBUG

        $step = "Get-CMApplication duplicate check ('$cmAppName')"
        $existing = Get-CMApplication -Name $cmAppName -ErrorAction SilentlyContinue
        if ($existing) {
            $existingApps = @($existing)
            if ($existingApps.Count -gt 1) {
                throw "Multiple existing MECM applications matched '$cmAppName'; refusing to package until the duplicate names are resolved."
            }

            $dtName = $cmAppName
            if (Test-MECMApplicationHasDeploymentType -ApplicationName $cmAppName -DeploymentTypeName $dtName) {
                Write-Log "Application already exists    : $cmAppName" -Level WARN
                Write-Log "Deployment type validated     : $dtName"
                return [UInt32]$existingApps[0].CI_ID
            }

            throw "Existing MECM application '$cmAppName' is missing deployment type '$dtName'. This looks like a partial prior package run; fix or remove the partial app before packaging again."
        }

        Write-Log "Creating CM Application      : $cmAppName"
        $step = "New-CMApplication ('$cmAppName')"
        $cmAppParams = @{
            Name             = $cmAppName
            Publisher        = $Manifest.Publisher
            SoftwareVersion  = $Manifest.SoftwareVersion
            Description      = $Comment
            AutoInstall      = $true
            ErrorAction      = 'Stop'
        }
        # Set Software Center display name if provided (omits channel/arch details)
        if ($Manifest.DisplayName) {
            $cmAppParams['LocalizedApplicationName'] = $Manifest.DisplayName
            Write-Log "Software Center name         : $($Manifest.DisplayName)"
        }
        $cmApp = New-CMApplication @cmAppParams

        Write-Log "Application CI_ID            : $($cmApp.CI_ID)"

        # Determine detection type (backward compat: missing Type = RegistryKeyValue)
        $detType = if ($Manifest.Detection.Type) { $Manifest.Detection.Type } else { 'RegistryKeyValue' }

        # Common deployment type parameters (splatted)
        $dtName = $cmAppName
        $dtParams = @{
            ApplicationName           = $cmAppName
            DeploymentTypeName        = $cmAppName + " - Install"
            ContentLocation           = $NetworkContentPath
            InstallationBehaviorType  = 'InstallForSystem'
            LogonRequirementType      = 'WhetherOrNotUserLoggedOn'
            EstimatedRuntimeMins      = $EstimatedRuntimeMins
            MaximumRuntimeMins        = $MaximumRuntimeMins
            ContentFallback           = $true
            SlowNetworkDeploymentMode = 'Download'
            UserInteractionMode       = 'Hidden'
            ErrorAction               = 'Stop'
        }

        if($PSAppDeployToolkitPath -and [string]::IsNullOrWhiteSpace($PSAppDeployToolkitPath) -eq $false) {
            $dtParams['InstallCommand']     = 'Deploy-Application.EXE INSTALL'
            $dtParams['UninstallCommand']   = 'Deploy-Application.EXE UNINSTALL'
        } else {
            $dtParams['InstallCommand']     = 'install.bat'
            $dtParams['UninstallCommand']   = 'uninstall.bat'
        }

        # Manifest field name matches the cmdlet's parameter TYPE
        # (PostExecutionBehavior); the actual parameter name is
        # -RebootBehavior. See Add-CMScriptDeploymentType docs.
        # Default to BasedOnExitCode when the manifest doesn't specify so
        # MSI 3010 "reboot required" exits actually propagate to the CCM
        # client. The cmdlet's own default is NoAction, which silently drops
        # 3010s and breaks install chains that need a reboot between apps.
        if ($Manifest.PostExecutionBehavior) {
            $dtParams['RebootBehavior'] = $Manifest.PostExecutionBehavior
        }
        else {
            $dtParams['RebootBehavior'] = 'BasedOnExitCode'
        }

        if ($Manifest.InstallationBehaviorType) {
            $dtParams['InstallationBehaviorType'] = $Manifest.InstallationBehaviorType
        }
        if ($Manifest.LogonRequirementType) {
            $dtParams['LogonRequirementType'] = $Manifest.LogonRequirementType
        }
        if ($Manifest.RequireUserInteraction -eq $true) {
            $dtParams['RequireUserInteraction'] = $true
        }

        if ($detType -eq 'Script') {
            # Script-based detection: pass script text, no clause objects needed
            $lang = if ($Manifest.Detection.ScriptLanguage) { $Manifest.Detection.ScriptLanguage } else { 'PowerShell' }
            $dtParams['ScriptLanguage'] = $lang
            $dtParams['ScriptText']     = $Manifest.Detection.ScriptText
        }
        else {
            # Clause-based detection: leave CM PSDrive to create clause objects
            # (CM PSDrive context can interfere with parameter binding)
            $step = "New detection clause(s) (type=$detType)"
            Set-Location C: -ErrorAction Stop

            if ($detType -eq 'Compound') {
                $clauses = @()
                foreach ($c in $Manifest.Detection.Clauses) {
                    $clauses += New-SingleDetectionClause -Det $c
                }
                $dtParams['AddDetectionClause'] = $clauses

                # OR connector: specify OR for each clause beyond the first
                # AND is the default and needs no explicit connector
                if ($Manifest.Detection.Connector -eq 'Or' -and $clauses.Count -ge 2) {
                    $connectors = @()
                    for ($i = 1; $i -lt $clauses.Count; $i++) {
                        $connectors += @{
                            LogicalName = $clauses[$i].Setting.LogicalName
                            Connector   = 'OR'
                        }
                    }
                    $dtParams['DetectionClauseConnector'] = $connectors
                }
            }
            else {
                # Single clause: RegistryKeyValue, RegistryKey, or File
                $clause = New-SingleDetectionClause -Det $Manifest.Detection
                $dtParams['AddDetectionClause'] = @($clause)
            }

            # Reconnect to CM site for Add-CMScriptDeploymentType
            $step = "Connect-CMSite reconnect (SiteCode=$SiteCode)"
            if (-not (Connect-CMSite -SiteCode $SiteCode)) {
                throw "CM site reconnection failed."
            }
        }

        Write-Log ("Deployment type parameters   : {0}" -f (($dtParams.Keys | Sort-Object | ForEach-Object { "{0}='{1}'" -f $_, $dtParams[$_] }) -join ' ')) -Level DEBUG

        Write-Log "Adding Script Deployment Type : $dtName"
        $step = "Add-CMScriptDeploymentType ('$dtName')"
        Add-CMScriptDeploymentType @dtParams | Out-Null

        $step = "Remove-CMApplicationRevisionHistory (CI_ID=$($cmApp.CI_ID))"
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        # Optional auto-distribute to a DP group. Settings live in
        # AppPackager.preferences.json alongside the GUI. Packagers invoked
        # from the CLI with no prefs file silently skip this step.
        try {
            $prefsPath = Join-Path $PSScriptRoot '..\AppPackager.preferences.json'
            if (Test-Path -LiteralPath $prefsPath) {
                $prefsRaw = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($prefsRaw)) {
                    $prefsObj = $prefsRaw | ConvertFrom-Json -ErrorAction Stop
                    $cd = $prefsObj.ContentDistribution
                    if ($cd -and [bool]$cd.AutoDistribute -and -not [string]::IsNullOrWhiteSpace([string]$cd.DPGroupName)) {
                        $dpGroup = [string]$cd.DPGroupName
                        Write-Log "Distributing content         : DP group '$dpGroup'"
                        try {
                            Start-CMContentDistribution -ApplicationName $appName -DistributionPointGroupName $dpGroup -ErrorAction Stop
                            Write-Log "Content distribution         : initiated"
                        } catch {
                            # "already been targeted" is the canonical pattern for a DP group that already holds this app
                            if ($_.Exception.Message -match 'already been targeted|already distributed') {
                                Write-Log "Content distribution         : already targeted (treated as success)"
                            } else {
                                Write-Log "Content distribution failed  : $($_.Exception.Message)" -Level WARN
                            }
                        }
                    }
                }
            }
        } catch {
            Write-Log "Could not read prefs for auto-distribute: $($_.Exception.Message)" -Level WARN
        }

        Write-Log ""
        Write-Log "Created MECM application     : $appName"

        return [UInt32]$cmApp.CI_ID
    }
    catch {
        Write-LogErrorRecord -ErrorRecord $_ -Context ("New-MECMApplicationFromManifest failed during step: {0}" -f $step)
        throw
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Packager preferences
# ---------------------------------------------------------------------------

function Get-PackagerPreferences {
    <#
    .SYNOPSIS
        Reads packager-preferences.json from the Packagers folder.
    #>
    $prefsPath = Join-Path $PSScriptRoot "packager-preferences.json"
    if (-not (Test-Path -LiteralPath $prefsPath)) {
        Write-Log "Preferences file not found: $prefsPath" -Level WARN
        return $null
    }
    $json = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
    return ($json | ConvertFrom-Json)
}

# ---------------------------------------------------------------------------
# ODT config XML generation
# ---------------------------------------------------------------------------

function ConvertTo-XmlAttributeValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function New-OdtConfigXml {
    <#
    .SYNOPSIS
        Generates a full ODT configuration XML string for download or install.

    .DESCRIPTION
        Builds the XML matching the production ODT template with all properties,
        excluded apps, AppSettings, logging, etc. Used by all M365 packager
        scripts for both download.xml and install.xml.

    .PARAMETER OfficeClientEdition
        Architecture: "32" or "64".

    .PARAMETER Version
        Full M365 version string (e.g. "16.0.19127.20532").

    .PARAMETER ProductIds
        Array of product IDs (e.g. @('O365ProPlusRetail') or
        @('O365ProPlusRetail', 'VisioProRetail')).

    .PARAMETER SourcePath
        SourcePath attribute for the Add element. For download: local content
        folder path. For install: ".".

    .PARAMETER Channel
        ODT channel name. Valid values: MonthlyEnterprise, Current.
        Default: MonthlyEnterprise.

    .PARAMETER CompanyName
        Value for the AppSettings Company name. Omit or pass empty to skip
        the AppSettings block entirely.
    #>
    param(
        [Parameter(Mandatory)][string]$OfficeClientEdition,
        [string]$Version,
        [Parameter(Mandatory)][string[]]$ProductIds,
        [string]$SourcePath,
        [ValidateSet('MonthlyEnterprise','Current')]
        [string]$Channel = 'MonthlyEnterprise',
        [string]$CompanyName,
        # ExcludeApp IDs per product (e.g. 'Groove','Lync','Teams'). Defaults
        # preserved from the prior hardcoded list so pre-pref callers behave
        # identically. Pass @() to include everything.
        [string[]]$ExcludeApps = @('Groove','Lync','OneDrive','Teams','Bing')
    )

    $addAttrs = @()
    if ($SourcePath) {
        $addAttrs += 'SourcePath="{0}"' -f (ConvertTo-XmlAttributeValue $SourcePath)
    }
    $addAttrs += 'OfficeClientEdition="{0}"' -f (ConvertTo-XmlAttributeValue $OfficeClientEdition)
    $addAttrs += 'Channel="{0}"' -f (ConvertTo-XmlAttributeValue $Channel)
    $addAttrs += 'OfficeMgmtCOM="TRUE"'
    if ($Version -and $Version -ne 'Latest') {
        $addAttrs += 'Version="{0}"' -f (ConvertTo-XmlAttributeValue $Version)
    }
    $addAttrs += 'MigrateArch="TRUE"'

    $lines = @('<Configuration>')
    $lines += '  <Add {0}>' -f ($addAttrs -join ' ')

    foreach ($prodId in $ProductIds) {
        $lines += '    <Product ID="{0}">' -f (ConvertTo-XmlAttributeValue $prodId)
        $lines += '      <Language ID="en-us" />'
        foreach ($appId in $ExcludeApps) {
            if (-not [string]::IsNullOrWhiteSpace($appId)) {
                $lines += '      <ExcludeApp ID="{0}" />' -f (ConvertTo-XmlAttributeValue $appId)
            }
        }
        $lines += '    </Product>'
    }

    $lines += '  </Add>'
    $lines += '  <Property Name="SharedComputerLicensing" Value="1" />'
    $lines += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
    $lines += '  <Property Name="DeviceBasedLicensing" Value="0" />'
    $lines += '  <Property Name="PinIconsToTaskbar" Value="FALSE" />'
    $lines += '  <Property Name="SCLCacheOverride" Value="0" />'
    $lines += '  <RemoveMSI />'
    if ($CompanyName) {
        $lines += '  <AppSettings>'
        $lines += '    <Setup Name="Company" Value="{0}" />' -f (ConvertTo-XmlAttributeValue $CompanyName)
        $lines += '  </AppSettings>'
    }
    $lines += '  <Display Level="None" AcceptEULA="TRUE" />'
    $lines += '  <Logging Level="Standard" Path="%programdata%\Appdeploy\Office2016" />'
    $lines += '</Configuration>'

    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Java vendor release helpers
# ---------------------------------------------------------------------------

function Get-LatestTemurinRelease {
    <#
    .SYNOPSIS
        Queries the Eclipse Adoptium API for the latest Temurin release.
    .DESCRIPTION
        Returns a hashtable with Version, DownloadUrl, and FileName for the
        latest Temurin JRE or JDK MSI installer. The -LTS suffix is stripped
        from the version string.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER ImageType
        'jre' or 'jdk'.
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [Parameter(Mandatory)][ValidateSet('jre','jdk')][string]$ImageType,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.adoptium.net/v3/assets/latest/$FeatureVersion/hotspot?architecture=$Architecture&image_type=$ImageType&os=windows"
    Write-Log "Adoptium API URL             : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Adoptium API." }

        $data = ConvertFrom-Json $json

        $asset = $data | Where-Object { $_.binary.installer.name -match '\.msi$' } | Select-Object -First 1
        if (-not $asset) { throw "No MSI installer found for Temurin $ImageType $FeatureVersion ($Architecture)." }

        $downloadUrl = $asset.binary.installer.link
        $fileName    = $asset.binary.installer.name
        $rawVersion  = $asset.version.semver

        if ([string]::IsNullOrWhiteSpace($rawVersion)) { throw "version.semver is empty in Adoptium API response." }

        $version = $rawVersion -replace '[\.\-]\d*\.?LTS$', ''

        Write-Log ("Temurin {0} {1} version      : {2}" -f $ImageType, $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Temurin release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}


function Get-LatestCorrettoRelease {
    <#
    .SYNOPSIS
        Queries the GitHub API for the latest Amazon Corretto JDK release.
    .DESCRIPTION
        Returns a hashtable with Version (4-part normalized), DownloadUrl, and
        FileName. Corretto uses 5-part versioning; the 5th part (Corretto patch)
        is stripped to produce a 4-part version compatible with the GUI regex
        and the .NET [version] type.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.github.com/repos/corretto/corretto-$FeatureVersion/releases/latest"
    Write-Log "Corretto GitHub API URL      : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error -A "PowerShell" $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Corretto GitHub API." }

        $release = ConvertFrom-Json $json

        $tagVersion = $release.tag_name
        if ([string]::IsNullOrWhiteSpace($tagVersion)) { throw "tag_name is empty in Corretto release response." }

        # Construct MSI filename from known pattern
        # v8: amazon-corretto-{TAG}-windows-{ARCH}-jdk.msi
        # v11+: amazon-corretto-{TAG}-windows-{ARCH}.msi
        if ($FeatureVersion -le 8) {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture-jdk.msi"
        }
        else {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture.msi"
        }
        $downloadUrl = "https://corretto.aws/downloads/resources/$tagVersion/$fileName"

        # Normalize 5-part version to 4 parts (strip Corretto patch)
        $parts = $tagVersion -split '\.'
        if ($parts.Count -ge 5) {
            $version = ($parts[0..3] -join '.')
        }
        else {
            $version = $tagVersion
        }

        Write-Log ("Corretto {0} raw version     : {1}" -f $FeatureVersion, $tagVersion) -Quiet:$Quiet
        Write-Log ("Corretto {0} normalized      : {1}" -f $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Corretto release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}

# ---------------------------------------------------------------------------
# Packager history (per-app timestamps for batch/scheduled workflow)
# ---------------------------------------------------------------------------
# Storage lives under %LOCALAPPDATA%\AppPackager\app-history.json. It is NEVER
# tracked in the repo (per the no-JSON rule) and is auto-created on first use.
#
# Schema (top level is a dictionary keyed by packager base name, e.g. 'package-chrome'):
# {
#   "package-chrome": {
#     "LastChecked":      "2026-04-19T17:00:00Z",
#     "LastStaged":       "2026-04-19T17:05:00Z",
#     "LastPackaged":     "2026-04-19T17:10:00Z",
#     "LastKnownVersion": "140.0.7339.80",
#     "LastResult":       "Updated"     # Checked | NoChange | Updated | Failed
#   },
#   ...
# }

function Get-PackagerHistoryPath {
    <#
    .SYNOPSIS
        Returns the user-profile path where packager history is persisted.
    #>
    $dir = Join-Path $env:LOCALAPPDATA 'AppPackager'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Join-Path $dir 'app-history.json'
}

function Convert-PackagerHistoryTimestampToString {
    param($Value)

    if ($Value -is [datetime]) {
        return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    return $Value
}

function Read-PackagerHistory {
    <#
    .SYNOPSIS
        Loads the packager history dictionary. Returns an empty hashtable if
        the file doesn't exist or is unreadable.
    #>
    $path = Get-PackagerHistoryPath
    if (-not (Test-Path -LiteralPath $path)) { return @{} }
    try {
        $raw  = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
        $obj  = $raw | ConvertFrom-Json -ErrorAction Stop
        $hash = @{}
        foreach ($p in $obj.PSObject.Properties) {
            $entry = $p.Value
            if ($entry -and $entry.PSObject -and $entry.PSObject.Properties) {
                foreach ($datePropName in @('LastChecked','LastStaged','LastPackaged')) {
                    $dateProp = $entry.PSObject.Properties[$datePropName]
                    if ($dateProp) {
                        $dateProp.Value = Convert-PackagerHistoryTimestampToString -Value $dateProp.Value
                    }
                }
            }
            $hash[$p.Name] = $entry
        }
        return $hash
    }
    catch {
        Write-Log ("Read-PackagerHistory: could not read {0}: {1}" -f $path, $_.Exception.Message) -Level WARN
        return @{}
    }
}

function Save-PackagerHistory {
    <#
    .SYNOPSIS
        Writes the history dictionary back to disk as UTF-8 JSON.
    .PARAMETER History
        Hashtable keyed by packager base name.
    #>
    param([Parameter(Mandatory)][hashtable]$History)
    $path = Get-PackagerHistoryPath
    try {
        ($History | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $path -Encoding UTF8
    }
    catch {
        Write-Log ("Save-PackagerHistory: could not write {0}: {1}" -f $path, $_.Exception.Message) -Level ERROR
        throw
    }
}

function Update-PackagerHistory {
    <#
    .SYNOPSIS
        Records an event against a packager in the history file.
    .PARAMETER PackagerName
        Script base name without extension, e.g. 'package-chrome'.
    .PARAMETER Event
        What happened: Checked | Staged | Packaged
    .PARAMETER Version
        Version string associated with the event (the latest version checked,
        staged, or packaged). Optional for non-version events.
    .PARAMETER Result
        Outcome classification used by the summary: Checked | NoChange |
        Updated | Failed. Optional.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagerName,
        [Parameter(Mandatory)][ValidateSet('Checked','Staged','Packaged')][string]$Event,
        [string]$Version,
        [ValidateSet('Checked','NoChange','Updated','Failed')][string]$Result
    )

    $history = Read-PackagerHistory
    if (-not $history.ContainsKey($PackagerName)) {
        $history[$PackagerName] = [pscustomobject]@{
            LastChecked      = $null
            LastStaged       = $null
            LastPackaged     = $null
            LastKnownVersion = $null
            LastResult       = $null
        }
    }
    $entry = $history[$PackagerName]

    # PSCustomObject from ConvertFrom-Json is read-only for property-set; normalize to hashtable-like.
    if ($entry -isnot [hashtable]) {
        $h = @{}
        foreach ($p in $entry.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $entry = $h
        $history[$PackagerName] = $entry
    }

    $nowUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    switch ($Event) {
        'Checked'  { $entry['LastChecked']  = $nowUtc }
        'Staged'   { $entry['LastStaged']   = $nowUtc }
        'Packaged' { $entry['LastPackaged'] = $nowUtc }
    }
    if ($PSBoundParameters.ContainsKey('Version') -and -not [string]::IsNullOrWhiteSpace($Version)) {
        $entry['LastKnownVersion'] = $Version
    }
    if ($PSBoundParameters.ContainsKey('Result')) {
        $entry['LastResult'] = $Result
    }

    Save-PackagerHistory -History $history
}

# ---------------------------------------------------------------------------
# Module export (belt-and-suspenders with .psd1 FunctionsToExport)
# ---------------------------------------------------------------------------

Export-ModuleMember -Function *
