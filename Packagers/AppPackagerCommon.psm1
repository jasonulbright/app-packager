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
# Shared core
# ---------------------------------------------------------------------------

# SuiteCommon (vendored at ..\Lib\SuiteCommon) supplies logging, settings,
# and the CM connection core. Guarded, not -Force: a -Force reimport of
# this module (background runspace init) must not reset SuiteCommon's
# attached logging state.
if (-not (Get-Module -Name SuiteCommon)) {
    Import-Module -Name (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
}
# Packager scripts and operators set APP_PACKAGER_VERBOSE; SuiteCommon
# gates DEBUG on SUITE_VERBOSE. Bridge once at import.
if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_VERBOSE) -and [string]::IsNullOrWhiteSpace($env:SUITE_VERBOSE)) {
    $env:SUITE_VERBOSE = $env:APP_PACKAGER_VERBOSE
}

# ---------------------------------------------------------------------------
# Download with retry
# ---------------------------------------------------------------------------

function Invoke-DownloadWithRetry {
    <#
    .SYNOPSIS
        Downloads a file via curl.exe, falling back to Invoke-WebRequest, with
        a single retry on failure.

    .DESCRIPTION
        curl.exe is primary (curl.exe -L --fail --silent --show-error -o <file> <url>):
        the in-box Schannel build trusts the Windows certificate store and
        negotiates modern TLS regardless of per-machine .NET registry state.

        When curl fails, the attempt falls back to Invoke-WebRequest, which
        discovers WinINET/system proxy settings curl does not read. The
        fallback sends default credentials to the system proxy (Kerberos/NTLM
        auth proxies) and suppresses the 5.1 progress bar that cripples
        download throughput. TLS 1.2 is forced at module import.

        Throws on final failure (both methods, all attempts).

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
            Write-Log ("Retrying download (attempt {0} of {1}) after {2}s delay..." -f $attempt, $maxAttempts, $RetryDelaySec) -Level WARN -Quiet:$Quiet
            Start-Sleep -Seconds $RetryDelaySec
        }

        $allArgs = @('-L', '--fail', '--silent', '--show-error') + $ExtraCurlArgs + @('-o', $OutFile, $Url)
        & curl.exe @allArgs 2>$null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            return
        }

        Write-Log ("curl.exe failed (exit code {0}); falling back to Invoke-WebRequest." -f $exitCode) -Level WARN -Quiet:$Quiet

        # curl leaves a partial file when the transfer dies mid-stream; remove
        # it so a fallback/retry never passes integrity checks with torn content.
        if (Test-Path -LiteralPath $OutFile) {
            Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
        }

        $savedProgress = $ProgressPreference
        try {
            # Authenticated system proxies (the environments where curl fails)
            # reject anonymous CONNECT; hand the default proxy our credentials.
            [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

            # 5.1 redraws the progress bar per buffer read; silencing it is the
            # difference between KB/s and full line speed on large installers.
            $ProgressPreference = 'SilentlyContinue'

            Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop

            Write-Log "Download succeeded via Invoke-WebRequest fallback." -Quiet:$Quiet
            return
        }
        catch {
            Write-Log ("Invoke-WebRequest fallback failed: {0}" -f $_.Exception.Message) -Level WARN -Quiet:$Quiet
            if (Test-Path -LiteralPath $OutFile) {
                Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
            }
        }
        finally {
            $ProgressPreference = $savedProgress
        }

        if ($attempt -lt $maxAttempts) {
            Write-Log ("Download attempt {0} failed via both curl.exe and Invoke-WebRequest. Will retry." -f $attempt) -Level WARN -Quiet:$Quiet
        }
    }

    $msg = "Download failed after $maxAttempts attempt(s) via both curl.exe and Invoke-WebRequest: $Url"
    Write-Log $msg -Level ERROR -Quiet:$Quiet
    throw $msg
}

# ---------------------------------------------------------------------------
# TLS 1.2 enforcement
# ---------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-NetworkContentPath {
    <#
    .SYNOPSIS
        Creates and returns the network content folder for one application
        version in the configured share layout.

    .DESCRIPTION
        Nested (default): <FileServerPath>\Applications\<Vendor>\<App>\<Version>
        Flat:             <FileServerPath>\Applications\<Vendor>-<App>-<Version>

        Both layouts live under Applications\. Nested keeps an app's version
        folders adjacent (retention pruning deletes old version folders in
        place); Flat serves org conventions that mandate one folder per
        package. The layout choice must stay consistent across a site: mixing
        them leaves content split across two trees.
    #>
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder,
        [Parameter(Mandatory)][string]$Version,
        [ValidateSet('Nested', 'Flat')][string]$Layout = 'Nested'
    )

    if ($Layout -eq 'Flat') {
        $appsRoot    = Join-Path $FileServerPath 'Applications'
        $contentPath = Join-Path $appsRoot ('{0}-{1}-{2}' -f $VendorFolder, $AppFolder, $Version)
        Initialize-Folder -Path $appsRoot
        Initialize-Folder -Path $contentPath
        return $contentPath
    }

    $appRoot     = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $contentPath = Join-Path $appRoot $Version
    Initialize-Folder -Path $contentPath
    return $contentPath
}

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


function Connect-CMSite {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = $null
    )
    # Resolution stays app-side: preferences, the AdminUI connect script,
    # and APP_PACKAGER_CM_PROVIDER feed Resolve-CMProviderMachineName; the
    # drive and session mechanics live in SuiteCommon. Site verification
    # stays off: packager runs never queried the site on connect.
    $provider = Resolve-CMProviderMachineName -ProviderMachineName $ProviderMachineName
    if ([string]::IsNullOrWhiteSpace($provider) -and
        -not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        Write-Log "Configuration Manager PSDrive '$SiteCode' is not available and no provider machine name is configured. Set Provider Machine in MECM Preferences, copy the ProviderMachineName value from the AdminUI connect script, or set APP_PACKAGER_CM_PROVIDER." -Level ERROR
        return $false
    }
    return (SuiteCommon\Connect-CMSite -SiteCode $SiteCode -SMSProvider $provider -SkipSiteVerification)
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

    # A build staged under operator command overrides records them, so the
    # manifest distinguishes it from a stock build.
    $stageOverrides = Get-RequestedCommandOverrides
    if ($stageOverrides) {
        $ManifestData['CommandOverrides'] = @{
            Install   = [string]$stageOverrides.Install
            Uninstall = [string]$stageOverrides.Uninstall
        }
        Write-Log "Command overrides recorded in manifest."
    }

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
        Builds the path <FileServerPath>\Applications\<VendorFolder>\<AppFolder>,
        creating each level if it does not exist. Returns the final path.
    #>
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder
    )

    $appsRoot   = Join-Path $FileServerPath "Applications"
    $vendorPath = Join-Path $appsRoot $VendorFolder
    $appPath    = Join-Path $vendorPath $AppFolder

    Initialize-Folder -Path $appsRoot
    Initialize-Folder -Path $vendorPath
    Initialize-Folder -Path $appPath

    return $appPath
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

    # The filename lands inside a single-quoted literal in the generated
    # script; an unescaped apostrophe terminates the string early.
    $MsiFileName = $MsiFileName -replace "'", "''"

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

    # Filename and uninstall path land inside single-quoted literals in the
    # generated script; an unescaped apostrophe terminates the string early.
    # InstallArgs/UninstallArgs are excluded: they are caller-authored
    # PowerShell element lists interpolated into @() as-is.
    $InstallerFileName = $InstallerFileName -replace "'", "''"
    $UninstallCommand  = $UninstallCommand -replace "'", "''"

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

function Test-PsadtLayout {
    <#
    .SYNOPSIS
        Detects the PSADT toolkit generation in a folder and returns the MECM
        deployment type command lines for it.

    .DESCRIPTION
        v4 layouts carry Invoke-AppDeployToolkit.ps1 at the root (usually with
        Invoke-AppDeployToolkit.exe beside it); v3 layouts carry
        Deploy-Application.exe / Deploy-Application.ps1. The native .exe
        launcher is preferred when present because it hides the PowerShell
        window; otherwise the command line falls back to powershell.exe -File.

        DeployMode is deliberately omitted from the generated command lines by
        default: PSADT runs Interactive when a user session exists and degrades
        to NonInteractive otherwise, which is the close-app/defer behavior that
        justifies wrapping with PSADT at all. Pass -DeployMode to force one.

    .OUTPUTS
        [pscustomobject] Generation ('v3'|'v4'), EntryPoint (file name),
        InstallCommandLine, UninstallCommandLine.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('', 'Interactive', 'Silent', 'NonInteractive')]
        [string]$DeployMode = ''
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "PSADT toolkit folder not found: $Path"
    }

    $modeSuffix = if ([string]::IsNullOrWhiteSpace($DeployMode)) { '' } else { " -DeployMode $DeployMode" }

    $v4Exe = Join-Path $Path 'Invoke-AppDeployToolkit.exe'
    $v4Ps1 = Join-Path $Path 'Invoke-AppDeployToolkit.ps1'
    $v3Exe = Join-Path $Path 'Deploy-Application.exe'
    $v3Ps1 = Join-Path $Path 'Deploy-Application.ps1'

    if (Test-Path -LiteralPath $v4Ps1) {
        if (Test-Path -LiteralPath $v4Exe) {
            $entry = 'Invoke-AppDeployToolkit.exe'
            $prefix = $entry
        }
        else {
            $entry = 'Invoke-AppDeployToolkit.ps1'
            $prefix = "powershell.exe -NonInteractive -ExecutionPolicy Bypass -File `"$entry`""
        }
        return [pscustomobject]@{
            Generation           = 'v4'
            EntryPoint           = $entry
            InstallCommandLine   = "$prefix -DeploymentType Install$modeSuffix"
            UninstallCommandLine = "$prefix -DeploymentType Uninstall$modeSuffix"
        }
    }

    if ((Test-Path -LiteralPath $v3Exe) -or (Test-Path -LiteralPath $v3Ps1)) {
        if (Test-Path -LiteralPath $v3Exe) {
            $entry = 'Deploy-Application.exe'
            $prefix = $entry
        }
        else {
            $entry = 'Deploy-Application.ps1'
            $prefix = "powershell.exe -NonInteractive -ExecutionPolicy Bypass -File `"$entry`""
        }
        return [pscustomobject]@{
            Generation           = 'v3'
            EntryPoint           = $entry
            InstallCommandLine   = "$prefix -DeploymentType `"Install`"$modeSuffix"
            UninstallCommandLine = "$prefix -DeploymentType `"Uninstall`"$modeSuffix"
        }
    }

    throw "No PSADT entry point found in '$Path' (expected Invoke-AppDeployToolkit.ps1 [v4] or Deploy-Application.exe/.ps1 [v3])."
}

function Get-NextPatchVersion {
    <#
    .SYNOPSIS
        Returns the version string with its last dotted numeric component
        incremented by one ('8.0.20' -> '8.0.21').
    .DESCRIPTION
        Used by packagers whose detection accepts the packaged version OR its
        immediate successor, so last month's still-active deployment stops
        reinstalling over an in-place upgrade. Returns $null when the last
        component is not purely numeric (e.g. preview/rc suffixes), so callers
        can fall back to single-version detection.
    #>
    param([Parameter(Mandatory)][string]$Version)

    # Every dotted component must be purely numeric; '10.0.0-rc.1' splits to a
    # numeric final segment and must not be treated as an incrementable patch.
    $parts = $Version -split '\.'
    $n = 0
    foreach ($p in $parts) {
        if (-not [int]::TryParse($p, [ref]$n)) { return $null }
    }
    $parts[$parts.Count - 1] = [string]([int]$parts[$parts.Count - 1] + 1)
    return ($parts -join '.')
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


function Get-ConditionTemplatesPath {
    return (Join-Path $PSScriptRoot 'condition-templates.json')
}

function Get-DefaultConditionTemplates {
    <#
    .SYNOPSIS
        Returns the built-in condition template document.

    .DESCRIPTION
        Each template maps a stable Id (referenced by requirement specs) to
        a CM global condition. Kind selects the creation cmdlet (BuiltIn is
        never created, only resolved); RuleType selects the requirement
        rule cmdlet the spec value binds through.
    #>
    return [pscustomobject]@{
        SchemaVersion = 1
        Conditions    = @(
            [pscustomobject]@{
                Id                  = 'cpu-arch'
                GlobalConditionName = 'APKG - CPU Architecture'
                Kind                = 'Wql'
                RuleType            = 'CommonValue'
                DataType            = 'Integer'
                Namespace           = 'root\cimv2'
                Class               = 'Win32_Processor'
                Property            = 'Architecture'
                # Win32_OperatingSystem.OSArchitecture returns a localized
                # string; the numeric processor architecture compares the
                # same on every OS language.
                Values              = [pscustomobject]@{ x64 = '9'; ARM64 = '12' }
                Description         = 'Win32_Processor.Architecture (9 = x64, 12 = ARM64).'
            }
            [pscustomobject]@{
                Id                  = 'os-language'
                GlobalConditionName = 'Operating System Language'
                Kind                = 'BuiltIn'
                RuleType            = 'OperatingSystemLanguage'
                # The site ships two conditions under this name; PlatformType
                # 1 is Windows, 2 is Mobile.
                PlatformType        = 1
            }
            [pscustomobject]@{
                Id                  = 'vpn-connected'
                GlobalConditionName = 'APKG - VPN Connected'
                Kind                = 'Script'
                RuleType            = 'Boolean'
                DataType            = 'Boolean'
                AdapterPatterns     = @('Zscaler', 'Juniper', 'PANGP', 'Cisco AnyConnect', 'Fortinet')
                AliasPattern        = 'vpn'
                Description         = 'True when an IP-enabled network adapter description matches a configured VPN client pattern or an interface alias contains the alias pattern.'
            }
        )
    }
}

function Get-ConditionTemplates {
    <#
    .SYNOPSIS
        Returns the condition template document (file override or built-in defaults).

    .DESCRIPTION
        condition-templates.json beside the module overrides the built-in
        defaults when present and parseable; the Deployment Conditions
        options panel writes it. A missing or malformed file falls back to
        defaults so the Package phase never depends on the file existing.
    #>
    $path = Get-ConditionTemplatesPath
    if (Test-Path -LiteralPath $path) {
        try {
            $doc = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($doc -and $doc.Conditions) { return $doc }
            Write-Log ("Condition templates file has no Conditions; using built-in defaults: {0}" -f $path) -Level WARN
        }
        catch {
            Write-Log ("Condition templates file unreadable; using built-in defaults: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
    return Get-DefaultConditionTemplates
}

function Save-ConditionTemplates {
    param([Parameter(Mandatory)][pscustomobject]$Templates)
    $path = Get-ConditionTemplatesPath
    $Templates | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function New-VpnConditionScriptText {
    <#
    .SYNOPSIS
        Builds the client-side detection script for the VPN-connected global condition.

    .DESCRIPTION
        The generated script runs on clients inside the CM agent's
        requirement evaluation, so it must emit exactly one boolean and
        nothing else on the output stream.
    #>
    param(
        [string[]]$AdapterPatterns = @(),
        [string]$AliasPattern = ''
    )

    $quoted = @($AdapterPatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        "'" + ([string]$_ -replace "'", "''") + "'"
    })

    $lines = @(
        ('$patterns = @({0})' -f ($quoted -join ', '))
        '$found = $false'
        'foreach ($adapter in @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=''True''" -ErrorAction SilentlyContinue)) {'
        '    foreach ($p in $patterns) {'
        '        if ($adapter.Description -like (''*'' + $p + ''*'')) { $found = $true }'
        '    }'
        '}'
    )
    if (-not [string]::IsNullOrWhiteSpace($AliasPattern)) {
        $aliasLike = '*' + ([string]$AliasPattern -replace "'", "''") + '*'
        $lines += @(
            'if (-not $found) {'
            ('    foreach ($ip in @(Get-CimInstance -ClassName MSFT_NetIPAddress -Namespace ''root/StandardCimv2'' -ErrorAction SilentlyContinue)) {')
            ('        if ($ip.InterfaceAlias -like ''{0}'') {{ $found = $true }}' -f $aliasLike)
            '    }'
            '}'
        )
    }
    $lines += '$found'
    return ($lines -join "`r`n")
}

function Get-OrCreateGlobalConditionFromTemplate {
    <#
    .SYNOPSIS
        Resolves a condition template to a CM global condition, creating it when absent.

    .DESCRIPTION
        Must be called from the CM site drive. Matching is by name, so an
        existing site condition of any origin is attached rather than
        duplicated; renaming a template's GlobalConditionName points it at
        a site's own condition.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Template)

    $name = [string]$Template.GlobalConditionName
    if ([string]::IsNullOrWhiteSpace($name)) {
        throw "Condition template '$($Template.Id)' has no GlobalConditionName."
    }

    # A name can match several site conditions (the stock Operating System
    # Language exists once per PlatformType: 1 = Windows, 2 = Mobile).
    # Passing an array to a requirement rule cmdlet's -InputObject fails,
    # so narrow by the template's PlatformType and require a single match.
    $existing = @(Get-CMGlobalCondition -Name $name)
    if ($existing.Count -gt 1 -and $Template.PSObject.Properties['PlatformType'] -and $Template.PlatformType) {
        $existing = @($existing | Where-Object { [int]$_.PlatformType -eq [int]$Template.PlatformType })
    }
    if ($existing.Count -gt 1) {
        throw "Global condition name '$name' matches $($existing.Count) site conditions; add a PlatformType to the '$($Template.Id)' template or rename to a unique condition."
    }
    if ($existing.Count -eq 1) {
        Write-Log ("Global condition (existing)  : {0}" -f $name)
        return $existing[0]
    }

    switch ([string]$Template.Kind) {
        'BuiltIn' {
            throw "Built-in global condition '$name' was not found on the site."
        }
        'Wql' {
            Write-Log ("Global condition (creating)  : {0}" -f $name)
            return New-CMGlobalConditionWqlQuery `
                -Name $name `
                -DataType ([string]$Template.DataType) `
                -Namespace ([string]$Template.Namespace) `
                -Class ([string]$Template.Class) `
                -Property ([string]$Template.Property) `
                -Description ([string]$Template.Description)
        }
        'Script' {
            Write-Log ("Global condition (creating)  : {0}" -f $name)
            $scriptText = ''
            if ($Template.PSObject.Properties['ScriptText'] -and $Template.ScriptText) {
                $scriptText = (@($Template.ScriptText) -join "`r`n")
            }
            else {
                $scriptText = New-VpnConditionScriptText -AdapterPatterns @($Template.AdapterPatterns) -AliasPattern ([string]$Template.AliasPattern)
            }
            return New-CMGlobalConditionScript `
                -Name $name `
                -DataType ([string]$Template.DataType) `
                -ScriptLanguage PowerShell `
                -ScriptText $scriptText `
                -Description ([string]$Template.Description)
        }
        default {
            throw "Condition template '$($Template.Id)' has unsupported Kind '$($Template.Kind)'."
        }
    }
}

function Get-DeploymentTypeRequirementSpecs {
    <#
    .SYNOPSIS
        Collects requirement rule specs for the deployment type being created.

    .DESCRIPTION
        Two sources merge: a manifest Requirements array and the
        APP_PACKAGER_REQUIREMENTS environment JSON the GUI sets per app
        ({ SchemaVersion, Rules: [...] }). Malformed environment JSON
        throws instead of packaging without the rules the operator
        configured.
    #>
    param(
        [pscustomobject]$Manifest,
        [switch]$IgnoreEnvironment
    )

    $specs = @()
    if ($Manifest -and $Manifest.PSObject.Properties['Requirements'] -and $Manifest.Requirements) {
        $specs += @($Manifest.Requirements)
    }

    $envJson = $(if ($IgnoreEnvironment) { $null } else { $env:APP_PACKAGER_REQUIREMENTS })
    if (-not [string]::IsNullOrWhiteSpace($envJson)) {
        try {
            $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "APP_PACKAGER_REQUIREMENTS is not valid JSON: $($_.Exception.Message)"
        }
        if ($parsed -and $parsed.PSObject.Properties['Rules'] -and $parsed.Rules) {
            $specs += @($parsed.Rules)
        }
    }

    # No comma wrap: an empty array must unroll to zero pipeline objects
    # so callers' @(...) sees Count 0, not a one-element array-of-array.
    return $specs
}

function New-DeploymentTypeRequirementRules {
    <#
    .SYNOPSIS
        Builds CM requirement rule objects from condition templates and specs.

    .DESCRIPTION
        Must be called from the CM site drive. Unknown condition ids and
        unmapped values throw so a package run never creates a deployment
        type missing rules the operator asked for.

    .OUTPUTS
        Requirement rule array; empty when no specs are configured.
    #>
    param(
        [pscustomobject]$Manifest,
        [switch]$IgnoreEnvironment
    )

    $specs = @(Get-DeploymentTypeRequirementSpecs -Manifest $Manifest -IgnoreEnvironment:$IgnoreEnvironment)
    if ($specs.Count -eq 0) { return @() }

    $doc = Get-ConditionTemplates
    $rules = @()
    foreach ($spec in $specs) {
        $condId = [string]$spec.ConditionId
        $found = @($doc.Conditions | Where-Object { [string]$_.Id -eq $condId })
        if ($found.Count -eq 0) {
            throw "Requirement rule references unknown condition id '$condId'."
        }
        $template = $found[0]
        $gc = Get-OrCreateGlobalConditionFromTemplate -Template $template

        switch ([string]$template.RuleType) {
            'CommonValue' {
                $value1 = [string]$spec.Value
                if ($template.PSObject.Properties['Values'] -and $template.Values) {
                    $mapped = $template.Values.PSObject.Properties[$value1]
                    if (-not $mapped) {
                        $known = (@($template.Values.PSObject.Properties | ForEach-Object { $_.Name }) -join ', ')
                        throw "Condition '$condId' has no mapping for value '$value1' (known: $known)."
                    }
                    $value1 = [string]$mapped.Value
                }
                if ([string]::IsNullOrWhiteSpace($value1)) {
                    throw "Condition '$condId' requires a Value."
                }
                $rules += New-CMRequirementRuleCommonValue -InputObject $gc -RuleOperator IsEquals -Value1 $value1
            }
            'OperatingSystemLanguage' {
                $cultures = @()
                foreach ($c in @($spec.Cultures)) {
                    if ([string]::IsNullOrWhiteSpace([string]$c)) { continue }
                    $cultures += [System.Globalization.CultureInfo]::GetCultureInfo([string]$c)
                }
                if ($cultures.Count -eq 0) {
                    throw "Condition '$condId' requires at least one culture code (e.g. de-DE)."
                }
                $rules += New-CMRequirementRuleOperatingSystemLanguageValue -InputObject $gc -RuleOperator OneOf -Culture $cultures
            }
            'Boolean' {
                # [bool]'false' is $true; Convert.ToBoolean parses the
                # string form and throws on anything else.
                $boolValue = $false
                if ($spec.Value -is [bool]) { $boolValue = $spec.Value }
                else { $boolValue = [System.Convert]::ToBoolean([string]$spec.Value) }
                $rules += New-CMRequirementRuleBooleanValue -InputObject $gc -Value $boolValue
            }
            default {
                throw "Condition '$condId' has unsupported RuleType '$($template.RuleType)'."
            }
        }
        Write-Log ("Requirement rule             : {0} ({1})" -f $condId, $template.GlobalConditionName)
    }

    return $rules
}


function Get-ManifestDeploymentTypeSpecs {
    <#
    .SYNOPSIS
        Resolves a stage manifest into an ordered deployment type spec list.

    .DESCRIPTION
        A manifest without a DeploymentTypes array yields one spec built
        from the base fields (today's single-DT behavior, byte for byte).
        A DeploymentTypes array yields one spec per entry in manifest
        order; CM assigns deployment type priority by creation order and
        the client installs the first deployment type whose requirements
        pass, so authors list the most specific variant first and the
        unconditional fallback last. Entry fields override the base
        manifest field of the same name; ContentSubpath is relative to
        the network content root. On a multi-DT manifest the
        APP_PACKAGER_REQUIREMENTS environment JSON is ignored (per-DT
        Requirements arrays are authoritative) - the caller logs that.

    .OUTPUTS
        [pscustomobject[]] DtName, ContentLocation, InstallCommand,
        UninstallCommand, Detection, RequirementSource, plus behavior
        overrides (PostExecutionBehavior, InstallationBehaviorType,
        LogonRequirementType, RequireUserInteraction).
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [Parameter(Mandatory)][string]$AppName
    )

    $entries = @()
    if ($Manifest.PSObject.Properties['DeploymentTypes'] -and $Manifest.DeploymentTypes) {
        $entries = @($Manifest.DeploymentTypes)
    }

    $base = {
        param($Entry, $Field)
        if ($Entry -and $Entry.PSObject.Properties[$Field] -and $null -ne $Entry.$Field -and ('' -ne [string]$Entry.$Field -or $Entry.$Field -is [bool])) { return $Entry.$Field }
        if ($Manifest.PSObject.Properties[$Field]) { return $Manifest.$Field }
        return $null
    }

    if ($entries.Count -eq 0) {
        # RequirementSource $Manifest keeps the single-DT contract: manifest
        # Requirements plus the environment JSON, exactly as before.
        return ,([pscustomobject]@{
            DtName                   = $AppName
            ContentLocation          = $NetworkContentPath
            InstallCommand           = $(if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.InstallCommandLine)) { [string]$Manifest.InstallCommandLine } else { 'install.bat' })
            UninstallCommand         = $(if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.UninstallCommandLine)) { [string]$Manifest.UninstallCommandLine } else { 'uninstall.bat' })
            Detection                = $Manifest.Detection
            RequirementSource        = $Manifest
            PostExecutionBehavior    = $Manifest.PostExecutionBehavior
            InstallationBehaviorType = $Manifest.InstallationBehaviorType
            LogonRequirementType     = $Manifest.LogonRequirementType
            RequireUserInteraction   = ($Manifest.RequireUserInteraction -eq $true)
        })
    }

    $specs = @()
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $entries) {
        $suffix = [string]$entry.NameSuffix
        if ([string]::IsNullOrWhiteSpace($suffix)) {
            throw "Every DeploymentTypes entry requires a NameSuffix; the deployment type name becomes '<AppName> - <NameSuffix>'."
        }
        $dtName = "$AppName - $suffix"
        if (-not $seen.Add($dtName)) {
            throw "DeploymentTypes entries produce a duplicate deployment type name '$dtName'."
        }

        $content = $NetworkContentPath
        if (-not [string]::IsNullOrWhiteSpace([string]$entry.ContentSubpath)) {
            $content = Join-Path $NetworkContentPath ([string]$entry.ContentSubpath)
        }

        $detection = & $base $entry 'Detection'
        if (-not $detection) {
            throw "DeploymentTypes entry '$suffix' has no Detection and the manifest has no base Detection."
        }

        # Requirements resolve strictly from the entry: a variant without
        # rules is the unconditional fallback by design, and inheriting the
        # base or environment rules here would gate the fallback by accident.
        # @() over a missing property yields @($null) with one element; the
        # Where-Object keeps a rule-less entry a true unconditional fallback.
        $reqSource = [pscustomobject]@{ Requirements = @(@($entry.Requirements) | Where-Object { $null -ne $_ }) }

        $installOverride = [string](& $base $entry 'InstallCommandLine')
        $uninstallOverride = [string](& $base $entry 'UninstallCommandLine')
        $specs += [pscustomobject]@{
            DtName                   = $dtName
            ContentLocation          = $content
            InstallCommand           = $(if (-not [string]::IsNullOrWhiteSpace($installOverride)) { $installOverride } else { 'install.bat' })
            UninstallCommand         = $(if (-not [string]::IsNullOrWhiteSpace($uninstallOverride)) { $uninstallOverride } else { 'uninstall.bat' })
            Detection                = $detection
            RequirementSource        = $reqSource
            PostExecutionBehavior    = (& $base $entry 'PostExecutionBehavior')
            InstallationBehaviorType = (& $base $entry 'InstallationBehaviorType')
            LogonRequirementType     = (& $base $entry 'LogonRequirementType')
            RequireUserInteraction   = ((& $base $entry 'RequireUserInteraction') -eq $true)
        }
    }
    return $specs
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
        [Parameter(Mandatory)][string]$SiteCode,
        [AllowEmptyString()][string]$Comment = '',
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30
    )

    $orig = Get-Location

    $contentVerification = Compare-StageFileHashes -Root $NetworkContentPath -Expected $Manifest.FileHashes
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
        if ([string]::IsNullOrWhiteSpace([string]$appName)) {
            throw "Stage manifest AppName is null or empty; cannot create an MECM application. Re-run the Stage phase and verify the manifest."
        }

        Write-Log ("Manifest fields              : AppName='{0}' Publisher='{1}' SoftwareVersion='{2}' DetectionType='{3}'" -f $appName, $Manifest.Publisher, $Manifest.SoftwareVersion, $Manifest.Detection.Type) -Level DEBUG

        $step = 'Deployment type spec resolution'
        $dtSpecs = @(Get-ManifestDeploymentTypeSpecs -Manifest $Manifest -NetworkContentPath $NetworkContentPath -AppName $appName)
        $isMultiDt = ($dtSpecs.Count -gt 1 -or $dtSpecs[0].DtName -ne $appName)
        if ($isMultiDt) {
            Write-Log ("Deployment types (manifest)  : {0}" -f (($dtSpecs | ForEach-Object { $_.DtName }) -join ', '))
            if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_REQUIREMENTS)) {
                Write-Log "APP_PACKAGER_REQUIREMENTS ignored: a DeploymentTypes manifest carries its own per-deployment-type Requirements." -Level WARN
            }
        }

        # Operator command overrides replace the deployment type commands
        # on single-DT apps; a multi-DT manifest carries per-variant
        # commands, so the flat override is refused rather than applied to
        # every variant.
        $commandOverrides = Get-RequestedCommandOverrides
        if ($commandOverrides) {
            if ($isMultiDt) {
                Write-Log "APP_PACKAGER_COMMANDS ignored: a DeploymentTypes manifest carries per-deployment-type commands." -Level WARN
            }
            else {
                if (-not [string]::IsNullOrWhiteSpace($commandOverrides.Install)) {
                    $dtSpecs[0].InstallCommand = $commandOverrides.Install
                    Write-Log ("Install command (override)   : {0}" -f $commandOverrides.Install)
                }
                if (-not [string]::IsNullOrWhiteSpace($commandOverrides.Uninstall)) {
                    $dtSpecs[0].UninstallCommand = $commandOverrides.Uninstall
                    Write-Log ("Uninstall command (override) : {0}" -f $commandOverrides.Uninstall)
                }
            }
        }

        $step = "Get-CMApplication duplicate check ('$appName')"
        $existing = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
        $cmApp = $null
        $replaceDtNames = @()
        if ($existing) {
            $existingApps = @($existing)
            if ($existingApps.Count -gt 1) {
                throw "Multiple existing MECM applications matched '$appName'; refusing to package until the duplicate names are resolved."
            }
            $cmApp = $existingApps[0]

            $missingDts = @($dtSpecs | Where-Object { -not (Test-MECMApplicationHasDeploymentType -ApplicationName $appName -DeploymentTypeName $_.DtName) })
            $existingVersion = [string]$cmApp.SoftwareVersion

            if ($existingVersion -eq [string]$Manifest.SoftwareVersion) {
                if ($missingDts.Count -eq 0) {
                    Write-Log "Application already exists    : $appName (v$existingVersion, unchanged)" -Level WARN
                    Write-Log ("Deployment type(s) validated : {0}" -f (($dtSpecs | ForEach-Object { $_.DtName }) -join ', '))
                    return [UInt32]$cmApp.CI_ID
                }
                throw ("Existing MECM application '$appName' is missing deployment type(s): {0}. This looks like a partial prior package run; fix or remove the partial app before packaging again." -f (($missingDts | ForEach-Object { $_.DtName }) -join ', '))
            }

            # Version change on a reused application name (version-less CMName by
            # design): replace every deployment type with the new set pointing at
            # the new content. New deployment types are created under staging
            # names first, because a deployed application refuses to remove its
            # last deployment type.
            Write-Log "Application already exists    : $appName (v$existingVersion -> v$($Manifest.SoftwareVersion), replacing deployment types)" -Level WARN
            $step = "Get-CMDeploymentType inventory ('$appName')"
            $replaceDtNames = @(Get-CMDeploymentType -ApplicationName $appName -ErrorAction Stop | ForEach-Object { [string]$_.LocalizedDisplayName })
            if ($replaceDtNames.Count -eq 0) {
                Write-Log "Existing app has no deployment type; adding the new set." -Level WARN
            }
        }

        # Requirement rules resolve for every deployment type before any
        # create/replace so a bad spec fails the run without leaving a
        # partial application behind. Rule objects must be built on the CM
        # drive; they survive the later detection-clause drive switch. On a
        # multi-DT manifest the per-entry Requirements are authoritative and
        # the environment JSON is excluded.
        $step = 'Requirement rules'
        foreach ($spec in $dtSpecs) {
            $spec | Add-Member -NotePropertyName RequirementRules -NotePropertyValue @(New-DeploymentTypeRequirementRules -Manifest $spec.RequirementSource -IgnoreEnvironment:$isMultiDt)
        }

        if (-not $cmApp) {
            Write-Log "Creating CM Application      : $appName"
            $step = "New-CMApplication ('$appName')"
            $cmAppParams = @{
                Name             = $appName
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
        }

        # Create every deployment type in spec order: CM assigns priority by
        # creation order and the client installs the first deployment type
        # whose requirements pass, so the most specific variant is listed
        # first and the unconditional fallback last. On a version replace
        # each new deployment type starts under a staging name; the old set
        # is removed and the staged names take the canonical names only
        # after every new deployment type created successfully.
        $stagedRenames = @()
        foreach ($spec in $dtSpecs) {
            $det = $spec.Detection
            $detType = if ($det.Type) { $det.Type } else { 'RegistryKeyValue' }

            $dtName = $spec.DtName
            $dtCreateName = if ($replaceDtNames.Count -gt 0) { "$dtName (staging)" } else { $dtName }
            # Deployment type command lines default to the generated .bat
            # wrappers; manifests may override both (PSADT-wrapped apps point
            # at the toolkit entry, e.g. Invoke-AppDeployToolkit.exe
            # -DeploymentType Install).
            $installCommand = [string]$spec.InstallCommand
            $uninstallCommand = [string]$spec.UninstallCommand
            if ($installCommand -ne 'install.bat') {
                Write-Log "Install command (manifest)   : $installCommand"
            }
            if ($uninstallCommand -ne 'uninstall.bat') {
                Write-Log "Uninstall command (manifest) : $uninstallCommand"
            }

            $dtParams = @{
                ApplicationName           = $appName
                DeploymentTypeName        = $dtCreateName
                ContentLocation           = $spec.ContentLocation
                InstallCommand            = $installCommand
                UninstallCommand          = $uninstallCommand
                InstallationBehaviorType  = 'InstallForSystem'
                LogonRequirementType      = 'WhetherOrNotUserLoggedOn'
                EstimatedRuntimeMins      = $EstimatedRuntimeMins
                MaximumRuntimeMins        = $MaximumRuntimeMins
                ContentFallback           = $true
                SlowNetworkDeploymentMode = 'Download'
                UserInteractionMode       = 'Hidden'
                ErrorAction               = 'Stop'
            }

            # Manifest field name matches the cmdlet's parameter TYPE
            # (PostExecutionBehavior); the actual parameter name is
            # -RebootBehavior. See Add-CMScriptDeploymentType docs.
            # Default to BasedOnExitCode when the manifest doesn't specify so
            # MSI 3010 "reboot required" exits actually propagate to the CCM
            # client. The cmdlet's own default is NoAction, which silently drops
            # 3010s and breaks install chains that need a reboot between apps.
            if ($spec.PostExecutionBehavior) {
                $dtParams['RebootBehavior'] = $spec.PostExecutionBehavior
            }
            else {
                $dtParams['RebootBehavior'] = 'BasedOnExitCode'
            }

            if ($spec.InstallationBehaviorType) {
                $dtParams['InstallationBehaviorType'] = $spec.InstallationBehaviorType
            }
            if ($spec.LogonRequirementType) {
                $dtParams['LogonRequirementType'] = $spec.LogonRequirementType
            }
            if ($spec.RequireUserInteraction -eq $true) {
                $dtParams['RequireUserInteraction'] = $true
            }

            if ($detType -eq 'Script') {
                # Script-based detection: pass script text, no clause objects needed
                $lang = if ($det.ScriptLanguage) { $det.ScriptLanguage } else { 'PowerShell' }
                $dtParams['ScriptLanguage'] = $lang
                $dtParams['ScriptText']     = $det.ScriptText
            }
            else {
                # Clause-based detection: leave CM PSDrive to create clause objects
                # (CM PSDrive context can interfere with parameter binding)
                $step = "New detection clause(s) (type=$detType, dt='$dtName')"
                Set-Location C: -ErrorAction Stop

                if ($detType -eq 'Compound') {
                    $clauses = @()
                    foreach ($c in $det.Clauses) {
                        $clauses += New-SingleDetectionClause -Det $c
                    }
                    $dtParams['AddDetectionClause'] = $clauses

                    $groupSizes = @()
                    if ($det.PSObject.Properties.Name -contains 'GroupSizes' -and $null -ne $det.GroupSizes) {
                        $groupSizes = @($det.GroupSizes | ForEach-Object { [int]$_ })
                    }

                    if ($groupSizes.Count -gt 0) {
                        # (run1 AND ...) OR (run2 AND ...): exactly two contiguous
                        # clause runs. Only the second run is passed to
                        # -GroupDetectionClauses; the cmdlet's left-associative
                        # expression build parenthesizes the first run on its own,
                        # so grouping both runs is impossible (String[] holds one
                        # group) and unnecessary. OR attaches to the first clause
                        # of the second run; clauses inside a run keep AND.
                        if ($groupSizes.Count -ne 2 -or (($groupSizes[0] + $groupSizes[1]) -ne $clauses.Count) -or ($groupSizes -contains 0)) {
                            throw "Detection.GroupSizes must be exactly two non-zero sizes summing to the clause count (got: '$($groupSizes -join ',')' for $($clauses.Count) clauses)."
                        }
                        $secondStart = $groupSizes[0]
                        $dtParams['GroupDetectionClauses'] = @($clauses[$secondStart..($clauses.Count - 1)] | ForEach-Object { $_.Setting.LogicalName })
                        $dtParams['DetectionClauseConnector'] = @(@{
                            LogicalName = $clauses[$secondStart].Setting.LogicalName
                            Connector   = 'OR'
                        })
                        Write-Log ("Detection expression         : (clauses 1..{0}) OR (clauses {1}..{2})" -f $groupSizes[0], ($secondStart + 1), $clauses.Count)
                    }
                    # OR connector: specify OR for each clause beyond the first
                    # AND is the default and needs no explicit connector
                    elseif ($det.Connector -eq 'Or' -and $clauses.Count -ge 2) {
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
                    $clause = New-SingleDetectionClause -Det $det
                    $dtParams['AddDetectionClause'] = @($clause)
                }

                # Reconnect to CM site for Add-CMScriptDeploymentType
                $step = "Connect-CMSite reconnect (SiteCode=$SiteCode)"
                if (-not (Connect-CMSite -SiteCode $SiteCode)) {
                    throw "CM site reconnection failed."
                }
            }

            $requirementRules = @($spec.RequirementRules)
            if ($requirementRules.Count -gt 0) {
                $dtParams['AddRequirement'] = $requirementRules
                Write-Log ("Requirement rules attached   : {0} ('{1}')" -f $requirementRules.Count, $dtName)
            }

            Write-Log ("Deployment type parameters   : {0}" -f (($dtParams.Keys | Sort-Object | ForEach-Object { "{0}='{1}'" -f $_, $dtParams[$_] }) -join ' ')) -Level DEBUG

            Write-Log "Adding Script Deployment Type : $dtCreateName"
            $step = "Add-CMScriptDeploymentType ('$dtCreateName')"
            Add-CMScriptDeploymentType @dtParams | Out-Null

            if ($dtCreateName -ne $dtName) {
                $stagedRenames += [pscustomobject]@{ From = $dtCreateName; To = $dtName }
            }
        }

        foreach ($oldDt in $replaceDtNames) {
            $step = "Remove-CMDeploymentType ('$oldDt')"
            Remove-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $oldDt -Force -ErrorAction Stop
            Write-Log "Removed old deployment type  : $oldDt"
        }
        foreach ($rename in $stagedRenames) {
            $step = "Set-CMDeploymentType rename ('$($rename.From)' -> '$($rename.To)')"
            Set-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $rename.From -NewDeploymentTypeName $rename.To -ErrorAction Stop
            Write-Log "Renamed deployment type      : $($rename.To)"
        }

        if ($existing) {
            $step = "Set-CMApplication version update ('$appName')"
            $setAppParams = @{
                Name            = $appName
                SoftwareVersion = [string]$Manifest.SoftwareVersion
                ErrorAction     = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($Comment)) { $setAppParams['Description'] = $Comment }
            Set-CMApplication @setAppParams
            Write-Log "Updated application version  : $($Manifest.SoftwareVersion)"
        }

        $step = "Remove-CMApplicationRevisionHistory (CI_ID=$($cmApp.CI_ID))"
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        # Optional auto-distribute to a DP group, and optional test-collection
        # deployment after successful distribution. Settings live in
        # AppPackager.preferences.json alongside the GUI. Packagers invoked
        # from the CLI with no prefs file silently skip both steps.
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

                        # Availability of the test-deployment option is gated in the
                        # GUI (requires auto-distribute + DP group); at runtime the
                        # prefs are taken as-is.
                        $testCollection = if ($null -ne $cd.TestCollectionName) { ([string]$cd.TestCollectionName).Trim() } else { '' }
                        if ($cd.DeployToTestCollection -and -not [string]::IsNullOrWhiteSpace($testCollection)) {
                            try {
                                $collection = Get-CMDeviceCollection -Name $testCollection -ErrorAction SilentlyContinue
                                if (-not $collection -and [bool]$cd.CreateTestCollectionIfMissing) {
                                    Write-Log "Creating test collection     : '$testCollection' (direct membership, limited to All Systems)"
                                    New-CMDeviceCollection -Name $testCollection -LimitingCollectionName 'All Systems' -ErrorAction Stop | Out-Null
                                    $collection = Get-CMDeviceCollection -Name $testCollection -ErrorAction SilentlyContinue
                                }
                                if (-not $collection) {
                                    Write-Log "Test deployment skipped      : collection '$testCollection' not found (enable 'Create collection if it does not exist' or create it first)" -Level WARN
                                }
                                else {
                                    Write-Log "Deploying to test collection : '$testCollection' (Available, immediate)"
                                    New-CMApplicationDeployment `
                                        -Name $appName `
                                        -CollectionName $testCollection `
                                        -DeployAction Install `
                                        -DeployPurpose Available `
                                        -AvailableDateTime (Get-Date) `
                                        -ErrorAction Stop | Out-Null
                                    Write-Log "Test deployment              : created"
                                }
                            } catch {
                                if ($_.Exception.Message -match 'already (exists|deployed|been deployed)') {
                                    Write-Log "Test deployment              : already exists (treated as success)"
                                } else {
                                    Write-Log "Test deployment failed       : $($_.Exception.Message)" -Level WARN
                                }
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
# Intune Win32 content prep
# ---------------------------------------------------------------------------

function Install-IntuneWinAppUtil {
    <#
    .SYNOPSIS
        Downloads IntuneWinAppUtil.exe into a local tool folder.

    .DESCRIPTION
        Fetches the Microsoft Win32 Content Prep Tool from its official
        GitHub repository and verifies the Authenticode signature is Valid
        and Microsoft-signed before the file lands at its final path. A
        download that fails verification is deleted and the function
        throws. The tool is downloaded per workstation on first use and is
        never redistributed with AppPackager.

    .OUTPUTS
        [string] Full path to the verified IntuneWinAppUtil.exe.
    #>
    param(
        [Parameter(Mandatory)][string]$DestinationFolder,
        [string]$DownloadUrl = 'https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe'
    )

    Initialize-Folder -Path $DestinationFolder
    $target = Join-Path $DestinationFolder 'IntuneWinAppUtil.exe'
    $temp   = Join-Path $DestinationFolder ('IntuneWinAppUtil.' + [Guid]::NewGuid().ToString('N') + '.tmp')

    try {
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $temp

        $sig = Get-AuthenticodeSignature -LiteralPath $temp
        if ($sig.Status -ne 'Valid') {
            throw ("IntuneWinAppUtil.exe download signature status is '{0}', expected 'Valid'; file discarded." -f $sig.Status)
        }
        $subject = if ($sig.SignerCertificate) { [string]$sig.SignerCertificate.Subject } else { '' }
        if ($subject -notmatch 'O=Microsoft Corporation') {
            throw ("IntuneWinAppUtil.exe download signer is '{0}', expected a Microsoft Corporation certificate; file discarded." -f $subject)
        }

        Move-Item -LiteralPath $temp -Destination $target -Force
        Write-Log ("IntuneWinAppUtil.exe verified and installed: {0}" -f $target)
        return $target
    }
    finally {
        if (Test-Path -LiteralPath $temp) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-IntuneWinPackage {
    <#
    .SYNOPSIS
        Produces a .intunewin file from a staged content folder.

    .DESCRIPTION
        Runs the Microsoft Win32 Content Prep Tool against ContentFolder
        with SetupFile as the setup reference. The tool writes
        <setup-basename>.intunewin into OutputFolder; when OutputName is
        given the file is renamed to it. OutputFolder must not be the
        content folder itself: stage-manifest hash verification treats any
        file added to staged or network content as an integrity failure.

    .OUTPUTS
        [pscustomobject] IntuneWinPath, SizeBytes, Sha256, ToolExitCode,
        DurationSec.
    #>
    param(
        [Parameter(Mandatory)][string]$ToolPath,
        [Parameter(Mandatory)][string]$ContentFolder,
        [Parameter(Mandatory)][string]$SetupFile,
        [Parameter(Mandatory)][string]$OutputFolder,
        [string]$OutputName = '',
        [int]$TimeoutSec = 900
    )

    if (-not (Test-Path -LiteralPath $ToolPath)) {
        throw "IntuneWinAppUtil.exe not found: $ToolPath"
    }
    if (-not (Test-Path -LiteralPath $ContentFolder)) {
        throw "Content folder not found: $ContentFolder"
    }
    $setupPath = Join-Path $ContentFolder $SetupFile
    if (-not (Test-Path -LiteralPath $setupPath)) {
        throw "Setup file not found in content folder: $setupPath"
    }
    $contentFull = (Resolve-Path -LiteralPath $ContentFolder).ProviderPath.TrimEnd('\')
    Initialize-Folder -Path $OutputFolder
    $outputFull = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath.TrimEnd('\')
    if ($outputFull -ieq $contentFull) {
        throw "OutputFolder must differ from ContentFolder; a .intunewin inside the content folder fails stage hash verification."
    }

    $expected = Join-Path $outputFull ([IO.Path]::GetFileNameWithoutExtension($SetupFile) + '.intunewin')

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = $ToolPath
    # -q answers the tool's overwrite prompt; without it a pre-existing
    # output file stalls the run waiting for console input.
    $psi.Arguments = ('-c "{0}" -s "{1}" -o "{2}" -q' -f $contentFull, $setupPath, $outputFull)
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $stdOutTask = $proc.StandardOutput.ReadToEndAsync()
        $stdErrTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            try { $proc.Kill() } catch { }
            throw ("IntuneWinAppUtil.exe did not finish within {0}s; process terminated." -f $TimeoutSec)
        }
        $sw.Stop()
        $exitCode = $proc.ExitCode
        $stdOut = [string]$stdOutTask.Result
        $stdErr = [string]$stdErrTask.Result
    }
    finally {
        $proc.Dispose()
    }

    if ($exitCode -ne 0) {
        $tail = @((($stdErr, $stdOut) -join "`n") -split "`r?`n" | Where-Object { $_ } | Select-Object -Last 5) -join ' | '
        throw ("IntuneWinAppUtil.exe exit code {0}: {1}" -f $exitCode, $tail)
    }
    if (-not (Test-Path -LiteralPath $expected)) {
        throw ("IntuneWinAppUtil.exe reported success but the output file is missing: {0}" -f $expected)
    }

    $final = $expected
    if (-not [string]::IsNullOrWhiteSpace($OutputName)) {
        $final = Join-Path $outputFull $OutputName
        Move-Item -LiteralPath $expected -Destination $final -Force
    }

    $item = Get-Item -LiteralPath $final
    return [pscustomobject]@{
        IntuneWinPath = [string]$item.FullName
        SizeBytes     = [long]$item.Length
        Sha256        = (Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash
        ToolExitCode  = [int]$exitCode
        DurationSec   = [int]$sw.Elapsed.TotalSeconds
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
# Ad-hoc drop intake: analyze a dropped installer, stage it through the
# normal manifest path, optionally graduate it into the packager catalog.
# ---------------------------------------------------------------------------

function Get-InstallerAnalysis {
    <#
    .SYNOPSIS
        Runs the vendored installer analysis over one file and returns the
        deployment-relevant digest.

    .DESCRIPTION
        Orchestrates InstallerAnalysisCommon (vendored at
        ..\Lib\InstallerAnalysisCommon): engine identification by binary
        signature, MSI properties when applicable, package metadata,
        silent-switch prediction, and the aggregated deployment fields.
        Confidence is the roadmap's gate: MSI-derived identity is
        authoritative; everything else is predicted and requires operator
        confirmation before Package.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the aggregate analysis for one installer.')]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Installer not found: $Path"
    }
    if (-not (Get-Module -Name InstallerAnalysisCommon)) {
        Import-Module -Name (Join-Path $PSScriptRoot '..\Lib\InstallerAnalysisCommon\InstallerAnalysisCommon.psd1') -Global -DisableNameChecking
    }

    $fileInfo = Get-InstallerFileInfo -Path $Path
    $type     = Get-InstallerType -Path $Path

    $msiProps = $null
    $msiSummary = $null
    if ($type -eq 'MSI') {
        $msiProps   = Get-MsiProperties -MsiPath $Path
        $msiSummary = Get-MsiSummaryInfo -MsiPath $Path
    }
    $pkgMeta  = Get-PackageMetadataFor -Path $Path -InstallerType $type
    $switches = Get-SilentSwitches -InstallerType $type -FilePath $Path -MsiProperties $msiProps
    $fields   = Get-DeploymentFields -FileInfo $fileInfo -MsiProperties $msiProps -Switches $switches `
        -PackageMetadata $pkgMeta -InstallerType $type -MsiSummary $msiSummary

    $productCode = ''
    if ($msiProps -and $msiProps.Contains('ProductCode')) { $productCode = [string]$msiProps['ProductCode'] }

    # Get-SilentSwitches returns full command lines; the wrapper generators
    # want bare arguments, so strip the leading installer filename.
    $installCommand   = if ($switches) { [string]$switches.Install }   else { '' }
    $uninstallCommand = if ($switches) { [string]$switches.Uninstall } else { '' }
    $installArgs = ''
    $uninstallArgs = ''
    if ($type -eq 'MSI') {
        $installArgs   = '/qn /norestart'
        $uninstallArgs = '/qn /norestart'
    }
    else {
        $fileNameEscaped = [regex]::Escape((Split-Path -Leaf $Path))
        if ($installCommand -match ('^"?' + $fileNameEscaped + '"?\s+(?<args>\S.*)$')) {
            $installArgs = $Matches['args'].Trim()
        }
    }

    [pscustomobject]@{
        Path                 = (Resolve-Path -LiteralPath $Path).Path
        FileName             = Split-Path -Leaf $Path
        InstallerType        = $type
        AppName              = [string]$fields.DisplayName
        Publisher            = [string]$fields.Vendor
        SoftwareVersion      = [string]$fields.DisplayVersion
        ProductCode          = $productCode
        InstallArgs          = $installArgs
        UninstallArgs        = $uninstallArgs
        InstallCommand       = $installCommand
        UninstallCommand     = $uninstallCommand
        UninstallRegistryKey = [string]$fields.UninstallRegistryKey
        Architecture         = if ($msiSummary -and $msiSummary.Architecture) { [string]$msiSummary.Architecture } else { [string]$fileInfo.Architecture }
        Confidence           = if ($type -eq 'MSI') { 'Authoritative' } else { 'Predicted' }
        Switches             = $switches
        Fields               = $fields
    }
}

function New-AdHocStage {
    <#
    .SYNOPSIS
        Stages a dropped installer as a versioned content folder with
        wrappers and a stage manifest - the same shape every packager
        produces, so Package consumes it unchanged.

    .DESCRIPTION
        Layout: <DownloadRoot>\<AppFolder>\<Version>\ holding the installer,
        the install/uninstall wrapper set, and stage-manifest.json with
        SHA256 hashes. Detection is RegistryKeyValue on the ARP key: the
        ProductCode key for an MSI (deterministic), or the analysis-
        predicted uninstall key otherwise (operator-confirmed upstream).
    #>
    param(
        [Parameter(Mandatory)]$Analysis,
        [Parameter(Mandatory)][string]$DownloadRoot,
        [string]$AppName,
        [string]$Publisher,
        [string]$SoftwareVersion,
        [string]$InstallArgs,
        [string]$UninstallArgs,
        [string]$UninstallCommand
    )

    if (-not $AppName)         { $AppName         = [string]$Analysis.AppName }
    if (-not $Publisher)       { $Publisher       = [string]$Analysis.Publisher }
    if (-not $SoftwareVersion) { $SoftwareVersion = [string]$Analysis.SoftwareVersion }
    if (-not $InstallArgs)     { $InstallArgs     = [string]$Analysis.InstallArgs }
    if (-not $PSBoundParameters.ContainsKey('UninstallArgs') -or $null -eq $UninstallArgs) { $UninstallArgs = [string]$Analysis.UninstallArgs }
    if ([string]::IsNullOrWhiteSpace($AppName))         { throw 'Ad-hoc stage requires an application name.' }
    if ([string]::IsNullOrWhiteSpace($SoftwareVersion)) { throw 'Ad-hoc stage requires a version.' }
    if ($Analysis.InstallerType -eq 'MSP') {
        throw 'MSI patches (.msp) are not supported for ad-hoc staging; they apply against an installed base MSI, not as a standalone deployment.'
    }
    if ($Analysis.InstallerType -ne 'MSI' -and [string]::IsNullOrWhiteSpace($InstallArgs)) {
        throw 'Ad-hoc stage requires install arguments for a non-MSI installer.'
    }
    if ([string]::IsNullOrWhiteSpace($Publisher))       { $Publisher = 'Unknown' }

    $installerName = [string]$Analysis.FileName
    # Wrapper .ps1 files are written -Encoding ASCII (deployment targets read
    # them with Windows PowerShell defaults); a non-ASCII filename would be
    # silently mangled inside the wrapper and fail only on target devices.
    if ($installerName -match '[^\x00-\x7F]') {
        throw "Installer filename contains non-ASCII characters; rename the file to plain ASCII before staging: $installerName"
    }

    # Folder-name sanitation: identity fields become path segments. A value
    # that sanitizes to nothing must not silently collapse into the parent
    # folder (Join-Path with '' returns the parent).
    $sanitize = { param($s) (($s -replace '[\\/:*?"<>|]', '') -replace '\s+', ' ').Trim() }
    $vendorFolder = & $sanitize $Publisher
    $appFolder    = & $sanitize $AppName
    $versionFolder = & $sanitize $SoftwareVersion
    if ([string]::IsNullOrWhiteSpace($vendorFolder))  { $vendorFolder = 'Unknown' }
    if ([string]::IsNullOrWhiteSpace($appFolder))     { throw "Application name '$AppName' contains no usable path characters." }
    if ([string]::IsNullOrWhiteSpace($versionFolder)) { throw "Version '$SoftwareVersion' contains no usable path characters." }

    $baseRoot = Join-Path (Join-Path $DownloadRoot $vendorFolder) $appFolder
    $stagedPath = Join-Path $baseRoot $versionFolder
    Initialize-Folder -Path $stagedPath
    $stagedInstaller = Join-Path $stagedPath $installerName
    Copy-Item -LiteralPath $Analysis.Path -Destination $stagedInstaller -Force
    Write-Log "Staged dropped installer     : $stagedInstaller"

    $isMsi = ($Analysis.InstallerType -eq 'MSI')
    # New-ExeWrapperContent interpolates InstallArgs/UninstallArgs into an
    # @() literal, so plain space-separated args become a quoted element list.
    $toArgList = {
        param($s)
        (($s -split '\s+') | Where-Object { $_ } | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', '
    }
    if ($isMsi) {
        $wrapperContent = New-MsiWrapperContent -MsiFileName $installerName
    }
    else {
        if ([string]::IsNullOrWhiteSpace($UninstallCommand)) { $UninstallCommand = [string]$Analysis.UninstallCommand }
        # Split "path args" into -FilePath and -ArgumentList halves; the path
        # may be quoted (analysis emits '"uninstall.exe" /S').
        $uninstallExe  = $UninstallCommand
        $uninstallTail = $UninstallArgs
        if ($UninstallCommand -match '^\s*"(?<exe>[^"]+)"\s*(?<tail>.*)$' -or
            $UninstallCommand -match '^\s*(?<exe>\S+)\s+(?<tail>.+)$') {
            $uninstallExe = $Matches['exe']
            if ([string]::IsNullOrWhiteSpace($uninstallTail)) { $uninstallTail = $Matches['tail'].Trim() }
        }
        if ([string]::IsNullOrWhiteSpace($uninstallExe)) {
            # Wrapper still ships; the operator-visible manifest carries the
            # gap so the generated uninstall is an explicit no-op, not a lie.
            $uninstallExe = 'cmd.exe'
            $uninstallTail = '/c exit 0'
        }
        $wrapperContent = New-ExeWrapperContent -InstallerFileName $installerName `
            -InstallArgs (& $toArgList $InstallArgs) `
            -UninstallCommand $uninstallExe `
            -UninstallArgs (& $toArgList $uninstallTail)
    }
    Write-ContentWrappers -OutputPath $stagedPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    $is64 = ([string]$Analysis.Architecture) -notmatch '^(x86|Intel)$'
    if ($isMsi -and $Analysis.ProductCode) {
        $arpRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $Analysis.ProductCode
    }
    else {
        # Analysis predicts a full hive path; the manifest wants the
        # HKLM-relative form. A WOW6432Node prediction means a 32-bit view.
        $arpRelative = ([string]$Analysis.UninstallRegistryKey) -replace '^HK(LM|CU):\\', ''
        if ($arpRelative -match '(?i)\\?WOW6432Node\\') {
            $arpRelative = $arpRelative -replace '(?i)WOW6432Node\\', ''
            $is64 = $false
        }
        if ([string]::IsNullOrWhiteSpace($arpRelative)) {
            $arpRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $AppName
        }
    }

    $manifestPath = Join-Path $stagedPath 'stage-manifest.json'
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $SoftwareVersion
        InstallerFile   = $installerName
        InstallerType   = if ($isMsi) { 'MSI' } else { 'EXE' }
        DetectedEngine  = [string]$Analysis.InstallerType
        InstallArgs     = $InstallArgs
        UninstallArgs   = $UninstallArgs
        UninstallCommand = if ($isMsi) { '' } else { $UninstallCommand }
        ProductCode     = [string]$Analysis.ProductCode
        RunningProcess  = @()
        AdHocSource     = [string]$Analysis.Path
        Detection       = @{
            Type                = 'RegistryKeyValue'
            RegistryKeyRelative = $arpRelative
            ValueName           = 'DisplayVersion'
            DisplayName         = $AppName
            DisplayVersion      = $SoftwareVersion
            Is64Bit             = $is64
        }
    }
    Write-Log "Ad-hoc stage complete        : $stagedPath"

    [pscustomobject]@{
        StagedPath   = $stagedPath
        ManifestPath = $manifestPath
        VendorFolder = $vendorFolder
        AppFolder    = $appFolder
    }
}

function Invoke-AdHocPackage {
    <#
    .SYNOPSIS
        Packages an ad-hoc staged folder: copies content to the network
        version folder and creates the MECM application from the manifest -
        the same path every packager takes.
    #>
    param(
        [Parameter(Mandatory)][string]$StagedPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder,
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$Comment = '',
        [ValidateSet('Nested', 'Flat')][string]$ContentLayout = 'Nested',
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30
    )

    $manifest = Read-StageManifest -Path (Join-Path $StagedPath 'stage-manifest.json')
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    # Version becomes a path segment; strip the same characters the stage
    # sanitized so local and network folder names agree.
    $versionSegment = ((([string]$manifest.SoftwareVersion) -replace '[\\/:*?"<>|]', '') -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($versionSegment)) {
        throw ("Manifest version '{0}' contains no usable path characters." -f $manifest.SoftwareVersion)
    }
    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath `
        -VendorFolder $VendorFolder -AppFolder $AppFolder `
        -Version $versionSegment -Layout $ContentLayout
    Initialize-Folder -Path $networkContentPath

    foreach ($f in (Get-ChildItem -Path $StagedPath -File)) {
        if ($f.Name -eq 'stage-manifest.json') { continue }
        Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $networkContentPath $f.Name) -Force
    }
    Write-Log "Ad-hoc content on network    : $networkContentPath"

    return New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
}

function New-PackagerFromDrop {
    <#
    .SYNOPSIS
        Graduates an analyzed drop into the packager catalog: writes
        Packagers\package-<slug>.ps1 from the matching template with the
        analysis-filled identity values.

    .DESCRIPTION
        Identity, folder segments, and installer filename are baked; the
        download-source resolution stays a TODO by nature - a dropped file
        carries no URL. The generated file satisfies the grid's discovery
        contract immediately (header metadata, GetLatestVersionOnly, phase
        switches) and refuses to overwrite an existing packager.
    #>
    param(
        [Parameter(Mandatory)]$Analysis,
        [Parameter(Mandatory)][string]$PackagersRoot,
        [string]$AppName,
        [string]$Publisher,
        [string]$SoftwareVersion
    )

    if (-not $AppName)         { $AppName         = [string]$Analysis.AppName }
    if (-not $Publisher)       { $Publisher       = [string]$Analysis.Publisher }
    if (-not $SoftwareVersion) { $SoftwareVersion = [string]$Analysis.SoftwareVersion }
    if ([string]::IsNullOrWhiteSpace($AppName)) { throw 'A packager needs an application name.' }

    $isMsi = ($Analysis.InstallerType -eq 'MSI')
    $templateName = if ($isMsi) { 'package-template-msi.ps1' } else { 'package-template-exe.ps1' }
    $templatePath = Join-Path (Split-Path $PackagersRoot -Parent) (Join-Path 'Samples' $templateName)
    if (-not (Test-Path -LiteralPath $templatePath)) { throw "Template not found: $templatePath" }

    $slug = (($AppName.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-'))
    $target = Join-Path $PackagersRoot ("package-{0}.ps1" -f $slug)
    if (Test-Path -LiteralPath $target) { throw "Packager already exists: $target" }

    $sanitize = { param($s) (($s -replace '[\\/:*?"<>|]', '') -replace '\s+', ' ').Trim() }
    $vendorFolder = & $sanitize $Publisher
    $appFolder    = & $sanitize $AppName

    # Identity values originate in the dropped file's own version resource /
    # MSI tables: untrusted text. Substitutions into live double-quoted
    # literals must neutralize `, $, and " or the resource string executes
    # as code the next time the generated packager runs; substitutions into
    # the doc comment must not carry a comment terminator.
    $escapeDq = { param($s) (([string]$s -replace '`', '``') -replace '\$', '`$') -replace '"', '`"' }
    $escapeComment = { param($s) [string]$s -replace '#>', '# >' }
    $appNameDq   = & $escapeDq $AppName
    $publisherDq = & $escapeDq $Publisher
    $appNameCmt   = & $escapeComment $AppName
    $publisherCmt = & $escapeComment $Publisher

    # Literal String.Replace throughout: values may contain characters regex
    # replacement strings would reinterpret.
    $c = Get-Content -LiteralPath $templatePath -Raw
    $c = $c.Replace('Vendor: TODO', "Vendor: $publisherCmt")
    $c = $c.Replace('App: TODO', "App: $appNameCmt")
    $c = $c.Replace('CMName: TODO', "CMName: $appNameCmt")
    $c = $c.Replace('$VendorFolder = "TODO"', ('$VendorFolder = "{0}"' -f (& $escapeDq $vendorFolder)))
    $c = $c.Replace('$AppFolder    = "TODO"', ('$AppFolder    = "{0}"' -f (& $escapeDq $appFolder)))
    $c = $c.Replace('Packages TODO (x64)', "Packages $appNameCmt")
    $c = $c.Replace('"TODO - STAGE phase"', ('"{0} - STAGE phase"' -f $appNameDq))
    $c = $c.Replace('"TODO - PACKAGE phase"', ('"{0} - PACKAGE phase"' -f $appNameDq))
    $c = $c.Replace('"TODO Auto-Packager starting"', ('"{0} Auto-Packager starting"' -f $appNameDq))
    if ($isMsi) {
        $c = $c.Replace('$MsiFileName      = "TODO-installer.msi"', ('$MsiFileName      = "{0}"' -f (& $escapeDq $Analysis.FileName)))
    }
    else {
        $c = $c.Replace('AppName          = "TODO"', ('AppName          = "{0}"' -f $appNameDq))
        $c = $c.Replace('Publisher        = "TODO"', ('Publisher        = "{0}"' -f $publisherDq))
        if (-not [string]::IsNullOrWhiteSpace([string]$Analysis.InstallArgs)) {
            $c = $c.Replace('$installArgs   = "/S"', ('$installArgs   = "{0}"' -f (& $escapeDq $Analysis.InstallArgs)))
        }
    }

    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($c, [ref]$null, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw ("Generated packager does not parse ({0}); refusing to write it. First error: {1}" -f $parseErrors.Count, $parseErrors[0].Message)
    }

    Set-Content -LiteralPath $target -Value $c -Encoding UTF8
    Write-Log "Packager written             : $target (download source is a TODO; identity and folders are filled)"
    return $target
}


# ---------------------------------------------------------------------------
# Module export (belt-and-suspenders with .psd1 FunctionsToExport)
# ---------------------------------------------------------------------------

Export-ModuleMember -Function *

function Get-RequestedPackagerVariants {
    <#
    .SYNOPSIS
        Parses the APP_PACKAGER_VARIANTS environment JSON the GUI sets for
        a SupportsVariants packager.

    .DESCRIPTION
        Returns $null when no variant split is requested. Malformed JSON
        throws instead of silently packaging without the operator's split.

    .OUTPUTS
        [pscustomobject] Split ('Architecture' | 'Language' | 'Network')
        and Languages (culture codes; only meaningful for a Language
        split), or $null.
    #>
    $envJson = $env:APP_PACKAGER_VARIANTS
    if ([string]::IsNullOrWhiteSpace($envJson)) { return $null }
    try {
        $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "APP_PACKAGER_VARIANTS is not valid JSON: $($_.Exception.Message)"
    }
    $split = [string]$parsed.Split
    if ($split -notin @('Architecture', 'Language', 'Network')) {
        throw "APP_PACKAGER_VARIANTS Split must be Architecture, Language, or Network (got '$split')."
    }
    $languages = @()
    if ($parsed.PSObject.Properties['Languages'] -and $parsed.Languages) {
        $languages = @($parsed.Languages | ForEach-Object { [string]$_ } | Where-Object { $_ })
    }
    if ($split -eq 'Language' -and $languages.Count -eq 0) {
        throw 'APP_PACKAGER_VARIANTS requests a Language split with no Languages.'
    }
    return [pscustomobject]@{ Split = $split; Languages = $languages }
}

Export-ModuleMember -Function Get-RequestedPackagerVariants

function Get-RequestedCommandOverrides {
    <#
    .SYNOPSIS
        Parses the APP_PACKAGER_COMMANDS environment JSON the GUI sets
        when the operator overrides an app's install or uninstall command.

    .DESCRIPTION
        Returns $null when no override is configured. Malformed JSON or an
        override with neither command throws instead of silently packaging
        with the wrong commands.

    .OUTPUTS
        [pscustomobject] Install and Uninstall (either may be empty), or
        $null.
    #>
    $envJson = $env:APP_PACKAGER_COMMANDS
    if ([string]::IsNullOrWhiteSpace($envJson)) { return $null }
    try {
        $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "APP_PACKAGER_COMMANDS is not valid JSON: $($_.Exception.Message)"
    }
    $install = [string]$parsed.Install
    $uninstall = [string]$parsed.Uninstall
    if ([string]::IsNullOrWhiteSpace($install) -and [string]::IsNullOrWhiteSpace($uninstall)) {
        throw 'APP_PACKAGER_COMMANDS carries neither an Install nor an Uninstall command.'
    }
    return [pscustomobject]@{ Install = $install.Trim(); Uninstall = $uninstall.Trim() }
}

Export-ModuleMember -Function Get-RequestedCommandOverrides

# ---------------------------------------------------------------------------
# Intune Win32 publishing (Graph)
# ---------------------------------------------------------------------------

function Get-IntuneWinEncryptionInfo {
    <#
    .SYNOPSIS
        Reads the metadata and encryption info from a .intunewin file.

    .DESCRIPTION
        A .intunewin file is a zip whose IntuneWinPackage/Metadata/Detection.xml
        carries the fields the Graph commit action needs (fileEncryptionInfo)
        plus the setup file name and unencrypted size. The encrypted payload
        itself is IntuneWinPackage/Contents/IntunePackage.intunewin.

    .OUTPUTS
        [pscustomobject] Name, FileName, SetupFile, UnencryptedContentSize,
        EncryptionKey, MacKey, InitializationVector, Mac, ProfileIdentifier,
        FileDigest, FileDigestAlgorithm.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "IntuneWin file not found: $Path" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'IntuneWinPackage/Metadata/Detection.xml' } | Select-Object -First 1
        if (-not $entry) { throw "Detection.xml not found inside $Path; not a valid .intunewin package." }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { [xml]$doc = $reader.ReadToEnd() } finally { $reader.Dispose() }

        $info = $doc.ApplicationInfo
        $enc = $info.EncryptionInfo
        if (-not $enc) { throw "Detection.xml carries no EncryptionInfo; the package cannot be committed." }

        return [pscustomobject]@{
            Name                   = [string]$info.Name
            FileName               = [string]$info.FileName
            SetupFile              = [string]$info.SetupFile
            UnencryptedContentSize = [long]$info.UnencryptedContentSize
            EncryptionKey          = [string]$enc.EncryptionKey
            MacKey                 = [string]$enc.MacKey
            InitializationVector   = [string]$enc.InitializationVector
            Mac                    = [string]$enc.Mac
            ProfileIdentifier      = [string]$enc.ProfileIdentifier
            FileDigest             = [string]$enc.FileDigest
            FileDigestAlgorithm    = [string]$enc.FileDigestAlgorithm
        }
    }
    finally { $zip.Dispose() }
}

function Export-IntuneWinPayload {
    <#
    .SYNOPSIS
        Extracts the encrypted payload from a .intunewin file.

    .DESCRIPTION
        The bytes uploaded to Azure Storage are the inner encrypted
        IntunePackage.intunewin, not the outer zip.

    .OUTPUTS
        [pscustomobject] Path and Size of the extracted payload.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'IntuneWinPackage/Contents/IntunePackage.intunewin' } | Select-Object -First 1
        if (-not $entry) { throw "Encrypted payload not found inside $Path; not a valid .intunewin package." }
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $Destination, $true)
        return [pscustomobject]@{ Path = $Destination; Size = (Get-Item -LiteralPath $Destination).Length }
    }
    finally { $zip.Dispose() }
}

function Get-MsGraphToken {
    <#
    .SYNOPSIS
        Acquires an app-only Graph token via the client-credentials flow.
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $resp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
    } -ErrorAction Stop
    if (-not $resp.access_token) { throw 'Token endpoint returned no access_token.' }
    return [string]$resp.access_token
}

function Invoke-GraphJson {
    <#
    .SYNOPSIS
        One Graph REST call. Central so retries and tests hang off one seam.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Token,
        $Body = $null
    )

    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) { $params['Body'] = ($Body | ConvertTo-Json -Depth 12) }
    return Invoke-RestMethod @params
}

function Invoke-AzureBlobUpload {
    <#
    .SYNOPSIS
        Uploads a file to an Azure block-blob SAS URI in chunks.

    .DESCRIPTION
        Standard block-blob protocol: PUT each chunk to
        <uri>&comp=block&blockid=<base64 id>, then PUT the ordered block
        list to <uri>&comp=blocklist. 6 MiB chunks keep each request under
        proxy limits without ballooning the block count.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$FilePath,
        [int]$ChunkSizeMB = 6
    )

    $chunkSize = $ChunkSizeMB * 1MB
    $stream = [System.IO.File]::OpenRead($FilePath)
    $blockIds = New-Object System.Collections.Generic.List[string]
    try {
        $buffer = New-Object byte[] $chunkSize
        $index = 0
        while (($read = $stream.Read($buffer, 0, $chunkSize)) -gt 0) {
            $blockId = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('block-{0:D6}' -f $index)))
            $blockIds.Add($blockId)
            $chunk = if ($read -eq $chunkSize) { $buffer } else { $buffer[0..($read - 1)] }
            $blockUri = '{0}&comp=block&blockid={1}' -f $Uri, [uri]::EscapeDataString($blockId)
            Invoke-RestMethod -Method Put -Uri $blockUri -Body $chunk -Headers @{ 'x-ms-blob-type' = 'BlockBlob' } -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
            $index++
        }
        $blockListXml = '<?xml version="1.0" encoding="utf-8"?><BlockList>' + (($blockIds | ForEach-Object { "<Latest>$_</Latest>" }) -join '') + '</BlockList>'
        Invoke-RestMethod -Method Put -Uri ('{0}&comp=blocklist' -f $Uri) -Body $blockListXml -ContentType 'text/plain; charset=iso-8859-1' -ErrorAction Stop | Out-Null
        Write-Log ("Uploaded payload             : {0} block(s)" -f $blockIds.Count)
    }
    finally { $stream.Dispose() }
}

function ConvertTo-IntuneWin32Rules {
    <#
    .SYNOPSIS
        Maps a stage manifest Detection onto Graph win32LobApp rules.

    .DESCRIPTION
        Field names per the v1.0 win32LobApp*Rule resources. RegistryKeyValue
        compares DisplayVersion-style strings with a version operation when
        the expected value parses as a version, string equality otherwise.
        Compound detections map clause-per-rule (Graph ANDs all detection
        rules; OR-connected compounds are refused rather than silently
        narrowed).
    #>
    param([Parameter(Mandatory)][pscustomobject]$Manifest)

    $det = $Manifest.Detection
    $type = if ($det.Type) { [string]$det.Type } else { 'RegistryKeyValue' }

    $mapOne = {
        param($d, $dType)
        switch ($dType) {
            'RegistryKeyValue' {
                $expected = if ($d.ExpectedValue) { [string]$d.ExpectedValue } else { [string]$d.DisplayVersion }
                $valName = if ($d.ValueName) { [string]$d.ValueName } else { 'DisplayVersion' }
                $ver = $null
                $opType = if ([version]::TryParse($expected, [ref]$ver)) { 'version' } else { 'string' }
                $op = if ($d.Operator -eq 'GreaterEquals') { 'greaterThanOrEqual' } else { 'equal' }
                @{
                    '@odata.type'        = '#microsoft.graph.win32LobAppRegistryRule'
                    ruleType             = 'detection'
                    check32BitOn64System = (-not [bool]$d.Is64Bit)
                    keyPath              = ('HKEY_LOCAL_MACHINE\' + ([string]$d.RegistryKeyRelative))
                    valueName            = $valName
                    operationType        = $opType
                    operator             = $op
                    comparisonValue      = $expected
                }
            }
            'RegistryKey' {
                @{
                    '@odata.type'        = '#microsoft.graph.win32LobAppRegistryRule'
                    ruleType             = 'detection'
                    check32BitOn64System = (-not [bool]$d.Is64Bit)
                    keyPath              = ('HKEY_LOCAL_MACHINE\' + ([string]$d.RegistryKeyRelative))
                    valueName            = $null
                    operationType        = 'exists'
                    operator             = 'notConfigured'
                    comparisonValue      = $null
                }
            }
            'File' {
                $propType = [string]$d.PropertyType
                if ($propType -eq 'Version') {
                    $op = if ($d.Operator -eq 'GreaterEquals') { 'greaterThanOrEqual' } else { 'equal' }
                    @{
                        '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
                        ruleType             = 'detection'
                        path                 = [string]$d.FilePath
                        fileOrFolderName     = [string]$d.FileName
                        check32BitOn64System = (-not [bool]$d.Is64Bit)
                        operationType        = 'version'
                        operator             = $op
                        comparisonValue      = [string]$d.ExpectedValue
                    }
                }
                else {
                    @{
                        '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
                        ruleType             = 'detection'
                        path                 = [string]$d.FilePath
                        fileOrFolderName     = [string]$d.FileName
                        check32BitOn64System = (-not [bool]$d.Is64Bit)
                        operationType        = 'exists'
                        operator             = 'notConfigured'
                        comparisonValue      = $null
                    }
                }
            }
            'Script' {
                @{
                    '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptRule'
                    ruleType              = 'detection'
                    enforceSignatureCheck = $false
                    runAs32Bit            = $false
                    scriptContent         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$d.ScriptText))
                    operationType         = 'notConfigured'
                    operator              = 'notConfigured'
                }
            }
            default { throw "Detection type '$dType' has no Intune rule mapping." }
        }
    }

    if ($type -eq 'Compound') {
        if ([string]$det.Connector -eq 'Or' -or ($det.PSObject.Properties['GroupSizes'] -and $det.GroupSizes)) {
            throw 'OR-connected compound detections cannot map to Intune rules (Graph ANDs all detection rules); use a Script detection for this app.'
        }
        return @($det.Clauses | ForEach-Object { & $mapOne $_ ([string]$_.Type) })
    }
    return @(& $mapOne $det $type)
}

function Publish-IntuneWin32App {
    <#
    .SYNOPSIS
        Publishes a packaged .intunewin as an Intune Win32 app via Graph.

    .DESCRIPTION
        The complete v1.0 flow: create the win32LobApp (metadata, commands,
        detection rules from the stage manifest), create a content version
        and file entry, wait for the Azure Storage URI, upload the encrypted
        payload in blocks, commit with the package's fileEncryptionInfo,
        wait for commit, and patch committedContentVersion. Assignment is
        left to the operator in the Intune console.

        Requires an Entra app registration with application permission
        DeviceManagementApps.ReadWrite.All (admin-consented).

    .OUTPUTS
        [string] The created mobile app id.
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$IntuneWinPath,
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [AllowEmptyString()][string]$Description = '',
        [string]$MinimumSupportedWindowsRelease = 'Windows10_21H2',
        [ValidateSet('x86', 'x64')][string]$Architecture = 'x64',
        [string]$GraphBase = 'https://graph.microsoft.com/v1.0',
        [int]$PollTimeoutSec = 600
    )

    $meta = Get-IntuneWinEncryptionInfo -Path $IntuneWinPath
    $rules = @(ConvertTo-IntuneWin32Rules -Manifest $Manifest)

    $installCommand = if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.InstallCommandLine)) { [string]$Manifest.InstallCommandLine } else { 'install.bat' }
    $uninstallCommand = if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.UninstallCommandLine)) { [string]$Manifest.UninstallCommandLine } else { 'uninstall.bat' }

    $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

    # Repeat publishes update the existing app (new content version on the
    # same app identity) instead of stacking duplicates in the console.
    $filterName = ([string]$Manifest.AppName) -replace "'", "''"
    $existing = Invoke-GraphJson -Method GET -Uri ("$GraphBase/deviceAppManagement/mobileApps?`$filter=displayName eq '$filterName'") -Token $token
    $existingApp = @($existing.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.win32LobApp' }) | Select-Object -First 1

    $appBody = @{
        '@odata.type'                  = '#microsoft.graph.win32LobApp'
        displayName                    = [string]$Manifest.AppName
        description                    = $(if ($Description) { $Description } else { [string]$Manifest.AppName })
        publisher                      = [string]$Manifest.Publisher
        fileName                       = (Split-Path -Leaf $IntuneWinPath)
        setupFilePath                  = $meta.SetupFile
        installCommandLine             = $installCommand
        uninstallCommandLine           = $uninstallCommand
        applicableArchitectures        = $Architecture
        minimumSupportedWindowsRelease = $MinimumSupportedWindowsRelease
        installExperience              = @{
            runAsAccount          = 'system'
            deviceRestartBehavior = 'basedOnReturnCode'
        }
        rules                          = $rules
        returnCodes                    = @(
            @{ returnCode = 0;    type = 'success' }
            @{ returnCode = 1707; type = 'success' }
            @{ returnCode = 3010; type = 'softReboot' }
            @{ returnCode = 1641; type = 'hardReboot' }
            @{ returnCode = 1618; type = 'retry' }
        )
    }
    if ($existingApp) {
        $appId = [string]$existingApp.id
        Write-Log ("Updating Intune Win32 app    : {0} (app id {1})" -f $Manifest.AppName, $appId)
        Invoke-GraphJson -Method PATCH -Uri "$GraphBase/deviceAppManagement/mobileApps/$appId" -Token $token -Body $appBody | Out-Null
    }
    else {
        Write-Log ("Creating Intune Win32 app    : {0}" -f $Manifest.AppName)
        $app = Invoke-GraphJson -Method POST -Uri "$GraphBase/deviceAppManagement/mobileApps" -Token $token -Body $appBody
        $appId = [string]$app.id
        Write-Log ("Intune app id                : {0}" -f $appId)
    }

    $lobBase = "$GraphBase/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp"
    $content = Invoke-GraphJson -Method POST -Uri "$lobBase/contentVersions" -Token $token -Body @{}
    $contentId = [string]$content.id

    $payload = Export-IntuneWinPayload -Path $IntuneWinPath -Destination ([System.IO.Path]::GetTempFileName())
    try {
        $fileBody = @{
            '@odata.type' = '#microsoft.graph.mobileAppContentFile'
            name          = $meta.FileName
            size          = [long]$meta.UnencryptedContentSize
            sizeEncrypted = [long]$payload.Size
            manifest      = $null
            isDependency  = $false
        }
        $file = Invoke-GraphJson -Method POST -Uri "$lobBase/contentVersions/$contentId/files" -Token $token -Body $fileBody
        $fileId = [string]$file.id
        $fileUri = "$lobBase/contentVersions/$contentId/files/$fileId"

        $deadline = (Get-Date).AddSeconds($PollTimeoutSec)
        do {
            Start-Sleep -Seconds 3
            $file = Invoke-GraphJson -Method GET -Uri $fileUri -Token $token
            if ([string]$file.uploadState -match 'Failed|TimedOut') { throw "Azure storage URI request failed: $($file.uploadState)" }
        } until ($file.azureStorageUri -or (Get-Date) -gt $deadline)
        if (-not $file.azureStorageUri) { throw 'Timed out waiting for the Azure storage upload URI.' }

        Write-Log "Uploading encrypted payload  : $([math]::Round($payload.Size / 1MB, 1)) MB"
        Invoke-AzureBlobUpload -Uri ([string]$file.azureStorageUri) -FilePath $payload.Path

        $commitBody = @{
            fileEncryptionInfo = @{
                encryptionKey        = $meta.EncryptionKey
                macKey               = $meta.MacKey
                initializationVector = $meta.InitializationVector
                mac                  = $meta.Mac
                profileIdentifier    = $meta.ProfileIdentifier
                fileDigest           = $meta.FileDigest
                fileDigestAlgorithm  = $meta.FileDigestAlgorithm
            }
        }
        Invoke-GraphJson -Method POST -Uri "$fileUri/commit" -Token $token -Body $commitBody | Out-Null

        $deadline = (Get-Date).AddSeconds($PollTimeoutSec)
        do {
            Start-Sleep -Seconds 5
            $file = Invoke-GraphJson -Method GET -Uri $fileUri -Token $token
            if ([string]$file.uploadState -match 'commitFileFailed|commitFileTimedOut') { throw "Content commit failed: $($file.uploadState)" }
        } until ([string]$file.uploadState -eq 'commitFileSuccess' -or (Get-Date) -gt $deadline)
        if ([string]$file.uploadState -ne 'commitFileSuccess') { throw "Timed out waiting for content commit (last state: $($file.uploadState))." }

        Invoke-GraphJson -Method PATCH -Uri "$GraphBase/deviceAppManagement/mobileApps/$appId" -Token $token -Body @{
            '@odata.type'           = '#microsoft.graph.win32LobApp'
            committedContentVersion = $contentId
        } | Out-Null

        Write-Log ("Published to Intune          : {0} (app id {1})" -f $Manifest.AppName, $appId)
        return $appId
    }
    finally {
        Remove-Item -LiteralPath $payload.Path -Force -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function Get-IntuneWinEncryptionInfo, Export-IntuneWinPayload, Get-MsGraphToken, Invoke-GraphJson, Invoke-AzureBlobUpload, ConvertTo-IntuneWin32Rules, Publish-IntuneWin32App
