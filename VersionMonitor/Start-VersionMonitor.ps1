<#
.SYNOPSIS
    Vendor Version Monitor - compares MECM-packaged versions against vendor releases.

.DESCRIPTION
    Discovers packager scripts from the sibling Packagers folder, reads CPE and
    URL metadata from their headers, queries MECM for deployed versions, checks
    each vendor for the latest available version, flags stale applications,
    optionally queries the NIST NVD for known CVEs, and produces a self-contained
    HTML report.

    Designed for headless/scheduled execution. No GUI, no MECM changes.

.PARAMETER ConfigPath
    Path to monitor-config.json. Defaults to $PSScriptRoot\monitor-config.json.

.PARAMETER SkipNVD
    Skip NVD CVE lookups. Useful for faster runs or when NVD is unreachable.

.PARAMETER SkipMECM
    Skip MECM queries. Vendor version checks still run. Useful for testing
    without a ConfigMgr connection.

.PARAMETER SimulateStale
    Load simulated MECM versions from simulate-overrides.json. Forces apps in
    the overrides file to appear stale so you can test CVE lookups and report
    rendering without a real MECM environment or genuinely stale apps.

.PARAMETER OverridesPath
    Path to simulate-overrides.json. Defaults to $PSScriptRoot\simulate-overrides.json.

.PARAMETER OutputPath
    Override the report output path from config.

.EXAMPLE
    .\Start-VersionMonitor.ps1
    Full run: MECM + vendor checks + NVD CVE lookups.

.EXAMPLE
    .\Start-VersionMonitor.ps1 -SkipMECM -SkipNVD
    Vendor version checks only. No MECM or NVD dependency.

.EXAMPLE
    .\Start-VersionMonitor.ps1 -SimulateStale
    Simulate stale versions from overrides file, run NVD lookups against them.

.NOTES
    ScriptName : Start-VersionMonitor.ps1
    Version    : 2.0.0
    Updated    : 2026-03-03
#>

param(
    [string]$ConfigPath    = (Join-Path $PSScriptRoot "monitor-config.json"),
    [switch]$SkipNVD,
    [switch]$SkipMECM,
    [switch]$SimulateStale,
    [string]$OverridesPath = (Join-Path $PSScriptRoot "simulate-overrides.json"),
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ---------------------------------------------------------------------------
# Resolve paths relative to this script's location inside applicationpackager
# ---------------------------------------------------------------------------

$vmRoot       = $PSScriptRoot
$appRoot      = Split-Path $vmRoot -Parent          # applicationpackager\1.0
$packagersDir = Join-Path $appRoot "Packagers"

# ---------------------------------------------------------------------------
# Load modules
# ---------------------------------------------------------------------------

$moduleRoot = Join-Path $vmRoot "Module"
Import-Module (Join-Path $moduleRoot "VersionMonitorCommon.psd1") -Force -DisableNameChecking

# Import AppPackagerCommon for Write-Log, Initialize-Logging
$appPackagerModule = Join-Path $packagersDir "AppPackagerCommon.psd1"
if (Test-Path -LiteralPath $appPackagerModule) {
    Import-Module $appPackagerModule -Force -DisableNameChecking
}
else {
    throw "AppPackagerCommon module not found at: $appPackagerModule"
}

# ---------------------------------------------------------------------------
# Load config
# ---------------------------------------------------------------------------

$config = Read-MonitorConfig -ConfigPath $ConfigPath

# Default log/report folders to subfolders of VersionMonitor if not set in config
$logFolder = if ($config.Logging.LogFolder) { $config.Logging.LogFolder } else { Join-Path $vmRoot "Logs" }
$reportFolder = if ($config.Report.OutputFolder) { $config.Report.OutputFolder } else { Join-Path $vmRoot "Reports" }

# Initialize logging
if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
$logPath = Join-Path $logFolder ("VersionMonitor-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $logPath

Write-Log "=== Vendor Version Monitor started ==="
Write-Log ("Packagers root: {0}" -f $packagersDir)
if ($SimulateStale) { Write-Log "SIMULATE MODE: Using overrides from $OverridesPath" -Level WARN }
if ($SkipMECM)      { Write-Log "Skipping MECM queries (-SkipMECM)" }
if ($SkipNVD)       { Write-Log "Skipping NVD lookups (-SkipNVD)" }

# ---------------------------------------------------------------------------
# Load simulation overrides
# ---------------------------------------------------------------------------

$simOverrides = @{}
if ($SimulateStale -and (Test-Path -LiteralPath $OverridesPath)) {
    try {
        $ovJson = Get-Content -LiteralPath $OverridesPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $ovJson.Overrides.PSObject.Properties) {
            $simOverrides[$prop.Name] = $prop.Value
        }
        Write-Log ("Loaded {0} simulation overrides" -f $simOverrides.Count)
    }
    catch {
        Write-Log ("Failed to load overrides: {0}" -f $_.Exception.Message) -Level WARN
    }
}

# ---------------------------------------------------------------------------
# Discover packagers (metadata + CPE/URLs come from script headers)
# ---------------------------------------------------------------------------

Write-Log ("Discovering packagers in: {0}" -f $packagersDir)
$packagers = Get-PackagerScripts -PackagersRoot $packagersDir
Write-Log ("Found {0} packager scripts" -f $packagers.Count)

# Build scan list directly from packager metadata (no catalog needed)
$appsToScan = @()
foreach ($pkg in $packagers) {
    $appsToScan += [pscustomobject]@{
        Script          = $pkg.Script
        FullPath        = $pkg.FullPath
        Application     = $pkg.Application
        Publisher       = $pkg.Vendor
        CMName          = $pkg.CMName
        CPE             = $pkg.CPE
        ReleaseNotesUrl = $pkg.ReleaseNotesUrl
        DownloadPageUrl = $pkg.DownloadPageUrl
        MecmVersion     = $null
        VendorVersion   = $null
        Status          = 'Pending'
        CVECount        = $null
        MaxCVSS         = $null
        MaxSeverity     = $null
        MaxCVEUrl       = $null
        CVEError        = $null
        ErrorMessage    = $null
    }
}
Write-Log ("{0} applications to scan" -f $appsToScan.Count)

# ---------------------------------------------------------------------------
# Query MECM
# ---------------------------------------------------------------------------

$mecmResults = @{}
if (-not $SkipMECM -and -not $SimulateStale) {
    Write-Log "Querying MECM for application versions..."
    try {
        $cmNames = @($appsToScan | ForEach-Object { $_.CMName } | Where-Object { $_ })
        $mecmResults = Get-MecmApplicationVersions -SiteCode $config.MECM.SiteCode -CMNames $cmNames
        $foundCount = @($mecmResults.Values | Where-Object { $_.Found }).Count
        Write-Log ("MECM: {0} of {1} applications found" -f $foundCount, $cmNames.Count)
    }
    catch {
        Write-Log ("MECM connection failed: {0}" -f $_.Exception.Message) -Level ERROR
        Write-Log "Continuing with vendor checks only"
    }
}

# Apply MECM versions (or simulation overrides)
foreach ($app in $appsToScan) {
    if ($SimulateStale -and $simOverrides.ContainsKey($app.Script)) {
        $app.MecmVersion = $simOverrides[$app.Script]
    }
    elseif ($mecmResults.ContainsKey($app.CMName) -and $mecmResults[$app.CMName].Found) {
        $app.MecmVersion = $mecmResults[$app.CMName].SoftwareVersion
    }
    elseif (-not $SkipMECM -and -not $SimulateStale) {
        $app.Status = 'Not in MECM'
    }
}

# ---------------------------------------------------------------------------
# Check vendor versions
# ---------------------------------------------------------------------------

Write-Log "Checking vendor versions..."
$downloadRoot = $config.DownloadRoot
if (-not (Test-Path -LiteralPath $downloadRoot)) {
    New-Item -ItemType Directory -Path $downloadRoot -Force | Out-Null
}

foreach ($app in $appsToScan) {
    Write-Log ("  Checking: {0}" -f $app.Application)
    try {
        $vendorVer = Invoke-VendorVersionCheck `
            -PackagerPath $app.FullPath `
            -SiteCode $config.MECM.SiteCode `
            -DownloadRoot $downloadRoot `
            -TimeoutSeconds 120

        $app.VendorVersion = $vendorVer
        Write-Log ("    Vendor version: {0}" -f $vendorVer)
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg.Length -gt 200) { $errMsg = $errMsg.Substring(0, 200) + '...' }
        Write-Log ("    FAILED: {0}" -f $errMsg) -Level WARN
        $app.ErrorMessage = $errMsg
        if ($app.Status -eq 'Pending') { $app.Status = 'Error' }
    }
}

# ---------------------------------------------------------------------------
# Compare versions
# ---------------------------------------------------------------------------

foreach ($app in $appsToScan) {
    if ($app.Status -ne 'Pending') { continue }

    if ($SimulateStale -and -not $simOverrides.ContainsKey($app.Script)) {
        # In simulate mode, apps without overrides use vendor version as MECM version (appear current)
        $app.MecmVersion = $app.VendorVersion
    }

    $cmp = Compare-Versions -MecmVersion $app.MecmVersion -VendorVersion $app.VendorVersion
    $app.Status = $cmp.Status
}

$staleApps = @($appsToScan | Where-Object { $_.Status -eq 'Stale' })
$currentApps = @($appsToScan | Where-Object { $_.Status -eq 'Current' })
Write-Log ("Version comparison complete: {0} current, {1} stale, {2} other" -f $currentApps.Count, $staleApps.Count, ($appsToScan.Count - $currentApps.Count - $staleApps.Count))

# ---------------------------------------------------------------------------
# NVD CVE lookup (stale apps only)
# ---------------------------------------------------------------------------

if (-not $SkipNVD -and $staleApps.Count -gt 0) {
    $staleWithCpe = @($staleApps | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CPE) })
    Write-Log ("Querying NVD for {0} stale applications ({1} have CPE)" -f $staleApps.Count, $staleWithCpe.Count)

    $rateLimit = if ($config.NVD.ApiKey) { 50 } else { $config.NVD.RateLimitPerWindow }
    $cachePath = Join-Path $vmRoot "nvd-cache.json"

    $nvdResults = Invoke-NvdBatchQuery `
        -StaleApps $staleApps `
        -ApiKey $config.NVD.ApiKey `
        -RateLimit $rateLimit `
        -WindowSeconds $config.NVD.WindowSeconds `
        -CachePath $cachePath `
        -CacheTtlMinutes $config.NVD.CacheTtlMinutes

    foreach ($app in $staleApps) {
        if ($nvdResults.ContainsKey($app.Script)) {
            $nvd = $nvdResults[$app.Script]
            $app.CVECount    = $nvd.CVECount
            $app.MaxCVSS     = $nvd.MaxCVSS
            $app.MaxSeverity = $nvd.MaxSeverity
            $app.CVEError    = $nvd.Error
            $maxCve = $nvd.CVEs | Sort-Object Score -Descending | Select-Object -First 1
            if ($maxCve) { $app.MaxCVEUrl = $maxCve.Url }
        }
    }
}
elseif ($SkipNVD) {
    Write-Log "NVD lookups skipped"
}
else {
    Write-Log "No stale applications - NVD lookups not needed"
}

# ---------------------------------------------------------------------------
# Generate HTML report
# ---------------------------------------------------------------------------

if (-not $OutputPath) {
    if (-not (Test-Path -LiteralPath $reportFolder)) {
        New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
    }
    $OutputPath = Join-Path $reportFolder ($config.Report.FileNamePattern -f (Get-Date -Format 'yyyy-MM-dd-HHmmss'))
}

Write-Log "Generating HTML report..."
$reportPath = Export-VersionMonitorHtml -Results $appsToScan -OutputPath $OutputPath
Write-Log ("Report written to: {0}" -f $reportPath)

# ---------------------------------------------------------------------------
# Notifications
# ---------------------------------------------------------------------------

if ($config.Notifications) {
    Send-ReportNotification `
        -ReportPath $reportPath `
        -NotificationConfig $config.Notifications `
        -StaleCount $staleApps.Count `
        -TotalCount $appsToScan.Count
}

# ---------------------------------------------------------------------------
# Cleanup old reports and logs
# ---------------------------------------------------------------------------

if ($config.Report.KeepReportDays -gt 0) {
    $cutoff = (Get-Date).AddDays(-$config.Report.KeepReportDays)
    Get-ChildItem -LiteralPath $reportFolder -Filter '*.html' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}
if ($config.Logging.KeepLogDays -gt 0) {
    $cutoff = (Get-Date).AddDays(-$config.Logging.KeepLogDays)
    Get-ChildItem -LiteralPath $logFolder -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

$stopwatch.Stop()
Write-Log ("=== Vendor Version Monitor completed in {0:N1}s ===" -f $stopwatch.Elapsed.TotalSeconds)
Write-Log ("    Total: {0} | Current: {1} | Stale: {2} | Error: {3}" -f $appsToScan.Count, $currentApps.Count, $staleApps.Count, @($appsToScan | Where-Object { $_.Status -like 'Error*' }).Count)

if ($staleApps.Count -gt 0) {
    Write-Log "    Stale applications:"
    foreach ($s in $staleApps) {
        $cveInfo = if ($s.CVECount -gt 0) { " | {0} CVEs (max CVSS {1:N1})" -f $s.CVECount, $s.MaxCVSS } else { '' }
        Write-Log ("      - {0}: {1} -> {2}{3}" -f $s.Application, $s.MecmVersion, $s.VendorVersion, $cveInfo)
    }
}

Write-Host "`nReport: $reportPath"
