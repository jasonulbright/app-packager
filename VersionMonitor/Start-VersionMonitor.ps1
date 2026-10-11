#Requires -Version 5.1
<#
.SYNOPSIS
    Scheduled entry point for the CVE monitor.

.DESCRIPTION
    Reads monitor-config.json beside this script and starts the application
    in batch mode with -CveReport. The run refreshes the latest and the
    ConfigMgr versions of every packager, looks up the CVEs that the update
    of each stale application fixes, writes the report in every format, and
    notifies through the drop folder and the webhook. It shares
    app-history.json and the CVE cache with the GUI, so the next interactive
    session shows the same versions and counts.

    Every parameter here overrides the value from the config file.

.PARAMETER ConfigPath
    Path to monitor-config.json. Defaults to the file beside this script. A
    missing file is created with default values on the first run.

.PARAMETER SkipLatest
    Do not query the vendors; use the latest versions from app-history.json.

.PARAMETER SkipMECM
    Do not query ConfigMgr; use the ConfigMgr versions from app-history.json.

.PARAMETER SkipNVD
    Do not look CVEs up; write the version report only.

.PARAMETER Apps
    Packager base names (package-chrome, package-7zip) to limit the run.

.PARAMETER OutputFolder
    Folder for the report files. Defaults to Report.OutputFolder from the
    config, then to Logs\Reports under the application folder.

.EXAMPLE
    .\Start-VersionMonitor.ps1

.EXAMPLE
    .\Start-VersionMonitor.ps1 -SkipMECM -SkipNVD

.EXAMPLE
    .\Start-VersionMonitor.ps1 -Apps package-chrome, package-firefox -SkipLatest
#>

[CmdletBinding()]
param(
    # $PSScriptRoot is empty inside a parameter default under Windows
    # PowerShell 5.1 with -File; the path is resolved below.
    [string]$ConfigPath = '',
    [switch]$SkipLatest,
    [switch]$SkipMECM,
    [switch]$SkipNVD,
    [string[]]$Apps,
    [Alias('OutputPath')][string]$OutputFolder,
    [string]$LogPath,
    # Accepted from earlier scheduled tasks; the simulation mode is gone.
    [switch]$SimulateStale,
    [string]$OverridesPath
)

if ($SimulateStale -or $OverridesPath) {
    Write-Warning 'The simulation mode (-SimulateStale, -OverridesPath) is no longer supported; the run continues without it.'
}

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $scriptRoot 'monitor-config.json' }
$appRoot = Split-Path $scriptRoot -Parent
$entry = Join-Path $appRoot 'start-apppackager.ps1'
if (-not (Test-Path -LiteralPath $entry)) { throw "start-apppackager.ps1 not found at: $entry" }

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    $default = [ordered]@{
        SchemaVersion = 2
        MECM = [ordered]@{ SiteCode = ''; ServerFQDN = '' }
        NVD = [ordered]@{ ApiKey = ''; CacheTtlMinutes = 360 }
        Report = [ordered]@{ OutputFolder = ''; Formats = @('Html', 'Csv', 'Json', 'Xlsx', 'Markdown'); KeepReportDays = 90 }
        Logging = [ordered]@{ LogFolder = ''; KeepLogDays = 30 }
        Notifications = [ordered]@{ Enabled = $false; DropFolder = ''; WebhookUrl = ''; OnlyOnStale = $true }
    }
    $default | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    Write-Host ('[monitor] Wrote default config: {0}' -f $ConfigPath)
}
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$get = { param($obj, $name, $fallback) if ($obj -and $obj.PSObject.Properties[$name] -and $null -ne $obj.PSObject.Properties[$name].Value) { $obj.PSObject.Properties[$name].Value } else { $fallback } }

$logFolder = [string](& $get $config.Logging 'LogFolder' '')
if (-not $logFolder) { $logFolder = Join-Path $scriptRoot 'Logs' }
if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
if (-not $LogPath) { $LogPath = Join-Path $logFolder ('VersionMonitor-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss')) }

$keepLogDays = [int](& $get $config.Logging 'KeepLogDays' 30)
if ($keepLogDays -gt 0) {
    $cutoff = (Get-Date).AddDays(-$keepLogDays)
    Get-ChildItem -LiteralPath $logFolder -Filter 'VersionMonitor-*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

$arguments = @{
    BatchMode     = $true
    CveReport     = $true
    LogPath       = $LogPath
    ReportFormats = @(& $get $config.Report 'Formats' @('Html', 'Csv', 'Json', 'Xlsx', 'Markdown'))
    KeepReportDays = [int](& $get $config.Report 'KeepReportDays' 0)
}
$siteCode = [string](& $get $config.MECM 'SiteCode' '')
if ($siteCode) { $arguments['SiteCode'] = $siteCode }
$provider = [string](& $get $config.MECM 'ServerFQDN' '')
if ($provider) { $arguments['ProviderMachineName'] = $provider }
if (-not $OutputFolder) { $OutputFolder = [string](& $get $config.Report 'OutputFolder' '') }
if ($OutputFolder) { $arguments['ReportFolder'] = $OutputFolder }
if ($Apps) { $arguments['Apps'] = $Apps }
if ($SkipLatest) { $arguments['SkipLatest'] = $true }
if ($SkipMECM) { $arguments['SkipMecm'] = $true }
if ($SkipNVD) { $arguments['SkipCve'] = $true }

$apiKey = [string](& $get $config.NVD 'ApiKey' '')
if ($apiKey -and -not $env:NVD_API_KEY) { $env:NVD_API_KEY = $apiKey }

$notify = $config.Notifications
if ($notify -and [bool](& $get $notify 'Enabled' $false)) {
    $drop = [string](& $get $notify 'DropFolder' '')
    $hook = [string](& $get $notify 'WebhookUrl' '')
    if ($drop) { $arguments['DropFolder'] = $drop }
    if ($hook) { $arguments['WebhookUrl'] = $hook }
    if (-not [bool](& $get $notify 'OnlyOnStale' $true)) { $arguments['NotifyAlways'] = $true }
}

& $entry @arguments
exit $LASTEXITCODE
