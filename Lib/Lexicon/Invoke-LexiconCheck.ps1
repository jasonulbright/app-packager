<#
.SYNOPSIS
    Scans the user-facing text of one repository against the ConfigMgr lexicon and reports the hits.

.DESCRIPTION
    Windows PowerShell 5.1 and PowerShell 7. Needs no ConfigMgr site and no network.
    Use it in CI for the C# repository APD (-RepoType Apd), or by hand for a report.
    Each hit shows file:line, the matched phrase, the lexicon id and the replacement rule.
    The exit code is 1 when -FailOn finds a hit, otherwise 0. Without -FailOn the script only reports.

.PARAMETER Root
    The repository folder to scan.

.PARAMETER RepoType
    Auto (default), PowerShell or Apd. Apd also scans src\**\*.cs string literals and the XAML files under src.

.PARAMETER LexiconPath
    Path of configmgr-lexicon.json. The default is the file next to this script.

.PARAMETER AllowListPath
    Path of lexicon-allowlist.json. The default is the file next to this script.

.PARAMETER FailOn
    None (default), Error, or Warning. Error fails on severity error. Warning fails on error and warning.
    Severity info never fails.

.PARAMETER OutputPath
    Optional. Writes the hits as CSV (ASCII).

.EXAMPLE
    .\Invoke-LexiconCheck.ps1 -Root C:\work\app-packager-desktop -RepoType Apd -FailOn Error
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Operator-facing console report.')]
param(
    [Parameter(Mandatory)][string]$Root,
    [ValidateSet('Auto', 'PowerShell', 'Apd')][string]$RepoType = 'Auto',
    [string]$LexiconPath = (Join-Path $PSScriptRoot 'configmgr-lexicon.json'),
    [string]$AllowListPath = (Join-Path $PSScriptRoot 'lexicon-allowlist.json'),
    [ValidateSet('None', 'Error', 'Warning')][string]$FailOn = 'None',
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Lexicon.psm1') -Force
$hits = @(Invoke-LexiconCheck -Root $Root -LexiconPath $LexiconPath -AllowListPath $AllowListPath -RepoType $RepoType)
# Banned names: every line of every tracked file (entries with scope tracked-files).
$banLexicon = Import-LexiconData -Path $LexiconPath
$banAllow = @(Import-LexiconAllowList -Path $AllowListPath)
$hits = @($hits) + @(Get-BannedNameHit -Root $Root -Lexicon $banLexicon -AllowList $banAllow)

foreach ($h in ($hits | Sort-Object RuleId, File, Line)) {
    Write-Host ('{0}:{1}: [{2} {3}] "{4}" -> {5}' -f $h.File, $h.Line, $h.RuleId, $h.Severity, $h.Phrase, $h.Replacement)
}
Write-Host ''
Write-Host ('{0} hits ({1} error, {2} warning, {3} info)' -f $hits.Count,
    @($hits | Where-Object Severity -eq 'error').Count,
    @($hits | Where-Object Severity -eq 'warning').Count,
    @($hits | Where-Object Severity -eq 'info').Count)

if ($OutputPath) {
    $hits | Select-Object File, Line, Kind, RuleId, Severity, Phrase, Text | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding ASCII
}

$fail = 0
if ($FailOn -eq 'Error') { $fail = @($hits | Where-Object Severity -eq 'error').Count }
if ($FailOn -eq 'Warning') { $fail = @($hits | Where-Object { $_.Severity -in 'error', 'warning' }).Count }
if ($fail -gt 0) { exit 1 }
exit 0
