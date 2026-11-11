#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the linter and the Pester tests of App Packager in groups, for CI.

.DESCRIPTION
    One command for CI and for the lab. The Checks workflow runs one group in each job:
      Lint     PSScriptAnalyzer on every PowerShell file that the repository owns.
               A finding of severity Error fails the run. Warning and Information
               findings are counted by rule and do not fail the run (see the note at
               the lint stage).
      Modules  Pester tests of the shared modules (Packagers\*.Tests.ps1).
      Smoke    The offline packager smoke test (Tests\PackagerSmoke.Tests.ps1).
               It runs the five offline checks against every packager.
      Tests    Every other Pester test in Tests\.
      Probes   The window probes (Tests\Invoke-*Smoke.ps1) under Windows PowerShell
               in an STA thread. They open windows, so they run only when
               SUITE_ALLOW_WINDOW_TESTS is 1. Without it the group is skipped.
      All      Lint, Modules, Smoke, Tests and Probes.
    A failed test fails the run. A group that finds no test fails too.
    Nothing here contacts a vendor, a lab computer or a certificate trust store.

.PARAMETER Group
    The group to run. Default All.

.EXAMPLE
    .\Invoke-Checks.ps1 -Group Modules
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console harness: Write-Host is the output channel.')]
[CmdletBinding()]
param(
    [ValidateSet('All', 'Lint', 'Modules', 'Smoke', 'Tests', 'Probes')]
    [string]$Group = 'All'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$failed = $false

function Test-GroupSelected {
    param([string]$Name)
    return ($Group -eq 'All' -or $Group -eq $Name)
}

# ---- Lint ------------------------------------------------------------------
if (Test-GroupSelected 'Lint') {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
    # Vendored code (Lib) and generated folders are not linted.
    $skipFolders = @('Lib', 'Logs', 'Icons', 'out', '.git', '.github')
    $targets = @(Get-ChildItem -LiteralPath $root -Recurse -File -Include '*.ps1', '*.psm1', '*.psd1' | Where-Object {
            $relative = $_.FullName.Substring($root.Length).TrimStart('\', '/')
            $top = ($relative -split '[\\/]')[0]
            $skipFolders -notcontains $top -and $relative -notlike 'Tests*\Fixtures\*' -and $relative -notlike 'Tests*/Fixtures/*'
        })
    $settings = Join-Path $root 'Tests\PSScriptAnalyzerSettings.psd1'
    # Test files pass state to mock bodies through a global variable, and they define variables in BeforeAll that the It blocks read. Both are intended.
    $testOnlyExclude = 'PSAvoidGlobalVars', 'PSUseDeclaredVarsMoreThanAssignments'
    # The tests build a throwaway password as a SecureString from plain text. The start script protects a key that an earlier build saved in plain text.
    # Both are by design, so the rule is off for every file.
    $allExclude = 'PSAvoidUsingConvertToSecureStringWithPlainText'
    $findings = @(foreach ($t in $targets) {
            if ($t.Name -like '*.Tests.ps1' -or $t.FullName -like '*\Tests\*') { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings -ExcludeRule ($testOnlyExclude + $allExclude) }
            else { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings -ExcludeRule $allExclude }
        })
    $blocking = @($findings | Where-Object { $_.Severity -eq 'Error' })
    Write-Host ('Lint: {0} files, {1} findings, {2} with severity Error' -f $targets.Count, $findings.Count, $blocking.Count)
    # The repository has a large set of older Warning and Information findings. The gate fails on Error only,
    # so a new Error fails the run. The counts by rule show the size of the rest.
    $findings | Group-Object -Property Severity, RuleName | Sort-Object -Property Count -Descending | ForEach-Object { Write-Host ('  {0,5}  {1}' -f $_.Count, $_.Name) }
    foreach ($f in $blocking) { Write-Host ('  ERROR {0}:{1} [{2}] {3}' -f (Split-Path $f.ScriptPath -Leaf), $f.Line, $f.RuleName, $f.Message) }
    if ($blocking.Count -gt 0) { $failed = $true }
}

# ---- Pester groups ---------------------------------------------------------
$pesterGroups = [ordered]@{
    Modules = @(Get-ChildItem -LiteralPath (Join-Path $root 'Packagers') -Filter '*.Tests.ps1' -File)
    Smoke   = @(Get-ChildItem -LiteralPath (Join-Path $root 'Tests') -Filter 'PackagerSmoke.Tests.ps1' -File)
    Tests   = @(Get-ChildItem -LiteralPath (Join-Path $root 'Tests') -Filter '*.Tests.ps1' -File | Where-Object { $_.Name -ne 'PackagerSmoke.Tests.ps1' })
}
foreach ($name in $pesterGroups.Keys) {
    if (-not (Test-GroupSelected $name)) { continue }
    # Pester 5 only: the tests use the 5.x configuration and syntax, and Pester 6 changes them.
    Import-Module Pester -MinimumVersion 5.0 -MaximumVersion 5.99.99 -ErrorAction Stop
    $files = @($pesterGroups[$name])
    if ($files.Count -eq 0) { Write-Host ('Group {0}: no test file found.' -f $name); $failed = $true; continue }
    $cfg = New-PesterConfiguration
    $cfg.Run.Path = $files.FullName
    $cfg.Run.PassThru = $true
    $cfg.Output.Verbosity = 'Minimal'
    # Native commands that the tests start write to stderr. Under 'Stop' Windows PowerShell 5.1 turns each line into a terminating error, so the tests run under 'Continue'.
    $ErrorActionPreference = 'Continue'
    $result = Invoke-Pester -Configuration $cfg
    $ErrorActionPreference = 'Stop'
    Write-Host ('{0}: {1} files, {2} run, {3} passed, {4} failed, {5} skipped, {6} containers failed' -f $name, $files.Count, $result.TotalCount, $result.PassedCount, $result.FailedCount, $result.SkippedCount, $result.FailedContainersCount)
    if ($result.TotalCount -eq 0) { Write-Host ('Group {0}: no test ran.' -f $name); $failed = $true }
    if ($result.FailedCount -gt 0 -or $result.FailedContainersCount -gt 0 -or $result.FailedBlocksCount -gt 0) { $failed = $true }
    if ($result.Result -ne 'Passed') { Write-Host ('Pester result: {0}' -f $result.Result); $failed = $true }
}

# ---- Window probes ---------------------------------------------------------
# The probes open windows. They run only with SUITE_ALLOW_WINDOW_TESTS=1, as in the other suite tools.
if (Test-GroupSelected 'Probes') {
    if ($env:SUITE_ALLOW_WINDOW_TESTS -ne '1') {
        Write-Host 'Window probes skipped: SUITE_ALLOW_WINDOW_TESTS is not 1.'
    }
    else {
        # Invoke-FullRegression.ps1 runs each probe in its own STA child process with a time limit.
        & (Join-Path $root 'Tests\Invoke-FullRegression.ps1') -SkipSmoke -SkipPester
        if ($LASTEXITCODE -ne 0) { $failed = $true }
    }
}

if ($failed) { Write-Host 'CHECKS FAILED'; exit 1 }
Write-Host 'CHECKS PASSED'
exit 0
