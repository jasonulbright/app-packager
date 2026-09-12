#Requires -Version 5.1

<#
.SYNOPSIS
    Runs every automated check in this repository and prints one summary table.

.DESCRIPTION
    Four stages, each in its own child process so a crashing host cannot take
    the runner with it:

      1. Offline packager smoke, with the level-1 catalog matrix.
      2. The combined Pester set under Windows PowerShell 5.1.
      3. The same Pester set under PowerShell 7, when pwsh is installed.
      4. Every Tests\Invoke-*Smoke.ps1 UI probe under powershell.exe -STA.

    Nothing here downloads an installer, contacts a vendor, touches a lab
    client, or writes to a certificate trust store. Stages that need one of
    those are skipped by name and reported as Skipped, which is not Passed.

    Exits non-zero when any stage fails.

.EXAMPLE
    .\Tests\Invoke-FullRegression.ps1

.EXAMPLE
    .\Tests\Invoke-FullRegression.ps1 -SkipPwsh7 -MatrixPath C:\temp\ap-matrix\catalog-matrix.csv
#>

[CmdletBinding()]
param(
    [string]$MatrixPath,

    [switch]$SkipSmoke,

    [switch]$SkipPester,

    [switch]$SkipPwsh7,

    [switch]$SkipUiProbes,

    # Needs a local copy of the vendor installer and re-runs itself elevated,
    # so it is not part of an unattended regression run.
    [string[]]$SkipProbe = @('Invoke-AdobeExtractSmoke.ps1'),

    [int]$ProbeTimeoutSec = 300,

    [int]$PesterTimeoutSec = 2700
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$script:Stages = New-Object System.Collections.ArrayList
$script:WorkFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('ap-regression-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $script:WorkFolder -Force | Out-Null

function Add-StageResult {
    param(
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][string]$HostLabel,
        [Parameter(Mandatory)][ValidateSet('Passed', 'Failed', 'Skipped')][string]$Status,
        [int]$Passed = 0,
        [int]$Failed = 0,
        [int]$Skipped = 0,
        [double]$Seconds = 0,
        [string]$Detail = ''
    )
    [void]$script:Stages.Add([pscustomobject]@{
        Stage   = $Stage
        Host    = $HostLabel
        Status  = $Status
        Passed  = $Passed
        Failed  = $Failed
        Skipped = $Skipped
        Seconds = [math]::Round($Seconds, 1)
        Detail  = $Detail
    })
}

function Invoke-RegressionProcess {
    <#
        Runs one child process with a timeout and returns its exit code and
        both streams. Both streams are drained by ReadToEnd tasks: a
        DataReceived handler would run on a thread-pool thread with no
        runspace and end the host without a message.
    #>
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [Parameter(Mandatory)][int]$TimeoutSec
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FileName
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.WorkingDirectory = $script:RepoRoot
    if ($startInfo.GetType().GetProperty('ArgumentList')) {
        foreach ($argument in $ArgumentList) { [void]$startInfo.ArgumentList.Add($argument) }
    }
    else {
        $startInfo.Arguments = (($ArgumentList | ForEach-Object {
            if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
        }) -join ' ')
    }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $started = [DateTime]::UtcNow
    [void]$process.Start()
    $outTask = $process.StandardOutput.ReadToEndAsync()
    $errTask = $process.StandardError.ReadToEndAsync()

    $timedOut = $false
    if (-not $process.WaitForExit($TimeoutSec * 1000)) {
        $timedOut = $true
        try { & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null } catch { }
        try { if (-not $process.HasExited) { $process.Kill() } } catch { }
        $process.WaitForExit()
    }

    $result = [pscustomobject]@{
        ExitCode = $(if ($timedOut) { -1 } else { $process.ExitCode })
        TimedOut = $timedOut
        StdOut   = [string]$outTask.Result
        StdErr   = [string]$errTask.Result
        Seconds  = ([DateTime]::UtcNow - $started).TotalSeconds
    }
    $process.Dispose()
    return $result
}

function Get-PesterRunnerScript {
    # One Pester invocation, written to a file so neither host has to quote a
    # long -Command string. The counts go to a JSON file rather than a stream:
    # the suites under test write to the host themselves, and 5.1 wraps a
    # redirected error stream in CLIXML.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$ResultPath)

    $body = @'
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath '__ROOT__'
Import-Module Pester -RequiredVersion 5.7.1 -Force
$paths = @(
    '.\Packagers\AppPackagerCommon.Tests.ps1'
    '.\Packagers\AppPackagerWorkbench.Tests.ps1'
    '.\Packagers\AppPackagerSigning.Tests.ps1'
    '.\Tests\PackagerSmoke.Tests.ps1'
    '.\Tests\PackageWorkflow.Tests.ps1'
    '.\Tests\UserDetection.Tests.ps1'
    '.\Tests\SigningCombinations.Tests.ps1'
    '.\Tests\UpdatePreservation.Tests.ps1'
) | Where-Object { Test-Path -LiteralPath $_ }
$result = Invoke-Pester -Path $paths -PassThru -Output None
$failures = @($result.Failed | ForEach-Object { $_.ExpandedPath + " :: " + [string]$_.ErrorRecord.Exception.Message.Split([char]10)[0] })
$payload = @{
    Passed  = $result.PassedCount
    Failed  = $result.FailedCount
    Skipped = $result.SkippedCount
    Failures = $failures
} | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText('__RESULT__', $payload)
if ($result.FailedCount -gt 0) { exit 1 }
exit 0
'@
    $body = $body.Replace('__ROOT__', $script:RepoRoot).Replace('__RESULT__', $ResultPath)
    [System.IO.File]::WriteAllText($Path, $body, (New-Object System.Text.UTF8Encoding($false)))
    return $Path
}

function Invoke-PesterStage {
    param(
        [Parameter(Mandatory)][string]$HostLabel,
        [Parameter(Mandatory)][string]$Executable,
        [Parameter(Mandatory)][string]$RunnerScript,
        [Parameter(Mandatory)][string]$ResultPath
    )

    if (Test-Path -LiteralPath $ResultPath) { Remove-Item -LiteralPath $ResultPath -Force }
    $result = Invoke-RegressionProcess -FileName $Executable `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $RunnerScript) `
        -TimeoutSec $PesterTimeoutSec

    if ($result.TimedOut -or -not (Test-Path -LiteralPath $ResultPath)) {
        $detail = if ($result.TimedOut) { "Timed out after ${PesterTimeoutSec}s." } else { 'Pester wrote no result file; the host ended early.' }
        Add-StageResult -Stage 'Pester (combined)' -HostLabel $HostLabel -Status Failed -Seconds $result.Seconds -Detail $detail
        return
    }

    $summary = [System.IO.File]::ReadAllText($ResultPath) | ConvertFrom-Json
    $status = if ([int]$summary.Failed -gt 0 -or $result.ExitCode -ne 0) { 'Failed' } else { 'Passed' }
    $detail = if ([int]$summary.Failed -gt 0) { (@($summary.Failures) -join '; ') } else { '' }
    Add-StageResult -Stage 'Pester (combined)' -HostLabel $HostLabel -Status $status `
        -Passed ([int]$summary.Passed) -Failed ([int]$summary.Failed) -Skipped ([int]$summary.Skipped) `
        -Seconds $result.Seconds -Detail $detail
}

function Get-ProbeVerdict {
    <#
        The UI probes report in two shapes: "PASS  <case>" lines closed by
        "<n> passed, <m> failed.", or "PASS: <case>" lines with a throw on
        failure. Both are read here so neither convention has to change.
    #>
    param(
        [AllowEmptyString()][string]$Output,
        [AllowEmptyString()][string]$ErrorOutput,
        [int]$ExitCode
    )

    $passed = 0
    $failed = 0
    $tally = [regex]::Match($Output, '(?im)^\s*(\d+)\s+passed,\s*(\d+)\s+failed')
    if ($tally.Success) {
        $passed = [int]$tally.Groups[1].Value
        $failed = [int]$tally.Groups[2].Value
    }
    else {
        $checks = [regex]::Match($Output, '(?im)^PASS[:,]?\s.*?(\d+)\s+checks?')
        if ($checks.Success) { $passed = [int]$checks.Groups[1].Value }
        else { $passed = @([regex]::Matches($Output, '(?im)^\s*PASS[: ]')).Count }
        $failed = @([regex]::Matches($Output, '(?im)^\s*FAIL[: ]')).Count
    }

    $detail = ''
    $ok = $true
    if ($ExitCode -ne 0) { $ok = $false; $detail = "Exit $ExitCode." }
    if ($failed -gt 0) { $ok = $false; $detail = ("{0} case(s) failed." -f $failed) }
    if ($passed -eq 0 -and $failed -eq 0) { $ok = $false; $detail = 'No PASS output.' }
    if (-not $ok) {
        $tail = ($ErrorOutput.Trim(), (@($Output -split "`r?`n" | Where-Object { $_ -match '(?i)^\s*FAIL' }) -join ' ') |
            Where-Object { $_ }) -join ' '
        if ($tail) { $detail = ($detail + ' ' + $tail).Trim() }
        if ($detail.Length -gt 240) { $detail = $detail.Substring(0, 240) + '...' }
    }

    return [pscustomobject]@{ Ok = $ok; Passed = $passed; Failed = $failed; Detail = $detail }
}

$powershellExe = (Get-Command powershell.exe -ErrorAction Stop).Source

try {
    # --- 1. Offline packager smoke ---------------------------------------
    if ($SkipSmoke) {
        Add-StageResult -Stage 'Offline smoke' -HostLabel 'powershell 5.1' -Status Skipped -Detail 'Skipped by -SkipSmoke.'
    }
    else {
        $smokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
            (Join-Path $PSScriptRoot 'Invoke-PackagerSmoke.ps1'), '-Matrix', '-Json')
        if ($MatrixPath) { $smokeArgs += @('-MatrixPath', $MatrixPath) }
        $smoke = Invoke-RegressionProcess -FileName $powershellExe -ArgumentList $smokeArgs -TimeoutSec $PesterTimeoutSec

        if ($smoke.TimedOut) {
            Add-StageResult -Stage 'Offline smoke' -HostLabel 'powershell 5.1' -Status Failed -Seconds $smoke.Seconds -Detail 'Timed out.'
        }
        else {
            $start = $smoke.StdOut.IndexOf('{')
            if ($start -lt 0) {
                Add-StageResult -Stage 'Offline smoke' -HostLabel 'powershell 5.1' -Status Failed -Seconds $smoke.Seconds -Detail ($smoke.StdErr.Trim())
            }
            else {
                $summary = $smoke.StdOut.Substring($start) | ConvertFrom-Json
                $status = if ([int]$summary.Failed -gt 0 -or $smoke.ExitCode -ne 0) { 'Failed' } else { 'Passed' }
                $detail = "$($summary.Packagers) script(s)"
                if ($summary.PSObject.Properties['MatrixPath'] -and $summary.MatrixPath) { $detail += "; matrix $($summary.MatrixPath)" }
                Add-StageResult -Stage 'Offline smoke' -HostLabel 'powershell 5.1' -Status $status `
                    -Passed ([int]$summary.Passed) -Failed ([int]$summary.Failed) -Skipped ([int]$summary.Skipped) `
                    -Seconds $smoke.Seconds -Detail $detail
            }
        }
    }

    # --- 2 and 3. Combined Pester on both hosts ---------------------------
    if ($SkipPester) {
        Add-StageResult -Stage 'Pester (combined)' -HostLabel 'powershell 5.1' -Status Skipped -Detail 'Skipped by -SkipPester.'
        Add-StageResult -Stage 'Pester (combined)' -HostLabel 'pwsh 7' -Status Skipped -Detail 'Skipped by -SkipPester.'
    }
    else {
        $resultPath = Join-Path $script:WorkFolder 'pester-result.json'
        $runnerScript = Get-PesterRunnerScript -Path (Join-Path $script:WorkFolder 'run-pester.ps1') -ResultPath $resultPath
        Invoke-PesterStage -HostLabel 'powershell 5.1' -Executable $powershellExe -RunnerScript $runnerScript -ResultPath $resultPath

        $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if ($SkipPwsh7) {
            Add-StageResult -Stage 'Pester (combined)' -HostLabel 'pwsh 7' -Status Skipped -Detail 'Skipped by -SkipPwsh7.'
        }
        elseif (-not $pwsh) {
            Add-StageResult -Stage 'Pester (combined)' -HostLabel 'pwsh 7' -Status Skipped -Detail 'pwsh.exe is not installed.'
        }
        else {
            Invoke-PesterStage -HostLabel 'pwsh 7' -Executable $pwsh.Source -RunnerScript $runnerScript -ResultPath $resultPath
        }
    }

    # --- 4. UI probes -----------------------------------------------------
    $probes = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'Invoke-*Smoke.ps1' -File |
        Where-Object { $_.Name -ne 'Invoke-PackagerSmoke.ps1' } | Sort-Object Name)
    foreach ($probe in $probes) {
        if ($SkipUiProbes) {
            Add-StageResult -Stage $probe.BaseName -HostLabel 'powershell 5.1 -STA' -Status Skipped -Detail 'Skipped by -SkipUiProbes.'
            continue
        }
        if ($SkipProbe -contains $probe.Name) {
            Add-StageResult -Stage $probe.BaseName -HostLabel 'powershell 5.1 -STA' -Status Skipped -Detail 'Needs a local installer or elevation.'
            continue
        }

        # Called with the call operator, not -File: the launch-scope probe
        # proves behavior that only holds when the script is dot-run into the
        # caller's scope.
        $result = Invoke-RegressionProcess -FileName $powershellExe `
            -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-Command',
                ('& "' + $probe.FullName + '"; exit $LASTEXITCODE')) `
            -TimeoutSec $ProbeTimeoutSec

        $verdict = Get-ProbeVerdict -Output $result.StdOut -ErrorOutput $result.StdErr -ExitCode $result.ExitCode
        if ($result.TimedOut) {
            Add-StageResult -Stage $probe.BaseName -HostLabel 'powershell 5.1 -STA' -Status Failed -Seconds $result.Seconds -Detail "Timed out after ${ProbeTimeoutSec}s."
        }
        elseif (-not $verdict.Ok) {
            Add-StageResult -Stage $probe.BaseName -HostLabel 'powershell 5.1 -STA' -Status Failed `
                -Passed $verdict.Passed -Failed ([math]::Max($verdict.Failed, 1)) -Seconds $result.Seconds -Detail $verdict.Detail
        }
        else {
            Add-StageResult -Stage $probe.BaseName -HostLabel 'powershell 5.1 -STA' -Status Passed -Passed $verdict.Passed -Seconds $result.Seconds
        }
    }
}
finally {
    Remove-Item -LiteralPath $script:WorkFolder -Recurse -Force -ErrorAction SilentlyContinue
}

$rows = @($script:Stages.ToArray())
Write-Host ''
Write-Host 'AppPackager full regression'
Write-Host ('=' * 96)
$rows | Format-Table -AutoSize Stage, Host, Status, Passed, Failed, Skipped, Seconds | Out-String -Width 200 | Write-Host

$failedRows = @($rows | Where-Object { $_.Status -eq 'Failed' })
Write-Host ('Stages: {0} passed, {1} failed, {2} skipped. Assertions/checks passed: {3}, failed: {4}.' -f
    @($rows | Where-Object { $_.Status -eq 'Passed' }).Count,
    $failedRows.Count,
    @($rows | Where-Object { $_.Status -eq 'Skipped' }).Count,
    (@($rows | Measure-Object -Property Passed -Sum).Sum),
    (@($rows | Measure-Object -Property Failed -Sum).Sum))

if ($failedRows.Count -gt 0) {
    Write-Host ''
    Write-Host 'Failures:'
    foreach ($row in $failedRows) {
        Write-Host ("  [{0}] {1}: {2}" -f $row.Host, $row.Stage, $row.Detail)
    }
    exit 1
}

exit 0
