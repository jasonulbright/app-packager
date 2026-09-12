#Requires -Modules Pester

<#
.SYNOPSIS
    Pester wrapper for the offline packager smoke harness.

.DESCRIPTION
    Runs Invoke-PackagerSmoke.ps1 in a child PowerShell process so failures
    are reported as Pester failures without terminating the test host.
#>

Describe 'Packager smoke harness' {
    It 'passes offline parse, metadata, and parameter-contract checks for all packagers' {
        $scriptPath = Join-Path $PSScriptRoot 'Invoke-PackagerSmoke.ps1'
        $stdoutPath = Join-Path $TestDrive 'packager-smoke.json'
        $stderrPath = Join-Path $TestDrive 'packager-smoke.err'
        $powershellExe = (Get-Command powershell.exe -ErrorAction Stop).Source

        $proc = Start-Process `
            -FilePath $powershellExe `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-Json') `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
        $stdout = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw } else { '' }

        $proc.ExitCode | Should -Be 0 -Because $stderr
        $summary = $stdout | ConvertFrom-Json
        $summary.Packagers | Should -BeGreaterThan 0
        $summary.Failed | Should -Be 0
        $summary.Passed | Should -BeGreaterThan 0
    }
}
