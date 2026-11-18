#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Parse-only guard (vault: testing-policy, "Process rules on a lab VM"). A script in tools\ stops only a process that it started and
# knows by process id. It never stops a process by a command-line pattern, because another session may own a process that matches.
# Nothing runs and no window opens.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ToolDir = Join-Path $script:Root 'tools'
    $script:Files = @(Get-ChildItem -LiteralPath $script:ToolDir -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.cmd', '.bat' })
    # Every code line of every tool script, as 'file:line text'. Comment lines are skipped.
    $script:Lines = foreach ($f in $script:Files) {
        $n = 0
        foreach ($l in (Get-Content -LiteralPath $f.FullName)) {
            $n++
            if ($l.TrimStart().StartsWith('#')) { continue }
            [pscustomobject]@{ Where = ('{0}:{1}' -f $f.Name, $n); Text = $l }
        }
    }
    function Get-Hits([string]$Pattern) { @($script:Lines | Where-Object { $_.Text -match $Pattern } | ForEach-Object { $_.Where }) }
}

Describe 'The tools stop only their own processes' {
    It 'finds the tool scripts' {
        $script:Files.Count | Should -BeGreaterThan 0
    }

    It 'never calls taskkill' {
        Get-Hits '(?i)\btaskkill\b' | Should -BeNullOrEmpty
    }

    It 'never matches a process by its command line' {
        Get-Hits 'CommandLine' | Should -BeNullOrEmpty
    }

    It 'stops a process only with Stop-Process -Id' {
        Get-Hits 'Stop-Process(?!\s+-Id\b)' | Should -BeNullOrEmpty
    }

    It 'never selects the processes to stop by name or by a filter' {
        Get-Hits 'Get-Process\s+(?!-Id\b)[^\r\n]*Stop-Process' | Should -BeNullOrEmpty
        Get-Hits 'Where-Object[^\r\n]*Stop-Process' | Should -BeNullOrEmpty
    }

    It 'reads Win32_Process only to find the children of a process by its id' {
        $bad = @($script:Lines | Where-Object { $_.Text -match 'Win32_Process' -and $_.Text -notmatch 'ParentProcessId' } | ForEach-Object { $_.Where })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'The capture records the processes that it starts' {
    It 'keeps a ledger of process ids in its own run folder' {
        $capture = Get-Content -LiteralPath (Join-Path $script:ToolDir 'guide-capture\Invoke-GuideCapture.ps1') -Raw
        $capture | Should -Match '\$stopRecorded = @'''
        $capture | Should -Match 'run\.pids'
        (Get-Content -LiteralPath (Join-Path $script:ToolDir 'guide-capture\main.ps1') -Raw) | Should -Match 'run\.pids'
        $lib = Get-Content -LiteralPath (Join-Path $script:ToolDir 'guide-capture\guilib.ps1') -Raw
        $lib | Should -Match 'function Register-RunProcess'
        $lib | Should -Match 'Register-RunProcess \$script:proc \$Script'
    }

    It 'stops a recorded process only when its start time still matches' {
        $capture = Get-Content -LiteralPath (Join-Path $script:ToolDir 'guide-capture\Invoke-GuideCapture.ps1') -Raw
        $capture | Should -Match 'StartTime\.Ticks -eq \[long\]\$ticks'
    }
}