#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The test runners stop only what they started. Real processes, no window.

.DESCRIPTION
    A stop of a process tree by parent ids (taskkill /T) can reach an older, unrelated process whose parent had the same id,
    because Windows uses process ids again. A lab script with such a walk ended the signed-in session of a lab client. The
    runners start every child in a job object (Tests\ProcessJob.ps1) and stop only the job.
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'ProcessJob.ps1')
    $script:Shell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:Sleep = '-NoProfile -NonInteractive -Command Start-Sleep -Seconds 120'

    function Wait-File([string]$Path, [int]$Seconds = 60) {
        $until = (Get-Date).AddSeconds($Seconds)
        while ((Get-Date) -lt $until) {
            $text = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
            if ($text -and $text.EndsWith("`n")) { return $text.Trim() }
            Start-Sleep -Milliseconds 200
        }
        throw "no text in $Path after $Seconds s"
    }
}

Describe 'The job of a test runner' {
    It 'ends the child and the grandchild whose parent is gone, and leaves a process with a dead parent that it did not start' {
        # The unrelated process: its parent has ended, so only a parent id links it to anything.
        $otherFile = Join-Path $TestDrive 'other.txt'
        $otherParent = Join-Path $TestDrive 'other-parent.ps1'
        Set-Content -LiteralPath $otherParent -Encoding Ascii -Value "`$c = Start-Process '$($script:Shell)' -ArgumentList '$($script:Sleep)' -PassThru -WindowStyle Hidden; Set-Content -LiteralPath '$otherFile' -Value `$c.Id"
        $starter = Start-Process $script:Shell -PassThru -WindowStyle Hidden -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $otherParent + '"')
        $null = $starter.Handle
        $other = Get-Process -Id ([int](Wait-File $otherFile)); $null = $other.Handle
        $starter.WaitForExit(30000) | Should -BeTrue

        $childFile = Join-Path $TestDrive 'child.txt'
        $middle = Join-Path $TestDrive 'middle.ps1'
        Set-Content -LiteralPath $middle -Encoding Ascii -Value "`$c = Start-Process '$($script:Shell)' -ArgumentList '$($script:Sleep)' -PassThru -WindowStyle Hidden; Set-Content -LiteralPath '$childFile' -Value `$c.Id"
        $child = $null
        $run = Start-JobProcess -FilePath $script:Shell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$middle`""
        try {
            $child = Get-Process -Id ([int](Wait-File $childFile)); $null = $child.Handle
            $run.Process.WaitForExit(30000) | Should -BeTrue
            Stop-JobProcess -Run $run
            $child.WaitForExit(15000) | Should -BeTrue
            $other.WaitForExit(2000) | Should -BeFalse
        }
        finally {
            Stop-JobProcess -Run $run
            foreach ($p in $other, $child) { if ($p -and -not $p.HasExited) { $p.Kill() } }
        }
    }

    It 'keeps quoted arguments and reads both output streams' {
        $out = Join-Path $TestDrive 'o.txt'
        $run = Start-JobProcess -FilePath $script:Shell -ArgumentList @('-NoProfile', '-Command', (ConvertTo-NativeArgument -Value '[Console]::Out.WriteLine("a ""b"" c\"); [Console]::Error.WriteLine(''e'')')) -StandardOutput $out -StandardError "$out.err"
        try { $run.Process.WaitForExit(60000) | Should -BeTrue } finally { Stop-JobProcess -Run $run }
        (Read-RunOutput -Path $out).Trim() | Should -Be 'a "b" c\'
        (Read-RunOutput -Path "$out.err").Trim() | Should -Be 'e'
    }
}

Describe 'The child runners of the two scripts' {
    BeforeAll {
        # The functions are read from the scripts, so the test runs the shipped text and not the whole script.
        function Get-ScriptFunctionText([string]$Script, [string[]]$Name) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $Script), [ref]$null, [ref]$null)
            $found = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Name -contains $n.Name }, $true))
            if ($found.Count -ne $Name.Count) { throw "functions not found in $Script" }
            ($found | ForEach-Object { $_.Extent.Text }) -join "`n"
        }
        . ([scriptblock]::Create((Get-ScriptFunctionText 'Invoke-PackagerSmoke.ps1' 'Invoke-ChildProcessBatch')))
        . ([scriptblock]::Create((Get-ScriptFunctionText 'Invoke-FullRegression.ps1' 'Invoke-RegressionProcess')))
        $script:RepoRoot = Split-Path -Parent $PSScriptRoot

        function New-TempCount { @(Get-ChildItem ([IO.Path]::GetTempPath()) -File | Where-Object { $_.Name -match '^ap-(smoke|run)-[0-9a-f]{32}\.out' }).Count }
    }

    It 'the smoke batch returns each result and ends the child of a child that timed out' {
        $before = New-TempCount
        $childFile = Join-Path $TestDrive 'batch-child.txt'
        $hang = Join-Path $TestDrive 'batch-hang.ps1'
        Set-Content -LiteralPath $hang -Encoding Ascii -Value @(
            "`$c = Start-Process '$($script:Shell)' -ArgumentList '$($script:Sleep)' -PassThru -WindowStyle Hidden"
            "Set-Content -LiteralPath '$childFile' -Value `$c.Id"
            'Start-Sleep -Seconds 120')
        $jobs = @(
            [pscustomobject]@{ Key = 'quick'; FileName = $script:Shell; ArgumentList = @('-NoProfile', '-Command', '[Console]::Out.WriteLine(''out''); [Console]::Error.WriteLine(''err''); exit 3'); TimeoutSec = 60 }
            [pscustomobject]@{ Key = 'hang'; FileName = $script:Shell; ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $hang); TimeoutSec = 20 }
        )
        $r = Invoke-ChildProcessBatch -Jobs $jobs -ThrottleLimit 2
        $r['quick'].ExitCode | Should -Be 3
        $r['quick'].TimedOut | Should -BeFalse
        $r['quick'].StdOut.Trim() | Should -Be 'out'
        $r['quick'].StdErr.Trim() | Should -Be 'err'
        $r['hang'].TimedOut | Should -BeTrue
        Get-Process -Id ([int](Wait-File $childFile 5)) -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        New-TempCount | Should -Be $before
    }

    It 'the regression runner returns the exit code and both streams, and reports a timeout' {
        $before = New-TempCount
        $ok = Invoke-RegressionProcess -FileName $script:Shell -ArgumentList @('-NoProfile', '-Command', '[Console]::Out.WriteLine((Get-Location).Path); [Console]::Error.WriteLine(''e''); exit 4') -TimeoutSec 60
        $ok.ExitCode | Should -Be 4
        $ok.TimedOut | Should -BeFalse
        $ok.StdOut.Trim() | Should -Be $script:RepoRoot
        $ok.StdErr.Trim() | Should -Be 'e'
        $slow = Invoke-RegressionProcess -FileName $script:Shell -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 120') -TimeoutSec 5
        $slow.TimedOut | Should -BeTrue
        $slow.ExitCode | Should -Be -1
        New-TempCount | Should -Be $before
    }
}

Describe 'The runner scripts have no other way to stop a process' {
    It '<_> uses the job helper' -ForEach @('Invoke-FullRegression.ps1', 'Invoke-PackagerSmoke.ps1') {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot $_) -Raw
        $text | Should -Match 'ProcessJob\.ps1'
        $text | Should -Match 'Start-JobProcess'
        $text | Should -Not -Match '(?i)taskkill|ParentProcessId|Win32_Process|\bStop-Process\b|\.Kill\('
    }
}
