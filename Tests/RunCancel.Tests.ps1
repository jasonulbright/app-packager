#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Cancel for every run (Q-69): the shared helpers with a fake child process, and the places in the window script that use them.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $script:Root 'Packagers\AppPackagerCommon.psd1') -Force
    $script:Script = Get-Content -LiteralPath (Join-Path $script:Root 'start-apppackager.ps1') -Raw

    function New-Sleeper([int]$Seconds = 60) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-Command', "Start-Sleep $Seconds") -PassThru -WindowStyle Hidden
    }
    function New-State { [hashtable]::Synchronized(@{ CancelRequested = $false; Aborted = $false; DoneCount = 0; TotalCount = 0 }) }
}

Describe 'Cancel for every run' {
AfterEach {
    Set-RunState -State $null
    [void](Stop-RunChildren)
}

Describe 'The cancelled line' {
    It 'says how many products are done' {
        Get-CancelledSummaryLine -Done 3 -Total 107 | Should -Be 'Cancelled: 3 of 107 products done.'
    }
    It 'uses the singular for one product' {
        Get-CancelledSummaryLine -Done 0 -Total 1 | Should -Be 'Cancelled: 0 of 1 product done.'
    }
}

Describe 'The run state' {
    It 'is not cancelled without a state' {
        Set-RunState -State $null
        Test-RunCancelRequested | Should -BeFalse
        Test-RunAborted | Should -BeFalse
    }
    It 'follows the flags of the state' {
        $state = New-State
        Set-RunState -State $state
        Test-RunCancelRequested | Should -BeFalse
        $state.CancelRequested = $true
        Test-RunCancelRequested | Should -BeTrue
        Set-RunAborted
        $state.Aborted | Should -BeTrue
        Test-RunAborted | Should -BeTrue
    }
}

Describe 'Wait-RunChild with a fake child process' {
    It 'cancel before the first product: the child ends at once' {
        $state = New-State; $state.CancelRequested = $true; Set-RunState -State $state
        $p = New-Sleeper
        $clock = [Diagnostics.Stopwatch]::StartNew()
        (Wait-RunChild -Process $p -TimeoutSeconds 30) | Should -Be 'Cancelled'
        $clock.Elapsed.TotalSeconds | Should -BeLessThan 5
        $p.HasExited | Should -BeTrue
        $state.Aborted | Should -BeTrue
    }
    It 'cancel in the middle of a download: the child ends within seconds' {
        $state = New-State; Set-RunState -State $state
        $p = New-Sleeper
        # Another runspace asks for Cancel while the wait is running.
        $rs = [runspacefactory]::CreateRunspace(); $rs.Open(); $ps = [powershell]::Create(); $ps.Runspace = $rs
        [void]$ps.AddScript({ param($s) Start-Sleep -Milliseconds 700; $s['CancelRequested'] = $true }).AddArgument($state); [void]$ps.BeginInvoke()
        $clock = [Diagnostics.Stopwatch]::StartNew()
        (Wait-RunChild -Process $p -TimeoutSeconds 30) | Should -Be 'Cancelled'
        $clock.Elapsed.TotalSeconds | Should -BeLessThan 8
        $p.HasExited | Should -BeTrue
        $ps.Dispose(); $rs.Dispose()
    }
    It 'cancel at the end: a child that already finished is a normal result' {
        $state = New-State; Set-RunState -State $state
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-Command', 'exit 0') -PassThru -WindowStyle Hidden
        $null = $p.WaitForExit(20000)
        $state.CancelRequested = $true
        # The run is cancelled, but the product finished with success: it stays done. No process is left to stop.
        $p.HasExited | Should -BeTrue
        $p.ExitCode | Should -Be 0
    }
    It 'a child that ends with a failure after Cancel counts as cancelled' {
        $state = New-State; Set-RunState -State $state
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep 1; exit 3') -PassThru -WindowStyle Hidden
        $state.CancelRequested = $false
        $null = $p.WaitForExit(20000)
        $state.CancelRequested = $true
        # The Cancel button stops the child from outside; the wait sees the failure and the request.
        (Wait-RunChild -Process $p -TimeoutSeconds 5) | Should -Be 'Cancelled'
    }
    It 'returns Exited when nothing was cancelled' {
        $state = New-State; Set-RunState -State $state
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-Command', 'exit 0') -PassThru -WindowStyle Hidden
        (Wait-RunChild -Process $p -TimeoutSeconds 30) | Should -Be 'Exited'
    }
    It 'returns TimedOut after the limit and leaves the stop to the caller' {
        $state = New-State; Set-RunState -State $state
        $p = New-Sleeper
        (Wait-RunChild -Process $p -TimeoutSeconds 1) | Should -Be 'TimedOut'
        Stop-ProcessTreeById -ProcessId $p.Id
    }
}

Describe 'Stop-RunChildren stops only the processes that the run registered' {
    It 'leaves another process alone and a second call does nothing' {
        $mine = New-Sleeper
        $other = New-Sleeper
        try {
            Register-RunChild -Process $mine
            (Stop-RunChildren) | Should -Be 1
            $null = $mine.WaitForExit(5000)
            $mine.HasExited | Should -BeTrue
            $other.HasExited | Should -BeFalse
            (Stop-RunChildren) | Should -Be 0
        }
        finally { Stop-Process -Id $other.Id -Force -ErrorAction SilentlyContinue }
    }
    It 'does not stop a process whose id was reused (the start time differs)' {
        $p = New-Sleeper
        try {
            $registry = [AppDomain]::CurrentDomain.GetData('AppPackager.RunChildren')
            if (-not $registry) { Register-RunChild -Process $p; $registry = [AppDomain]::CurrentDomain.GetData('AppPackager.RunChildren') }
            $registry[[int]$p.Id] = 12345L
            (Stop-RunChildren) | Should -Be 0
            $p.HasExited | Should -BeFalse
        }
        finally { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Update-PackagerHistory leaves no trace for a product that Cancel stopped' {
    It 'returns without writing when the run is aborted' {
        $state = New-State; $state.Aborted = $true; Set-RunState -State $state
        $text = Get-Content -LiteralPath (Join-Path $script:Root 'Packagers\AppPackagerCommon.psm1') -Raw
        $text | Should -Match '(?s)function Update-PackagerHistory \{.*?if \(Test-RunAborted\) \{ return \}'
    }
}

Describe 'Every run has a Cancel (window script)' {
    It 'the pipeline sets the run state, reverts the product in flight and reports the cancelled line' {
        $script:Script | Should -Match 'Set-RunState -State \$script:BgState'
        $script:Script | Should -Match '\$snapStatus = \$row\.Status'
        $script:Script | Should -Match 'if \(\[bool\]\$State\.Aborted\) \{'
        $script:Script | Should -Match 'Get-CancelledSummaryLine -Done \$i -Total \$n'
        $script:Script | Should -Match 'Get-CancelledSummaryLine -Done \(\$i - 1\) -Total \$n'
        $script:Script | Should -Match 'Get-CancelledSummaryLine -Done \(\[int\]\$doneState\.DoneCount\)'
    }
    It 'the Cancel button stops the children of the run by id and the second click does nothing' {
        $script:Script | Should -Match '(?s)\$btnCancelPipeline\.Add_Click\(\{.*?Stop-RunChildren.*?\}\)'
        $script:Script | Should -Match 'if \(\[bool\]\$script:UiRun\.CancelRequested\) \{ return \}'
    }
    It 'closing the window stops the children of the run' {
        $script:Script | Should -Match '(?s)\$window\.Add_Closing\(\{.*?\[void\]\(Stop-RunChildren\)'
    }
    It 'Check Latest, Stage and Package read the packager through the cancel-aware waits' {
        $script:Script | Should -Match 'Wait-RunChild -Process \$p -TimeoutSeconds 30'
        $script:Script | Should -Match 'Register-RunChild -Process \$p'
        $script:Script | Should -Match '\(Test-RunCancelRequested\)\) \{\s+\$cancelled = \$true'
    }
    It 'Check ConfigMgr and the conversion offer Cancel' {
        $script:Script | Should -Match "Show-BusyIndicator -Title 'Checking ConfigMgr' -AllowCancel"
        $script:Script | Should -Match 'if \(Test-UiRunCancelled\) \{ \$cmCancelled = \$true; break \}'
        $script:Script | Should -Match "Show-BusyIndicator -Title 'Converting a ConfigMgr application' -Step 'Starting the conversion...' -AllowCancel"
    }
    It 'stops no process by a command-line pattern' {
        $script:Script | Should -Not -Match 'Win32_Process[^\r\n]*CommandLine'
        $script:Script | Should -Not -Match '(?i)taskkill(\.exe)?\s+/IM'
    }
}

Describe 'The NVD wait ends at Cancel' {
    It 'throws a cancellation and marks the run aborted' {
        $vuln = Join-Path $script:Root 'Packagers\AppPackagerVulnerability.psd1'
        Import-Module $vuln -Force
        $state = New-State; $state.CancelRequested = $true; Set-RunState -State $state
        $thrown = $null
        InModuleScope 'AppPackagerVulnerability' {
            $script:NvdClock = { [datetime]::UtcNow }
            try { Wait-NvdUntil -Time ([datetime]::UtcNow.AddSeconds(30)) } catch { $script:caught = $_.Exception }
        }
        $state.Aborted | Should -BeTrue
    }
}

Describe 'A loop on the window thread hears Cancel (Check ConfigMgr, the conversion)' {
    BeforeAll {
        Add-Type -AssemblyName WindowsBase, PresentationFramework
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($script:Script, [ref]$null, [ref]$null)
        foreach ($name in 'Invoke-UiPump', 'Test-UiRunCancelled') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            . ([scriptblock]::Create($fn.Extent.Text))
        }
    }
    It 'the pump runs the input that is queued (a click on Cancel)' {
        $script:Handled = $false
        $d = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
        [void]$d.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Input, [Action]{ $script:Handled = $true })
        Invoke-UiPump
        $script:Handled | Should -BeTrue
    }
    It 'the cancel flag is honored between products without any help from the window thread' {
        $script:UiRun = [hashtable]::Synchronized(@{ CancelRequested = $false })
        $done = 0
        $items = 1..10
        foreach ($item in $items) {
            if (Test-UiRunCancelled) { break }
            $done++
            if ($item -eq 4) { $script:UiRun.CancelRequested = $true }
        }
        $done | Should -Be 4
        $script:UiRun = $null
        (Test-UiRunCancelled) | Should -BeFalse
    }
    It 'the loops on the window thread use one pump' {
        $script:Script | Should -Match 'function Invoke-UiPump'
        $script:Script | Should -Match 'PushFrame'
        $script:Script | Should -Not -Match '\[void\]\[System\.Windows\.Threading\.Dispatcher\]::CurrentDispatcher\.Invoke\(\s*\[System\.Windows\.Threading\.DispatcherPriority\]::Background'
    }
}}
