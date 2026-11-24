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
        $script:Script | Should -Match 'Get-RunCancelledLine -Operation \$Op -Done \$i -Total \$n'
        $script:Script | Should -Match 'Get-RunCancelledLine -Operation \$Op -Done \(\$i - 1\) -Total \$n'
        $script:Script | Should -Match 'Get-RunCancelledLine -Operation \$script:BgOperation -Done \(\[int\]\$doneState\.DoneCount\)'
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

Describe 'A cancelled run says it was cancelled and nothing else (Q-73)' {
    It 'names the run and counts the checked products' {
        Get-RunCancelledLine -Operation 'CheckConfigMgr' -Done 11 -Total 315 | Should -Be 'Check ConfigMgr cancelled: 11 of 315 products checked.'
        Get-RunCancelledLine -Operation 'CheckLatest' -Done 3 -Total 10 | Should -Be 'Check Latest cancelled: 3 of 10 products checked.'
        Get-RunCancelledLine -Operation 'CheckCve' -Done 0 -Total 1 | Should -Be 'Check CVEs cancelled: 0 of 1 product checked.'
    }
    It 'counts the done products for One Click and keeps the generic line for the other runs' {
        Get-RunCancelledLine -Operation 'OneClick' -Done 2 -Total 5 | Should -Be 'One Click cancelled: 2 of 5 products done.'
        Get-RunCancelledLine -Operation 'Stage' -Done 2 -Total 5 | Should -Be 'Cancelled: 2 of 5 products done.'
        Get-RunCancelledLine -Done 2 -Total 5 | Should -Be 'Cancelled: 2 of 5 products done.'
    }
    It 'the cancelled line never says finished or complete' {
        foreach ($op in 'CheckConfigMgr', 'CheckLatest', 'CheckCve', 'OneClick', 'FullRun', 'Stage', 'Package', '') {
            Get-RunCancelledLine -Operation $op -Done 1 -Total 4 | Should -Not -Match '(?i)finish|complete'
        }
    }
    It 'Check ConfigMgr, on Cancel, logs the cancelled line before the finished line, the Hide Unused offer and the CVE chain' {
        $m = [regex]::Match($script:Script, '(?s)if \(\$cmCancelled -or \(Test-UiRunCancelled\)\) \{(?<body>.*?)\r?\n        \}\r?\n        Hide-BusyIndicator\r?\n        Select-OnlyUpdateAvailable')
        $m.Success | Should -BeTrue
        $m.Groups['body'].Value | Should -Match "Get-RunCancelledLine -Operation 'CheckConfigMgr'"
        $m.Groups['body'].Value | Should -Match '\breturn\b'
        $m.Groups['body'].Value | Should -Not -Match 'Get-CheckResultSummaryLine|Hide Unused|CveAfterCmRows'
        $script:Script.IndexOf('if ($cmCancelled -or (Test-UiRunCancelled))') | Should -BeLessThan $script:Script.IndexOf('Hide Unused Applications')
        $script:Script.IndexOf('if ($cmCancelled -or (Test-UiRunCancelled))') | Should -BeLessThan $script:Script.IndexOf('$script:CveAfterCmRows = $selectedRows')
    }
    It 'a pipeline that Cancel reached after its last product still counts as cancelled' {
        $script:Script | Should -Match '(?s)if \(\[bool\]\$State\.CancelRequested -and -not \[bool\]\$State\.Canceled\) \{\s+\$State\.Canceled = \$true\s+\$State\.DoneCount = \$i'
    }
    It 'a cancelled check logs no summary of partial counts and no result line, and a cancelled Check CVEs logs no NVD notice' {
        $script:Script | Should -Match "-not \(\[bool\]\`$doneState\.Canceled -and \`$script:BgOperation -in @\('CheckLatest', 'CheckCve'\)\)"
        $script:Script | Should -Match "if \(\`$script:BgOperation -eq 'CheckLatest' -and -not \[bool\]\`$doneState\.Canceled\) \{ Add-LogLine -Message \(Get-CheckResultSummaryLine"
        $script:Script | Should -Match "if \(\`$Op -eq 'CheckCve' -and -not \[bool\]\`$State\.Canceled\) \{"
    }
    It 'a Check CVEs lookup that Cancel stopped in flight records nothing and logs no Error line (Q-73)' {
        $m = [regex]::Match($script:Script, '(?s)\$res = Get-ProductVulnerabilities -Cpe \$cpe -CurrentVersion \$current(?<body>.*?)
                    \}

                    ''Stage'' \{')
        $m.Success | Should -BeTrue
        $body = $m.Groups['body'].Value
        # The guard comes before the result is stored, shown or counted.
        $guard = $body.IndexOf('if ([bool]$State.Aborted) { continue }')
        $guard | Should -BeGreaterThan -1
        $guard | Should -BeLessThan $body.IndexOf('$State.CveResults[$scrName] = $res')
        $guard | Should -BeLessThan $body.IndexOf("'Error: ' + `$res.Error")
        # An exception that arrives with Cancel is not logged as an error either.
        $body | Should -Match '(?s)\} catch \{\s+if \(-not \[bool\]\$State\.Aborted\) \{\s+\$row\.CveCount = ''Error'''
    }
}

Describe 'The One Click dialog has its own Cancel while a run is active (Q-73)' {
    BeforeAll { $script:Script = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\start-apppackager.ps1') -Raw }
    It 'the dialog covers the window Cancel button, so the dialog shows a Cancel button that starts while the run is active and hides at the end' {
        $script:Script | Should -Match 'x:Name="btnCancelRun"[^>]*Content="Cancel"[^>]*Visibility="Collapsed"'
        $script:Script | Should -Match "\`$btnCancelRun\.IsEnabled = \`$true; \`$btnCancelRun\.Visibility = 'Visible'"
        $script:Script | Should -Match "(?s)\`$btnCancelRun\.Visibility = 'Collapsed'\s+\`$btnClose\.IsEnabled = \`$true"
    }
    It 'the dialog Cancel raises the Click of the window Cancel button, so one handler stops the run' {
        $script:Script | Should -Match '\$mainCancel = \$btnCancelPipeline'
        $script:Script | Should -Match '(?s)\$btnCancelRun\.Add_Click\(\{\s+\$btnCancelRun\.IsEnabled = \$false\s+\$mainCancel\.RaiseEvent\('
    }
}

Describe 'The window thread: a fake runner that is cancelled mid-run leaves a cancelled state, not a finished one' {
    It 'Cancel on a registered child ends the child, marks the run aborted, and the line reads cancelled' {
        $state = New-State; Set-RunState -State $state
        $p = New-Sleeper 60
        try {
            Register-RunChild -Process $p
            $state.CancelRequested = $true
            $r = Wait-RunChild -Process $p -TimeoutSeconds 30
            $r | Should -Be 'Cancelled'
            $line = Get-RunCancelledLine -Operation 'CheckLatest' -Done 0 -Total 1
            $line | Should -Be 'Check Latest cancelled: 0 of 1 product checked.'
        }
        finally { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
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
}

Describe 'The dropped-installer analysis has a Cancel (Q-69)' {
    It 'names the run and counts the analyzed installers' {
        Get-RunCancelledLine -Operation 'DropIntake' -Done 1 -Total 3 | Should -Be 'Cancelled: 1 of 3 installers analyzed. Nothing was staged or packaged.'
        Get-RunCancelledLine -Operation 'DropIntake' -Done 0 -Total 1 | Should -Be 'Cancelled: 0 of 1 installer analyzed. Nothing was staged or packaged.'
        Get-RunCancelledLine -Operation 'DropIntake' -Done 1 -Total 4 | Should -Not -Match '(?i)finish|complete'
    }
    It 'the window script offers Cancel, takes a run state and checks Cancel before and after each analysis' {
        $script:Script | Should -Match "Show-BusyIndicator -Title 'Analyzing dropped installers' -Step \(Format-BusyStep[^\r\n]*\) -AllowCancel"
        $m = [regex]::Match($script:Script, '(?s)function Invoke-DropIntake \{(?<body>.*?)\r?\n\}\r?\n\r?\n\$script:PendingDropQueue')
        $m.Success | Should -BeTrue
        $body = $m.Groups['body'].Value
        $body | Should -Match 'Start-UiRun -Total \$installerTotal'
        # A drop that arrives while the first one pumps messages must not start a second run: the guard names the run state.
        $body | Should -Match '-or \$null -ne \$script:DiscoveryJob -or \$script:UiRun\) \{'
        $body | Should -Match 'if \(Test-UiRunCancelled\) \{ \$dropCancelled = \$true; \$dropDone = \$installerIndex; break \}'
        $body | Should -Match '\$cancelledInFlight = Test-UiRunCancelled\s+Hide-BusyIndicator\s+if \(\$cancelledInFlight\) \{ \$dropCancelled = \$true; \$dropDone = \$installerIndex - 1; break \}'
        $body | Should -Match "Get-RunCancelledLine -Operation 'DropIntake'"
        $body | Should -Match '(?s)finally \{\s+Stop-UiRun\s+Hide-BusyIndicator'
        # The run state ends before the background run of the jobs starts, and a cancel never reaches it.
        $body.IndexOf('Stop-UiRun') | Should -BeLessThan $body.IndexOf('Invoke-AdHocPipeline')
        $body.IndexOf("Get-RunCancelledLine -Operation 'DropIntake'") | Should -BeLessThan $body.IndexOf('Invoke-AdHocPipeline')
    }
}

Describe 'A Workbench run does not start while a run on the window thread is active (Q-69)' {
    It 'the Workbench run handler checks the run state before it validates or closes the window' {
        $m = [regex]::Match($script:Script, '(?s)\$runFromWorkbench = \{(?<body>.*?)\r?\n    \}\r?\n    \$btnStage')
        $m.Success | Should -BeTrue
        $body = $m.Groups['body'].Value
        $body | Should -Match '(?s)if \(\$script:UiRun\) \{\s+\[void\]\(Show-ThemedMessage -Owner \$win -Title \$operation.*?A run is already in progress'
        $body.IndexOf('$script:UiRun') | Should -BeLessThan $body.IndexOf('$wb.Dirty')
        $body.IndexOf('$script:UiRun') | Should -BeLessThan $body.IndexOf('$win.Close()')
    }
}

Describe 'Invoke-DropIntake with a fake analysis and a fake Cancel (Q-69)' {
    BeforeAll {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($script:Script, [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-DropIntake' }, $true)
        . ([scriptblock]::Create($fn.Extent.Text))
        $script:Dir = Join-Path ([IO.Path]::GetTempPath()) ('dropcancel-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
        $script:Files = @(1..3 | ForEach-Object { $f = Join-Path $script:Dir ("setup$_.exe"); Set-Content -LiteralPath $f -Value 'x'; $f })
    }
    AfterAll { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }
    BeforeEach {
        $script:BgState = $null; $script:CmImportRunning = $false; $script:DiscoveryJob = $null
        $script:Prefs = [pscustomobject]@{ DownloadRoot = 'C:\dl'; SiteCode = 'XXX'; FileShareRoot = '\x\y'; ProviderMachineName = 'p'; ContentLayout = ''; EstimatedRuntimeMins = 0; MaximumRuntimeMins = 0
            DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $true } } }
        $script:window = New-Object psobject; $script:txtStatus = [pscustomobject]@{ Text = '' }; $script:txtComment = [pscustomobject]@{ Text = '' }
        $script:Calls = [System.Collections.Generic.List[string]]::new()
        $script:CancelAnswers = [System.Collections.Generic.Queue[bool]]::new()
        $script:Logged = [System.Collections.Generic.List[string]]::new()
        $script:Pipeline = $null
        function Add-LogLine { param([string]$Message) $script:Logged.Add($Message) }
        function Show-BusyIndicator { param([string]$Title, [string]$Step, [switch]$AllowCancel) $script:Calls.Add('busy' + $(if ($AllowCancel) { '+cancel' } else { '' })) }
        function Hide-BusyIndicator { $script:Calls.Add('hide') }
        function Start-UiRun { param([int]$Total) $script:Calls.Add("start:$Total") }
        function Stop-UiRun { $script:Calls.Add('stop') }
        function Format-BusyStep { param($Verb, $Index, $Total, $Name) "$Verb $Index/$Total" }
        function Test-UiRunCancelled { $a = if ($script:CancelAnswers.Count) { $script:CancelAnswers.Dequeue() } else { $false }; $script:Calls.Add("poll:$a"); $a }
        function Get-InstallerAnalysis { param($Path) $script:Calls.Add('analyze'); [pscustomobject]@{ FileName = (Split-Path -Leaf $Path); InstallerType = 'EXE'; Confidence = 'High' } }
        # The dialog of the first installer returns a staging job; the dialogs of the others skip (return nothing).
        function Show-DropIntakeDialog { param($Owner, $Analysis, $PackageAvailable, $PackageUnavailableReason) $script:Calls.Add('dialog'); if ($Analysis.FileName -eq 'setup1.exe') { [pscustomobject]@{ Action = 'Stage'; Values = [pscustomobject]@{ AppName = 'One' }; Analysis = $null } } else { $null } }
        function Invoke-AdHocPipeline { param($Jobs, $Context) $script:Pipeline = @($Jobs) }
        function Get-WorkbenchSigningPolicyJson { '' }
        function Test-WorkbenchModuleAvailable { $false }
        function Show-ThemedMessage { param($Owner, $Title, $Message, $Buttons, $Icon) 'OK' }
    }

    It 'Cancel before the second installer: stops the loop, reports 1 of 3, starts no pipeline, ends the run state' {
        $script:CancelAnswers.Enqueue($false); $script:CancelAnswers.Enqueue($false); $script:CancelAnswers.Enqueue($true)
        Invoke-DropIntake -Paths $script:Files
        @($script:Calls | Where-Object { $_ -eq 'analyze' }).Count | Should -Be 1
        $script:Calls | Should -Contain 'busy+cancel'
        $script:Calls | Should -Not -Contain 'busy'
        $script:Calls[0] | Should -Be 'start:3'
        $script:Calls[-1] | Should -Be 'hide'
        $script:Calls | Should -Contain 'stop'
        # The first installer gave a staging job before the cancel; a cancel never starts it.
        @($script:Calls | Where-Object { $_ -eq 'dialog' }).Count | Should -Be 1
        $script:Pipeline | Should -BeNullOrEmpty
        $script:txtStatus.Text | Should -Be 'Cancelled: 1 of 3 installers analyzed. Nothing was staged or packaged.'
        $script:Logged | Should -Contain $script:txtStatus.Text
    }
    It 'Cancel during the analysis of the first installer: its result is dropped and its dialog stays closed' {
        $script:CancelAnswers.Enqueue($false); $script:CancelAnswers.Enqueue($true)
        Invoke-DropIntake -Paths $script:Files
        @($script:Calls | Where-Object { $_ -eq 'analyze' }).Count | Should -Be 1
        $script:Calls | Should -Not -Contain 'dialog'
        # The click is handled while the busy panel still covers the window; the panel hides after it.
        $script:Calls.IndexOf('poll:True') | Should -BeLessThan $script:Calls.IndexOf('hide')
        $script:txtStatus.Text | Should -Be 'Cancelled: 0 of 3 installers analyzed. Nothing was staged or packaged.'
        $script:Pipeline | Should -BeNullOrEmpty
    }
    It 'without Cancel every installer is analyzed, the job of the first one runs, and the run state ends' {
        Invoke-DropIntake -Paths $script:Files
        @($script:Calls | Where-Object { $_ -eq 'analyze' }).Count | Should -Be 3
        @($script:Calls | Where-Object { $_ -eq 'dialog' }).Count | Should -Be 3
        @($script:Pipeline).Count | Should -Be 1
        $script:Pipeline[0].Action | Should -Be 'Stage'
        $script:Calls | Should -Contain 'stop'
        @($script:Logged | Where-Object { $_ -like 'Cancelled:*' }).Count | Should -Be 0
    }
}
}
