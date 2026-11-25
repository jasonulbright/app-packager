#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every long action of the main window shows the shared busy overlay and updates the grid without a rebuild per row.

.DESCRIPTION
    Walks the click handler of every action button of the main window in start-apppackager.ps1 (by AST, through the
    functions that the handler calls) and checks, for each action:
      - it reaches the shared busy path: Invoke-MultiAppPipeline (the overlay with the spinner, the background timer, the
        per-row grid update) or Show-BusyIndicator (the same overlay for an action on the window thread);
      - an action on the window thread hides the overlay in a finally block, so an error and a cancel hide it too;
      - the background timer hides the overlay when the run is done, and a run on the window thread never calls
        Items.Refresh() on the grid inside a loop (one refresh per row rebuilds the icon column per row);
      - the overlay of the window holds a running progress ring and sits in the grid row, above the grid.
#>

$script:Actions = @(
    @{ Button = 'btnCheckLatest'; Path = 'pipeline' }
    @{ Button = 'btnCheckMECM'; Path = 'window thread' }
    @{ Button = 'btnCheckCve'; Path = 'pipeline' }
    @{ Button = 'btnStage'; Path = 'pipeline' }
    @{ Button = 'btnPackage'; Path = 'pipeline' }
    @{ Button = 'btnPublishIntune'; Path = 'pipeline' }
    @{ Button = 'btnPublishWsus'; Path = 'pipeline' }
    @{ Button = 'btnFullRun'; Path = 'pipeline' }
    @{ Button = 'btnDiscover'; Path = 'window thread' }
    @{ Button = 'btnConvertCm'; Path = 'window thread' }
    @{ Button = 'btnCveReport'; Path = 'window thread' }
)

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:StartPath = Join-Path $script:Root 'start-apppackager.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:StartPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $script:Functions = @{}
    foreach ($f in $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $script:Functions[$f.Name] = $f }
    $script:Xaml = Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw

    function Get-Handler([string]$Button) {
        $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Add_Click' -and $n.Expression.Extent.Text -eq ('$' + $Button) }, $true)
    }
    # The handler and every function that it reaches (up to five calls deep), as AST nodes.
    function Get-ReachableNodes($Root) {
        $seen = @{}
        $nodes = New-Object System.Collections.Generic.List[object]
        $queue = New-Object System.Collections.Generic.Queue[object]
        $queue.Enqueue(@($Root, 0))
        while ($queue.Count -gt 0) {
            $item = $queue.Dequeue(); $node = $item[0]; $depth = $item[1]
            $nodes.Add($node)
            if ($depth -ge 5) { continue }
            foreach ($c in $node.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                $name = $c.GetCommandName()
                if ($name -and $script:Functions.ContainsKey($name) -and -not $seen.ContainsKey($name)) { $seen[$name] = $true; $queue.Enqueue(@($script:Functions[$name], ($depth + 1))) }
            }
        }
        , $nodes
    }
    function Test-InLoop($Node, $Stop) {
        $p = $Node.Parent
        while ($p -and $p -ne $Stop) {
            if ($p -is [System.Management.Automation.Language.LoopStatementAst]) { return $true }
            $p = $p.Parent
        }
        $false
    }
}

Describe 'The actions of the main window use the shared busy path' {
    It 'finds the click handler of <Button>' -ForEach $script:Actions {
        Get-Handler $Button | Should -Not -BeNullOrEmpty
    }

    It '<Button> runs on the <Path> path with the shared overlay' -ForEach $script:Actions {
        $nodes = Get-ReachableNodes (Get-Handler $Button)
        $text = ($nodes | ForEach-Object { $_.Extent.Text }) -join "`n"
        if ($Path -eq 'pipeline') { $text | Should -Match 'Invoke-MultiAppPipeline' }
        else { $text | Should -Match 'Show-BusyIndicator' }
    }

    It '<Button> never calls Items.Refresh() on the grid inside a loop' -ForEach $script:Actions {
        $nodes = Get-ReachableNodes (Get-Handler $Button)
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($node in $nodes) {
            foreach ($call in $node.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Refresh' -and $n.Expression.Extent.Text -match 'Items$' }, $true)) {
                if (Test-InLoop $call $node) { $bad.Add($call.Extent.StartLineNumber) }
            }
        }
        ($bad -join ', ') | Should -BeNullOrEmpty
    }

    It 'every function that shows the busy overlay on the window thread hides it in a finally block' {
        # Start-DiscoveryScan hands the work to a background runspace and returns; its timer ends the run (next test).
        $shown = $script:Functions.Values | Where-Object { $_.Extent.Text -match 'Show-BusyIndicator' -and $_.Name -notin 'Show-BusyIndicator', 'Start-DiscoveryScan' }
        $shown | Should -Not -BeNullOrEmpty
        foreach ($fn in $shown) {
            $tries = $fn.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Finally -and $n.Finally.Extent.Text -match 'Hide-BusyIndicator' }, $true)
            @($tries).Count | Should -BeGreaterThan 0 -Because "$($fn.Name) must hide the overlay on error and on cancel"
        }
    }

    It 'Discover hides the overlay when the scan ends, fails to start, fails in the timer, or is cancelled' {
        $start = $script:Functions['Start-DiscoveryScan']
        $start | Should -Not -BeNullOrEmpty
        # A start failure goes to Complete-DiscoveryScan; the timer ends through it and hides the overlay in its own catch block.
        $start.Extent.Text | Should -Match 'Complete-DiscoveryScan -Result'
        $tick = $start.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Add_Tick' }, $true) | Select-Object -First 1
        $tick | Should -Not -BeNullOrEmpty
        $catch = $tick.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true) | Select-Object -First 1
        $catch.Extent.Text | Should -Match 'Hide-BusyIndicator'
        $script:Functions['Complete-DiscoveryScan'].Extent.Text | Should -Match 'Hide-BusyIndicator'
    }

    It 'the handlers that show the overlay on the window thread hide it in a finally block' {
        foreach ($button in 'btnCheckMECM', 'btnConvertCm', 'btnCveReport') {
            $nodes = Get-ReachableNodes (Get-Handler $button)
            $withShow = @($nodes | Where-Object { $_.Extent.Text -match 'Show-BusyIndicator' })
            $withShow.Count | Should -BeGreaterThan 0
            $finals = 0
            foreach ($n in $nodes) { $finals += @($n.FindAll({ param($x) $x -is [System.Management.Automation.Language.TryStatementAst] -and $x.Finally -and $x.Finally.Extent.Text -match 'Hide-BusyIndicator' }, $true)).Count }
            $finals | Should -BeGreaterThan 0 -Because "$button must hide the overlay in a finally block"
        }
    }

    It 'Invoke-MultiAppPipeline shows the overlay and its timer hides it when the run is done, on success, error and cancel' {
        $fn = $script:Functions['Invoke-MultiAppPipeline']
        $fn | Should -Not -BeNullOrEmpty
        $fn.Extent.Text | Should -Match '\$progressOverlay\.Visibility = \[System\.Windows\.Visibility\]::Visible'
        $tick = $fn.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Add_Tick' }, $true) | Select-Object -First 1
        $tick | Should -Not -BeNullOrEmpty
        $done = $tick.Extent.Text.IndexOf('BgState.Done')
        $done | Should -BeGreaterThan 0
        $tick.Extent.Text.Substring($done) | Should -Match '\$progressOverlay\.Visibility = \[System\.Windows\.Visibility\]::Collapsed'
        # The timer finishes the grid at the end of the run and refreshes only through the per-row helper.
        $tick.Extent.Text | Should -Match 'Update-GridWhenChanged'
        $tick.Extent.Text | Should -Match 'Complete-GridUpdates'
        $tick.Extent.Text | Should -Not -Match '\$dataGrid\.Items\.Refresh\(\)'
    }

    It 'the overlay is drawn on its own, above the grid, with a running progress ring' {
        $script:Xaml | Should -Match '<Border x:Name="progressOverlay"[^>]*Grid\.Row="1"'
        $script:Xaml | Should -Match '<Controls:ProgressRing IsActive="True"'
    }
}
