#Requires -Modules Pester

<#
.SYNOPSIS
    Every long action shows the in-app progress display, not the wait cursor alone.

.DESCRIPTION
    The helpers are lifted out of start-apppackager.ps1 by AST and run against
    stand-ins for the overlay controls. A second test scans the script: a wait
    cursor may only be set where the same code also shows the in-app display.
#>

BeforeAll {
    Add-Type -AssemblyName PresentationFramework
    $script:ScriptPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Show-BusyIndicator', 'Set-BusyStep', 'Hide-BusyIndicator', 'Format-BusyStep', 'Update-UiNow') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
}

Describe 'Format-BusyStep' {
    It 'writes the verb, the position and the name in the format of the pipeline' {
        Format-BusyStep -Verb 'ConfigMgr' -Index 3 -Total 12 -Name 'Git' | Should -Be 'ConfigMgr 3/12: Git'
    }
}

Describe 'The shared busy indicator' {
    BeforeEach {
        $script:progressOverlay = [pscustomobject]@{ Visibility = [System.Windows.Visibility]::Collapsed }
        $script:txtProgressTitle = [pscustomobject]@{ Text = '' }
        $script:txtProgressStep = [pscustomobject]@{ Text = '' }
        $script:btnPausePipeline = [pscustomobject]@{ Visibility = [System.Windows.Visibility]::Visible; IsEnabled = $true; Content = 'Pause' }
        $script:btnCancelPipeline = [pscustomobject]@{ Visibility = [System.Windows.Visibility]::Visible; IsEnabled = $true }
        $script:window = [pscustomobject]@{ Cursor = $null }
        $progressOverlay = $script:progressOverlay; $txtProgressTitle = $script:txtProgressTitle; $txtProgressStep = $script:txtProgressStep
        $btnPausePipeline = $script:btnPausePipeline; $btnCancelPipeline = $script:btnCancelPipeline; $window = $script:window
    }

    It 'shows the overlay with a title and a status line, and the wait cursor as an addition' {
        Show-BusyIndicator -Title 'Checking ConfigMgr' -Step 'Starting...'
        $script:progressOverlay.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
        $script:txtProgressTitle.Text | Should -Be 'Checking ConfigMgr'
        $script:txtProgressStep.Text | Should -Be 'Starting...'
        $script:window.Cursor | Should -Be ([System.Windows.Input.Cursors]::Wait)
    }

    It 'hides Pause and Cancel for an action that has neither' {
        Show-BusyIndicator -Title 'Writing the CVE report'
        $script:btnPausePipeline.Visibility | Should -Be ([System.Windows.Visibility]::Collapsed)
        $script:btnCancelPipeline.Visibility | Should -Be ([System.Windows.Visibility]::Collapsed)
    }

    It 'keeps Cancel for an action that can be stopped' {
        Show-BusyIndicator -Title 'Discovering' -AllowCancel
        $script:btnPausePipeline.Visibility | Should -Be ([System.Windows.Visibility]::Collapsed)
        $script:btnCancelPipeline.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
        $script:btnCancelPipeline.IsEnabled | Should -BeTrue
    }

    It 'changes the status line while it runs' {
        Show-BusyIndicator -Title 'Checking ConfigMgr'
        Set-BusyStep -Text 'ConfigMgr 2/5: Git'
        $script:txtProgressStep.Text | Should -Be 'ConfigMgr 2/5: Git'
    }

    It 'restores the overlay, the cursor and both buttons when it ends' {
        Show-BusyIndicator -Title 'Discovering' -AllowCancel
        Hide-BusyIndicator
        $script:progressOverlay.Visibility | Should -Be ([System.Windows.Visibility]::Collapsed)
        $script:window.Cursor | Should -BeNullOrEmpty
        $script:btnPausePipeline.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
        $script:btnCancelPipeline.Visibility | Should -Be ([System.Windows.Visibility]::Visible)
        $script:btnPausePipeline.Content | Should -Be 'Pause'
    }

    It 'can be hidden twice' {
        Show-BusyIndicator -Title 'x'
        Hide-BusyIndicator
        { Hide-BusyIndicator } | Should -Not -Throw
    }
}

Describe 'No long action sets the wait cursor alone' {
    It 'shows the in-app display wherever the wait cursor is set' {
        $waitAssignments = $script:Ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left.Extent.Text -match '\.Cursor$' -and
                $n.Right.Extent.Text -match 'Cursors\]::Wait'
            }, $true)
        # The scan must find the sites, or it proves nothing.
        @($waitAssignments).Count | Should -BeGreaterThan 3
        $alone = New-Object System.Collections.Generic.List[string]
        foreach ($a in $waitAssignments) {
            # The helper itself sets the cursor as part of the display.
            $fn = $a.Parent
            while ($fn -and $fn -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
            if ($fn -and $fn.Name -in 'Show-BusyIndicator', 'Set-DialogBusyStatus') { continue }
            # An action shows the display through the helper, or through the
            # pipeline overlay it already used, or reports in the status line
            # of a dialog that the overlay cannot cover.
            # The scope is the nearest script block (an event handler) or, if
            # there is none, the enclosing function.
            $scope = $a.Parent
            while ($scope -and $scope -isnot [System.Management.Automation.Language.ScriptBlockExpressionAst] -and $scope -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $scope = $scope.Parent }
            if (-not $scope) { $scope = $a.Parent }
            $text = $scope.Extent.Text
            if ($text -notmatch 'Show-BusyIndicator|progressOverlay\.Visibility|Set-DialogBusyStatus') {
                $alone.Add(('line {0}: {1}' -f $a.Extent.StartLineNumber, $a.Extent.Text))
            }
        }
        $alone | Should -BeNullOrEmpty
    }
}

Describe 'A running discovery blocks other runs that use the display' {
    BeforeAll {
        foreach ($name in 'Test-PipelineBusy', 'Invoke-DropIntake') {
            $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
            if (-not $fn) { throw "function $name not found" }
            . ([scriptblock]::Create($fn.Extent.Text))
        }
        function Show-ThemedMessage { param($Owner, $Title, $Message, $Buttons, $Icon) $script:Shown += $Title; 'OK' }
        function Add-LogLine { param([string]$Message) $script:Lines += $Message }
    }
    BeforeEach {
        $script:BgState = $null; $script:CmImportRunning = $false; $script:DiscoveryJob = $null
        $script:Shown = @(); $script:Lines = @()
        $script:window = [pscustomobject]@{}
        $window = $script:window
    }

    It 'is not busy when nothing runs' {
        Test-PipelineBusy | Should -BeFalse
    }
    It 'is busy while the discovery job runs' {
        $script:DiscoveryJob = [pscustomobject]@{ Name = 'job' }
        Test-PipelineBusy | Should -BeTrue
    }
    It 'refuses an installer drop while the discovery job runs' {
        $script:DiscoveryJob = [pscustomobject]@{ Name = 'job' }
        Invoke-DropIntake -Paths @('C:\x\setup.exe')
        $script:Shown | Should -Be @('Pipeline Running')
        $script:Lines | Should -BeNullOrEmpty
    }
}
