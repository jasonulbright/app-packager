#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The busy timer does not rebuild the application grid on a tick that changed nothing.

.DESCRIPTION
    A rebuilt row shows its icon blank until the icon is set again, so a refresh on every
    100 ms busy tick made the icons blink and cost CPU. The helpers are lifted out of
    start-apppackager.ps1 by AST. A stand-in grid counts the refreshes. The icon cache
    returns the same frozen image on every call, so the decode count does not grow.
#>

BeforeAll {
    Add-Type -AssemblyName PresentationFramework
    $script:ScriptPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-GridRowSignature', 'Get-GridRowKey', 'Initialize-GridRowSignatures', 'Get-GridSignature', 'Update-GridRowCells', 'Update-GridWhenChanged', 'Complete-GridUpdates', 'Get-GridIconPlaceholder', 'Get-GridIconIndex', 'Get-GridIconSource', 'Reset-GridIconCache') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function New-StandInGrid {
        # A grid without a container generator: only Items.Refresh() can show a change.
        $items = [pscustomobject]@{ Refreshes = 0; SortDescriptions = @() }
        $items | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $this.Refreshes++ }
        [pscustomobject]@{ Items = $items }
    }
    function New-RealRowGrid {
        # A grid with row containers: one real visual tree per row (a Status TextBlock bound to the row, an icon Image
        # that is set once), a counter for Items.Refresh(), and a lookup like ItemContainerGenerator.ContainerFromItem.
        param([object[]]$Rows, [switch]$Sorted)
        $containers = @{}
        foreach ($row in $Rows) {
            $panel = New-Object System.Windows.Controls.StackPanel
            $text = New-Object System.Windows.Controls.TextBlock
            $text.DataContext = $row
            [void]$text.SetBinding([System.Windows.Controls.TextBlock]::TextProperty, (New-Object System.Windows.Data.Binding 'Status'))
            $image = New-Object System.Windows.Controls.Image
            $image.Source = Get-GridIconPlaceholder
            [void]$panel.Children.Add($image)
            [void]$panel.Children.Add($text)
            $containers[$row.Script] = [pscustomobject]@{ Panel = $panel; Text = $text; Image = $image; IconSource = $image.Source }
        }
        $items = [pscustomobject]@{ Refreshes = 0; SortDescriptions = $(if ($Sorted) { @('Status') } else { @() }) }
        $items | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $this.Refreshes++ }
        $generator = [pscustomobject]@{ Containers = $containers }
        $generator | Add-Member -MemberType ScriptMethod -Name ContainerFromItem -Value { param($item) $c = $this.Containers[[string]$item.Script]; if ($c) { $c.Panel } else { $null } }
        [pscustomobject]@{ Items = $items; ItemContainerGenerator = $generator; Containers = $containers }
    }
    function New-Row { param([string]$Script, [string]$Application) [pscustomobject]@{ Selected = $false; Script = $Script; Application = $Application; CurrentVersion = ''; LatestVersion = ''; Status = 'Ready'; CveCount = ''; CveCvss = ''; CveKev = ''; CveFixed = ''; CveNewest = ''; CveEpss = ''; CveUpdated = ''; CveChecked = ''; LastChecked = '' } }
}

Describe 'Update-GridWhenChanged on a grid without row containers' {
    BeforeEach {
        $script:GridRowSignatures = @{}
        $script:GridResortPending = $false
        $script:rows = @((New-Row 'package-git.ps1' 'Git'), (New-Row 'package-putty.ps1' 'Putty'))
        $script:grid = New-StandInGrid
    }

    It 'does nothing while 300 busy ticks change nothing' {
        1..300 | ForEach-Object { [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows) }
        $script:grid.Items.Refreshes | Should -Be 0
    }

    It 'falls back to one refresh per change when the grid cannot re-read a row' {
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        $script:rows[0].Status = 'Staging...'
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeTrue
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeFalse
        $script:grid.Items.Refreshes | Should -Be 1
    }

    It 'reports a change of the selection and accepts an empty grid' {
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        $script:rows[0].Selected = $true
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeTrue
        { Update-GridWhenChanged -Grid $script:grid -Rows @() } | Should -Not -Throw
    }
}

Describe 'A Check Latest run updates the cells of the checked row only' {
    BeforeEach {
        $script:GridRowSignatures = @{}
        $script:GridResortPending = $false
        $script:rows = @(1..40 | ForEach-Object { New-Row ("package-app$_.ps1") "App $_" })
        $script:grid = New-RealRowGrid -Rows $script:rows
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
    }

    It 'never refreshes the grid and never rebuilds a row or sets an icon again, over 40 checked applications' {
        foreach ($row in $script:rows) {
            $row.Status = 'Checking latest...'
            [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
            $row.LatestVersion = '1.0'
            $row.Status = 'Up to date'
            [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        }
        $script:grid.Items.Refreshes | Should -Be 0
        foreach ($row in $script:rows) {
            $c = $script:grid.Containers[$row.Script]
            $c.Text.Text | Should -Be 'Up to date'
            [object]::ReferenceEquals($c.Image.Source, $c.IconSource) | Should -BeTrue
        }
    }

    It 'shows the new status in the checked row and leaves the other rows as they were' {
        $script:rows[3].Status = 'Update available'
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeTrue
        $script:grid.Containers['package-app4.ps1'].Text.Text | Should -Be 'Update available'
        $script:grid.Containers['package-app5.ps1'].Text.Text | Should -Be 'Ready'
    }

    It 'skips a row that is not realized (scrolled out of the view) without an error' {
        $script:grid.Containers.Remove('package-app7.ps1')
        $script:rows[6].Status = 'Error'
        { Update-GridWhenChanged -Grid $script:grid -Rows $script:rows } | Should -Not -Throw
        $script:grid.Items.Refreshes | Should -Be 0
    }

    It 'refreshes a sorted grid once when the run ends, not per row' {
        $sorted = New-RealRowGrid -Rows $script:rows -Sorted
        $script:GridRowSignatures = @{}
        [void](Update-GridWhenChanged -Grid $sorted -Rows $script:rows)
        foreach ($row in $script:rows) { $row.Status = 'Up to date'; [void](Update-GridWhenChanged -Grid $sorted -Rows $script:rows) }
        $sorted.Items.Refreshes | Should -Be 0
        Complete-GridUpdates -Grid $sorted
        Complete-GridUpdates -Grid $sorted
        $sorted.Items.Refreshes | Should -Be 1
    }

    It 'does not refresh an unsorted grid at the end of the run' {
        $script:rows[0].Status = 'Up to date'
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        Complete-GridUpdates -Grid $script:grid
        $script:grid.Items.Refreshes | Should -Be 0
    }
}

Describe 'The grid icon cache' {
    It 'gives the same frozen image on every call, so a tick decodes nothing' {
        $script:GridIconCache = @{}
        $script:GridIconIndex = @{}
        $first = Get-GridIconSource -ScriptName 'package-nothing.ps1'
        $first.IsFrozen | Should -BeTrue
        1..300 | ForEach-Object { Get-GridIconSource -ScriptName 'package-nothing.ps1' } | ForEach-Object { [object]::ReferenceEquals($_, $first) | Should -BeTrue }
        $script:GridIconCache.Count | Should -Be 1
    }
}

Describe 'The busy timer' {
    It 'does not call Items.Refresh() directly on a tick' {
        $tick = $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Add_Tick' -and $n.Expression.Extent.Text -eq '$script:BgTimer' }, $true) | Select-Object -First 1
        $tick | Should -Not -BeNullOrEmpty
        $firstLines = ($tick.Extent.Text -split "`n" | Select-Object -First 70) -join "`n"
        $firstLines | Should -Match 'Update-GridWhenChanged'
    }
}
