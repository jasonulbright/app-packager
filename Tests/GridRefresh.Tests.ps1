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
    foreach ($name in 'Get-GridSignature', 'Update-GridWhenChanged', 'Get-GridIconPlaceholder', 'Get-GridIconIndex', 'Get-GridIconSource', 'Reset-GridIconCache') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function New-StandInGrid {
        $items = [pscustomobject]@{ Refreshes = 0 }
        $items | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $this.Refreshes++ }
        [pscustomobject]@{ Items = $items }
    }
}

Describe 'Update-GridWhenChanged' {
    BeforeEach {
        $script:GridSignature = $null
        $script:rows = @(
            [pscustomobject]@{ Selected = $false; Application = 'Git'; CurrentVersion = '2.1'; LatestVersion = '2.2'; Status = 'Ready'; CveCount = ''; CveCvss = ''; CveKev = ''; CveFixed = ''; CveNewest = ''; CveEpss = ''; CveUpdated = ''; CveChecked = ''; LastChecked = '' }
            [pscustomobject]@{ Selected = $true; Application = 'Putty'; CurrentVersion = ''; LatestVersion = ''; Status = 'Ready'; CveCount = ''; CveCvss = ''; CveKev = ''; CveFixed = ''; CveNewest = ''; CveEpss = ''; CveUpdated = ''; CveChecked = ''; LastChecked = '' }
        )
        $script:grid = New-StandInGrid
    }

    It 'refreshes once, then not again while 300 busy ticks change nothing' {
        1..300 | ForEach-Object { [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows) }
        $script:grid.Items.Refreshes | Should -Be 1
    }

    It 'refreshes again when a shown value changes' {
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        $script:rows[0].Status = 'Staging...'
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeTrue
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeFalse
        $script:grid.Items.Refreshes | Should -Be 2
    }

    It 'refreshes when a row is added or the selection changes' {
        [void](Update-GridWhenChanged -Grid $script:grid -Rows $script:rows)
        $script:rows[0].Selected = $true
        Update-GridWhenChanged -Grid $script:grid -Rows $script:rows | Should -BeTrue
        Update-GridWhenChanged -Grid $script:grid -Rows @($script:rows + $script:rows[0]) | Should -BeTrue
    }

    It 'accepts an empty grid' {
        { Update-GridWhenChanged -Grid $script:grid -Rows @() } | Should -Not -Throw
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
