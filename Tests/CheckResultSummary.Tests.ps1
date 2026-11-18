#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The log line that closes Check Latest and Check ConfigMgr (Get-CheckResultSummaryLine).

.DESCRIPTION
    A fake grid list. A product needs an update when its status starts with 'Update available'; this is the rule of
    Select-OnlyUpdateAvailable. A product whose check failed has the status 'Error' and is counted as failed, not as eligible.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function New-FakeRows {
        param([hashtable]$Counts)
        $rows = @()
        foreach ($status in $Counts.Keys) { 1..$Counts[$status] | ForEach-Object { $rows += [pscustomobject]@{ Application = "App $status $_"; Status = $status } } }
        return $rows
    }
}

Describe 'Get-CheckResultSummaryLine' {
    It 'states the eligible and the selected products of a ConfigMgr check' {
        $rows = New-FakeRows @{ 'Update available' = 12; 'Up to date' = 300; 'Not found in ConfigMgr' = 6 }
        Get-CheckResultSummaryLine -Check 'Check ConfigMgr' -Rows $rows -SelectedCount 12 |
            Should -Be 'Check ConfigMgr finished: 12 of 318 products need an update; 12 selected.'
    }

    It 'counts a product with CVE text after the status as eligible, like the selection does' {
        $rows = New-FakeRows @{ 'Update available (3 CVEs, 1 in KEV)' = 2; 'Up to date' = 1 }
        Get-CheckResultSummaryLine -Check 'Check ConfigMgr' -Rows $rows -SelectedCount 2 |
            Should -Be 'Check ConfigMgr finished: 2 of 3 products need an update; 2 selected.'
    }

    It 'gives a line for the zero case' {
        $rows = New-FakeRows @{ 'Up to date' = 5 }
        Get-CheckResultSummaryLine -Check 'Check ConfigMgr' -Rows $rows -SelectedCount 0 |
            Should -Be 'Check ConfigMgr finished: 0 of 5 products need an update; 0 selected.'
    }

    It 'gives a line when nothing was checked' {
        Get-CheckResultSummaryLine -Check 'Check ConfigMgr' -Rows @() -SelectedCount 0 |
            Should -Be 'Check ConfigMgr finished: 0 of 0 products need an update; 0 selected.'
    }

    It 'names the failed checks and does not count them as eligible' {
        $rows = New-FakeRows @{ 'Update available' = 3; 'Error' = 2; 'Up to date' = 5 }
        Get-CheckResultSummaryLine -Check 'Check ConfigMgr' -Rows $rows -SelectedCount 3 |
            Should -Be 'Check ConfigMgr finished: 3 of 10 products need an update; 3 selected; 2 checks failed.'
    }

    It 'uses the singular for one product and one failed check' {
        $rows = New-FakeRows @{ 'Error' = 1 }
        Get-CheckResultSummaryLine -Check 'Check Latest' -Rows $rows |
            Should -Be 'Check Latest finished: 0 of 1 product needs an update; 1 check failed.'
        $one = New-FakeRows @{ 'Update available' = 1 }
        Get-CheckResultSummaryLine -Check 'Check Latest' -Rows $one |
            Should -Be 'Check Latest finished: 1 of 1 product needs an update.'
    }

    It 'leaves out the selection when the check selects nothing' {
        $rows = New-FakeRows @{ 'Update available' = 4; 'Up to date' = 6 }
        Get-CheckResultSummaryLine -Check 'Check Latest' -Rows $rows |
            Should -Be 'Check Latest finished: 4 of 10 products need an update.'
    }

    It 'is called after both checks in the start script' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
        $text | Should -Match "Get-CheckResultSummaryLine -Check 'Check Latest'"
        $text | Should -Match "Get-CheckResultSummaryLine -Check 'Check ConfigMgr'"
    }
}