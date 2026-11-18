#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    AppFlow.DefaultCadenceDays is a saved preference with a default of 7 and a range of 1 to 365.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted') {
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($node.Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Read-PrefsWithAppFlow {
        param([hashtable]$AppFlow)
        $script:PrefsFile = Join-Path $TestDrive ('prefs-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ AppFlow = [pscustomobject]$AppFlow } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        Read-Preferences
    }
}

Describe 'AppFlow.DefaultCadenceDays' {
    It 'defaults to 7 when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).AppFlow.DefaultCadenceDays | Should -Be 7
    }

    It 'reads <Value>' -ForEach @(@{ Value = 1 }, @{ Value = 30 }, @{ Value = 365 }) {
        (Read-PrefsWithAppFlow @{ Action = 'Report'; DefaultCadenceDays = $Value }).AppFlow.DefaultCadenceDays | Should -Be $Value
    }

    It 'keeps 7 for <Value>' -ForEach @(@{ Value = 0 }, @{ Value = -3 }, @{ Value = 366 }, @{ Value = 'weekly' }) {
        (Read-PrefsWithAppFlow @{ Action = 'Report'; DefaultCadenceDays = $Value }).AppFlow.DefaultCadenceDays | Should -Be 7
    }

    It 'keeps 7 in a file from before the key' {
        (Read-PrefsWithAppFlow @{ Action = 'Report' }).AppFlow.DefaultCadenceDays | Should -Be 7
    }

    It 'is saved by the One Click panel' {
        $text = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-AppFlowPanel' }, $false)).Extent.Text
        $text | Should -Match 'DefaultCadenceDays\s*=\s*\$newCadence'
        $text | Should -Match 'txtOneClickCadenceDays'
    }
}
