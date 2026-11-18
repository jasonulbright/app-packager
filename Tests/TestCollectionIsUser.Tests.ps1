#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    ContentDistribution.TestCollectionIsUser is a saved preference (default off), has an Options field and picks the user collection cmdlets.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-OptionsSectionDefinition') {
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($node.Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    $script:CommonText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psm1') -Raw
}

Describe 'ContentDistribution.TestCollectionIsUser' {
    It 'defaults to off when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).ContentDistribution.TestCollectionIsUser | Should -BeFalse
    }

    It 'reads <Value>' -ForEach @(@{ Value = $true }, @{ Value = $false }) {
        $script:PrefsFile = Join-Path $TestDrive ('p-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ ContentDistribution = [pscustomobject]@{ AutoDistribute = $true; TestCollectionIsUser = $Value } } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).ContentDistribution.TestCollectionIsUser | Should -Be $Value
    }

    It 'keeps off in a file from before the key' {
        $script:PrefsFile = Join-Path $TestDrive 'old.json'
        ([pscustomobject]@{ ContentDistribution = [pscustomobject]@{ AutoDistribute = $true } } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).ContentDistribution.TestCollectionIsUser | Should -BeFalse
    }

    It 'has a field in Options, Packaging, Post Packaging' {
        $field = (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'postpackaging' }).Fields | Where-Object { $_.Id -eq 'dist.user' }
        $field.Pref | Should -Be 'ContentDistribution.TestCollectionIsUser'
    }

    It 'uses the user collection cmdlets and All Users as the limiting collection' {
        $script:CommonText | Should -Match 'Get-CMUserCollection -Name \$testCollection'
        $script:CommonText | Should -Match 'New-CMUserCollection -Name \$testCollection -LimitingCollectionName \$limitingName'
        $script:CommonText | Should -Match "'All Users' } else { 'All Systems' }"
    }
}
