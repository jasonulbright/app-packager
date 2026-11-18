#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    AppendArchitectureToTitle is a saved preference (default on), has an Options field and decides whether the packaged name keeps its trailing architecture tag.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Set-DeploymentTypeContentEnvironment', 'Get-OptionsSectionDefinition') {
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($node.Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
}

Describe 'AppendArchitectureToTitle preference' {
    It 'defaults to on when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).AppendArchitectureToTitle | Should -BeTrue
    }

    It 'reads <Value>' -ForEach @(@{ Value = $true }, @{ Value = $false }) {
        $script:PrefsFile = Join-Path $TestDrive ('p-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ AppendArchitectureToTitle = $Value } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).AppendArchitectureToTitle | Should -Be $Value
    }

    It 'keeps on in a file from before the key' {
        $script:PrefsFile = Join-Path $TestDrive 'old.json'
        ([pscustomobject]@{ SiteCode = 'ABC' } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).AppendArchitectureToTitle | Should -BeTrue
    }

    It 'sets the environment value for the packagers' -ForEach @(@{ On = $true; Expected = '1' }, @{ On = $false; Expected = '0' }) {
        $before = $env:APP_PACKAGER_TITLE_APPEND_ARCH
        try {
            Set-DeploymentTypeContentEnvironment -Prefs ([pscustomobject]@{ ContentFallback = 'Allow'; SlowNetworkDeploymentMode = 'Download'; AppendArchitectureToTitle = $On })
            $env:APP_PACKAGER_TITLE_APPEND_ARCH | Should -Be $Expected
        }
        finally { $env:APP_PACKAGER_TITLE_APPEND_ARCH = $before }
    }

    It 'has a field in Options, Packaging, Application Title' {
        $field = (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'title' }).Fields | Where-Object { $_.Id -eq 'title.arch' }
        $field.Pref | Should -Be 'AppendArchitectureToTitle'
    }
}

Describe 'Format-ApplicationTitleArchitecture' {
    It 'removes a trailing tag when off: <Title>' -ForEach @(
        @{ Title = '7-Zip (x64)'; Expected = '7-Zip' }
        @{ Title = 'Contoso App - 1.2.3 (x86)'; Expected = 'Contoso App - 1.2.3' }
        @{ Title = 'Contoso App (ARM64)'; Expected = 'Contoso App' }
        @{ Title = 'Contoso App (64-bit) '; Expected = 'Contoso App' }
    ) {
        Format-ApplicationTitleArchitecture -Title $Title -Append $false | Should -Be $Expected
    }

    It 'keeps the name when on, and a tag that is not trailing' {
        Format-ApplicationTitleArchitecture -Title '7-Zip (x64)' | Should -Be '7-Zip (x64)'
        Format-ApplicationTitleArchitecture -Title 'Contoso (x64) Edition' -Append $false | Should -Be 'Contoso (x64) Edition'
    }

    It 'turns a repeated tag into one' {
        Format-ApplicationTitleArchitecture -Title '7-Zip (x64) (x64)' | Should -Be '7-Zip (x64)'
    }

    It 'leaves a name without a tag alone' {
        Format-ApplicationTitleArchitecture -Title 'Notepad++' -Append $false | Should -Be 'Notepad++'
    }
}
