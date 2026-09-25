BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-signingsuite.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertFrom-SigningSuiteAssetName' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Signing Suite release asset' {
    It 'reads the version from the setup asset name' {
        ConvertFrom-SigningSuiteAssetName -Name 'SigningSuiteSetup-2026.09.15.0005.exe' | Should -Be '2026.09.15.0005'
    }

    It 'ignores every other release asset: <Name>' -TestCases @(
        @{ Name = 'SigningSuite-2026.09.15.0005.zip' }
        @{ Name = 'checksums.txt' }
        @{ Name = 'SigningSuiteSetup.exe' }
        @{ Name = '' }
    ) {
        param($Name)
        ConvertFrom-SigningSuiteAssetName -Name $Name | Should -BeNullOrEmpty
    }
}
