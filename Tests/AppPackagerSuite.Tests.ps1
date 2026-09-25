BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-apppackagersuite.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertFrom-AppPackagerSuiteAssetName' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'AppPackager Suite release asset' {
    It 'reads the version from the setup asset name' {
        ConvertFrom-AppPackagerSuiteAssetName -Name 'SuiteSetup-2026.09.21.0032.exe' | Should -Be '2026.09.21.0032'
    }

    It 'ignores every other release asset: <Name>' -TestCases @(
        @{ Name = 'AppPackagerSuite-2026.09.21.0032.zip' }
        @{ Name = 'checksums.txt' }
        @{ Name = 'SuiteSetup.exe' }
        @{ Name = '' }
    ) {
        param($Name)
        ConvertFrom-AppPackagerSuiteAssetName -Name $Name | Should -BeNullOrEmpty
    }
}
