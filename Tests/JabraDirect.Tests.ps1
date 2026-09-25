BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-jabradirect.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-JabraDirectVersionFromPage' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Jabra Direct release-notes version' {
    It 'reads the first release-version element' {
        $html = '<h3 data-testid="version-label">Release version</h3><span data-testid="release-version">8.2.23201</span>' +
                '<span data-testid="release-version">8.1.14601</span>'
        Get-JabraDirectVersionFromPage -Html $html | Should -Be '8.2.23201'
    }

    It 'tolerates whitespace around the version' {
        Get-JabraDirectVersionFromPage -Html "<span data-testid=`"release-version`">`n 6.27.03702 `n</span>" | Should -Be '6.27.03702'
    }

    It 'returns nothing when the page carries no marker: <Html>' -TestCases @(
        @{ Html = '<html>maintenance</html>' }
        @{ Html = '<span data-testid="release-version">latest</span>' }
        @{ Html = '' }
    ) {
        param($Html)
        Get-JabraDirectVersionFromPage -Html $Html | Should -BeNullOrEmpty
    }
}
