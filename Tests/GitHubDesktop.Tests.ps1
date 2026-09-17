BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-githubdesktop.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertFrom-GitHubDesktopDownloadUrl' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'GitHub Desktop download redirect' {
    It 'reads <Expected> from <Url>' -TestCases @(
        @{ Url = 'https://desktop.githubusercontent.com/releases/3.6.6-8b85519e/GitHubDesktopSetup-x64.exe'; Expected = '3.6.6' }
        @{ Url = 'https://desktop.githubusercontent.com/releases/3.7.10-0a1b2c3d/GitHubDesktopSetup-x64.exe'; Expected = '3.7.10' }
    ) {
        param($Url, $Expected)
        ConvertFrom-GitHubDesktopDownloadUrl -Url $Url | Should -Be $Expected
    }

    It 'refuses a URL that is not the versioned x64 installer: <Url>' -TestCases @(
        @{ Url = 'https://central.github.com/deployments/desktop/desktop/latest/win32' }
        @{ Url = 'https://desktop.githubusercontent.com/releases/3.6.6-8b85519e/GitHubDesktopSetup-x64.msi' }
        @{ Url = 'https://desktop.githubusercontent.com/releases/3.6.6-8b85519e/GitHubDesktopSetup-arm64.exe' }
        @{ Url = '' }
    ) {
        param($Url)
        ConvertFrom-GitHubDesktopDownloadUrl -Url $Url | Should -BeNullOrEmpty
    }
}
