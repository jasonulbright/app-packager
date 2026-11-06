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

Describe 'GitHub Desktop install wrapper waits for the detection key' {
    BeforeAll {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-githubdesktop.ps1'), [ref]$t, [ref]$e)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-GitHubDesktopInstallWrapper' }, $false)
        $script:InstallerFileName = 'GitHubDesktopSetup-x64.exe'
        . ([scriptblock]::Create($fn.Extent.Text))
        $script:Wrapper = New-GitHubDesktopInstallWrapper
    }

    It 'is valid PowerShell' {
        $t = $null; $e = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($script:Wrapper, [ref]$t, [ref]$e)
        @($e).Count | Should -Be 0
    }

    It 'does not return before the per-user uninstall key holds DisplayVersion' {
        # ConfigMgr evaluates detection when the install returns. Detection reads
        # this key, which appears after the application files.
        $script:Wrapper | Should -Match ([regex]::Escape('Uninstall\GitHubDesktop'))
        $script:Wrapper | Should -Match 'DisplayVersion'
        $loop = [regex]::Match($script:Wrapper, '(?s)while \(.*?\r?\n\}')
        $loop.Success | Should -BeTrue
        $loop.Value | Should -Match 'DisplayVersion'
    }

    It 'still fails with exit code 4 when the key does not appear in time' {
        $script:Wrapper | Should -Match 'exit 4'
    }
}
