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

Describe 'GitHub Desktop install wrapper needs both outputs' {
    BeforeAll {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-githubdesktop.ps1'), [ref]$t, [ref]$e)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-GitHubDesktopInstallWrapper' }, $false)
        $script:InstallerFileName = 'GitHubDesktopSetup-x64.exe'
        . ([scriptblock]::Create($fn.Extent.Text))
        $wrapperAst = [System.Management.Automation.Language.Parser]::ParseInput((New-GitHubDesktopInstallWrapper), [ref]$t, [ref]$e)
        $while = $wrapperAst.Find({ param($n) $n -is [System.Management.Automation.Language.WhileStatementAst] }, $true)
        # The loop runs while this condition is true.
        $script:WaitCondition = [scriptblock]::Create('param($installedExe, $arpKey) ' + $while.Condition.Extent.Text)
        $script:ExeFile = Join-Path $TestDrive 'GitHubDesktop.exe'
        Set-Content -LiteralPath $script:ExeFile -Value 'x'
        $script:KeyPath = 'HKCU:\Software\AppPackagerTest-' + [guid]::NewGuid().ToString('N')
        $script:MissingExe = Join-Path $TestDrive 'missing.exe'
        $script:MissingKey = 'HKCU:\Software\AppPackagerTest-' + [guid]::NewGuid().ToString('N')
        New-Item -Path $script:KeyPath -Force | Out-Null
        New-ItemProperty -Path $script:KeyPath -Name DisplayVersion -Value '3.6.6' -PropertyType String | Out-Null
    }
    AfterAll { Remove-Item -Path $script:KeyPath -Recurse -Force -ErrorAction SilentlyContinue }

    It 'stops waiting when the program file and the key both exist' {
        (& $script:WaitCondition $script:ExeFile $script:KeyPath) | Should -BeFalse
    }
    It 'keeps waiting when only the program file exists' {
        (& $script:WaitCondition $script:ExeFile $script:MissingKey) | Should -BeTrue
    }
    It 'keeps waiting when only the key exists' {
        (& $script:WaitCondition $script:MissingExe $script:KeyPath) | Should -BeTrue
    }
    It 'keeps waiting when neither exists' {
        (& $script:WaitCondition $script:MissingExe $script:MissingKey) | Should -BeTrue
    }
}
