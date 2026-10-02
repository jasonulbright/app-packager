BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-ultravnc.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $functions = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
    foreach ($name in 'Get-UltraVncSetupPath', 'New-UltraVncUninstallContent') {
        $fn = $functions | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $literals = $ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $n.Left.VariablePath.UserPath -eq 'ArpRegistryKey'
    }, $true)
    foreach ($literal in $literals) { . ([scriptblock]::Create($literal.Extent.Text)) }
}

Describe 'UltraVNC uninstall wrapper' {
    BeforeAll {
        $script:Content = New-UltraVncUninstallContent
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($script:Content, [ref]$tokens, [ref]$errors)
        $script:ContentErrors = $errors
    }

    It 'parses without errors and leaves no placeholder' {
        $script:ContentErrors | Should -BeNullOrEmpty
        $script:Content | Should -Not -Match '__[A-Z]+__'
    }

    It 'finds the uninstaller through the ARP entry, not a fixed folder' {
        $script:Content | Should -Match 'Ultravnc2_is1'
        $script:Content | Should -Match 'UninstallString'
        $script:Content | Should -Not -Match 'Program Files|ProgramFiles|unins000'
    }

    It 'fails when the entry exists and its uninstaller is missing' {
        $script:Content | Should -Match 'Test-Path -LiteralPath \$exe\)\) \{ Write-Error .*exit 1'
    }
}

Describe 'UltraVNC setup entry on the release detail page' {
    It 'reads the x64 setup path from <Name> links' -TestCases @(
        @{
            Name     = 'category'
            Html     = '<a href="/all/summary/3-setup/506-ultravnc-1830-x86-setup.html">x86</a><a href="/all/summary/3-setup/508-ultravnc-1830-x64-setup.html">x64</a><a href="/all/summary/3-setup/509-ultravnc-1830-msi-x64.html">msi</a>'
            Expected = '3-setup/508-ultravnc-1830-x64-setup.html'
        }
        @{
            Name     = 'component'
            Html     = '<a href="/component/jdownloads/summary/3-setup/508-ultravnc-1830-x64-setup.html">x64</a>'
            Expected = '3-setup/508-ultravnc-1830-x64-setup.html'
        }
        @{
            Name     = 'component with query'
            Html     = '<a href="/component/jdownloads/summary/3-setup/508-ultravnc-1830-x64-setup.html?Itemid=101">x64</a>'
            Expected = '3-setup/508-ultravnc-1830-x64-setup.html?Itemid=101'
        }
    ) {
        param($Html, $Expected)
        Get-UltraVncSetupPath -Html $Html -Compact '1830' | Should -Be $Expected
    }

    It 'returns nothing when the release has no x64 setup entry: <Name>' -TestCases @(
        @{ Name = 'other release'; Html = '<a href="/all/summary/3-setup/508-ultravnc-1829-x64-setup.html">x64</a>' }
        @{ Name = 'x86 and msi only'; Html = '<a href="/all/summary/3-setup/506-ultravnc-1830-x86-setup.html">x86</a><a href="/all/summary/3-setup/509-ultravnc-1830-msi-x64.html">msi</a>' }
        @{ Name = 'foreign host'; Html = '<a href="https://example.test/all/summary/3-setup/508-ultravnc-1830-x64-setup.html">x64</a>' }
        @{ Name = 'empty page'; Html = '' }
    ) {
        param($Html)
        Get-UltraVncSetupPath -Html $Html -Compact '1830' | Should -BeNullOrEmpty
    }
}
