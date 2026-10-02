BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-windowsadmincenter.ps1'), [ref]$t, [ref]$e)
    $script:ParseErrors = $e

    $literal = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $n.Left.VariablePath.UserPath -eq 'ArpRegistryKey'
    }, $true)
    . ([scriptblock]::Create($literal.Extent.Text))

    $version = '2.7.5.21'
    $table = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.HashtableAst] -and
        @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'RegistryKeyRelative' }).Count -gt 0
    }, $true)
    $script:Detection = [pscustomobject](& ([scriptblock]::Create($table.Extent.Text)))
    $script:AssertCall = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Assert-ArpDetectionKey'
    }, $true)
}

Describe 'Windows Admin Center detection' {
    It 'parses without errors' {
        $script:ParseErrors | Should -BeNullOrEmpty
    }

    It 'compares the version on the 64-bit Inno uninstall entry' {
        $det = $script:Detection
        $det.Type | Should -Be 'RegistryKeyValue'
        $det.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\9B27DF2F-5386-41DF-B52B-5DF81914B043_is1'
        $det.ValueName | Should -Be 'DisplayVersion'
        $det.PropertyType | Should -Be 'Version'
        $det.Operator | Should -Be 'GreaterEquals'
        $det.ExpectedValue | Should -Be '2.7.5.21'
        $det.Is64Bit | Should -BeTrue
    }

    It 'maps to one Intune registry version rule' {
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest ([pscustomobject]@{ Detection = $script:Detection }))
        $rules.Count | Should -Be 1
        $rules[0].keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\9B27DF2F-5386-41DF-B52B-5DF81914B043_is1'
        $rules[0].operationType | Should -Be 'version'
        $rules[0].operator | Should -Be 'greaterThanOrEqual'
        $rules[0].check32BitOn64System | Should -BeFalse
    }

    It 'checks the installer against the same key and registry view before staging' {
        $script:AssertCall | Should -Not -BeNullOrEmpty
        $script:AssertCall.Extent.Text | Should -Match ([regex]::Escape('-ExpectedKey $ArpRegistryKey'))
        $script:AssertCall.Extent.Text | Should -Match ([regex]::Escape('-Is64BitView $true'))
    }
}
