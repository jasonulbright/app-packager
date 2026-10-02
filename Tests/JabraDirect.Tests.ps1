BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-jabradirect.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $functions = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
    foreach ($name in 'Get-JabraDirectVersionFromPage', 'New-JabraDirectDetection', 'New-JabraDirectUninstallContent') {
        $fn = $functions | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $literals = $ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        @('ProductDisplayName', 'BundleUpgradeCode') -contains $n.Left.VariablePath.UserPath
    }, $true)
    foreach ($literal in $literals) { . ([scriptblock]::Create($literal.Extent.Text)) }
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

Describe 'Jabra Direct detection' {
    It 'names the product and the bundle upgrade code' {
        $ProductDisplayName | Should -Be 'Jabra Direct'
        $BundleUpgradeCode | Should -Be '{356870DC-69C0-4757-B09A-22C1786C104C}'
    }

    It 'compares the version of the installed executable' {
        $det = New-JabraDirectDetection -Version '8.2.23201'
        $det.Type | Should -Be 'File'
        $det.FilePath | Should -Be (Join-Path $env:ProgramFiles 'Jabra\Direct6')
        $det.FileName | Should -Be 'jabra-direct.exe'
        $det.PropertyType | Should -Be 'Version'
        $det.Operator | Should -Be 'GreaterEquals'
        $det.Is64Bit | Should -BeTrue
    }

    It 'expects the release version as printed' {
        (New-JabraDirectDetection -Version '8.3.26701').ExpectedValue | Should -Be '8.3.26701'
    }

    It 'maps to one Intune file version rule' {
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest ([pscustomobject]@{ Detection = [pscustomobject](New-JabraDirectDetection -Version '8.3.26701') }))
        $rules.Count | Should -Be 1
        $rules[0].fileOrFolderName | Should -Be 'jabra-direct.exe'
        $rules[0].operationType | Should -Be 'version'
        $rules[0].operator | Should -Be 'greaterThanOrEqual'
        $rules[0].comparisonValue | Should -Be '8.3.26701'
    }
}

Describe 'Jabra Direct uninstall wrapper' {
    BeforeAll {
        $script:Content = New-JabraDirectUninstallContent
        $tokens = $null; $errors = $null
        $contentAst = [System.Management.Automation.Language.Parser]::ParseInput($script:Content, [ref]$tokens, [ref]$errors)
        $script:ContentErrors = $errors
        $where = $contentAst.Find({ param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Where-Object'
        }, $true)
        $script:Filter = $where.CommandElements[1].ScriptBlock.GetScriptBlock()
    }

    It 'parses without errors and leaves no placeholder' {
        $script:ContentErrors | Should -BeNullOrEmpty
        $script:Content | Should -Not -Match '__[A-Z]+__'
    }

    It 'does not address an uninstall key by bundle id' {
        $script:Content | Should -Not -Match 'Uninstall\\\{'
    }

    It 'selects the bundle entry: <Name>' -TestCases @(
        @{ Name = 'by display name'; Entry = [pscustomobject]@{ DisplayName = 'Jabra Direct'; QuietUninstallString = '"C:\Cache\{a}\JabraDirectSetup.exe" /uninstall /quiet'; BundleUpgradeCode = $null } }
        @{ Name = 'by upgrade code list'; Entry = [pscustomobject]@{ DisplayName = 'Renamed'; QuietUninstallString = '"C:\Cache\{a}\JabraDirectSetup.exe" /uninstall /quiet'; BundleUpgradeCode = @('{00000000-0000-0000-0000-000000000000}', '{356870DC-69C0-4757-B09A-22C1786C104C}') } }
        @{ Name = 'by upgrade code text'; Entry = [pscustomobject]@{ DisplayName = 'Renamed'; QuietUninstallString = '"C:\Cache\{a}\JabraDirectSetup.exe" /uninstall /quiet'; BundleUpgradeCode = '{356870DC-69C0-4757-B09A-22C1786C104C}' } }
    ) {
        param($Entry)
        @($Entry | Where-Object $script:Filter).Count | Should -Be 1
    }

    It 'skips the entry: <Name>' -TestCases @(
        @{ Name = 'a package inside the bundle'; Entry = [pscustomobject]@{ DisplayName = 'Jabra Direct'; UninstallString = 'MsiExec.exe /X{00000000-0000-0000-0000-000000000000}'; BundleUpgradeCode = $null } }
        @{ Name = 'another product'; Entry = [pscustomobject]@{ DisplayName = 'Other Product'; QuietUninstallString = '"C:\Cache\{b}\Other.exe" /uninstall /quiet'; BundleUpgradeCode = '{11111111-1111-1111-1111-111111111111}' } }
    ) {
        param($Entry)
        @($Entry | Where-Object $script:Filter).Count | Should -Be 0
    }
}
