BeforeDiscovery {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
}

BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
    function Get-AuthoredManifest {
        param([string]$File)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            "$PSScriptRoot\..\Packagers\$File", [ref]$null, [ref]$null)
        $command = $ast.Find({ param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
                $node.GetCommandName() -eq 'Write-StageManifest'
        }, $true)
        $literal = $command.CommandElements | Where-Object {
            $_ -is [System.Management.Automation.Language.HashtableAst]
        }
        $version = '12.10.0'
        $installerFileName = 'test.exe'
        return ([pscustomobject]((& ([scriptblock]::Create($literal.Extent.Text))) |
            ConvertTo-Json -Depth 10 | ConvertFrom-Json))
    }
}

Describe 'Shipped user detection' {
    It 'uses the stable HKCU uninstall version for <File>' -ForEach @(
        @{ File = 'package-vscode-user.ps1'; Key = '{771FD6B0-FA20-440A-A002-3B3BAC16DC50}_is1' }
        @{ File = 'package-postman.ps1'; Key = 'Postman' }
    ) {
        $manifest = Get-AuthoredManifest $File
        $manifest.InstallationBehaviorType | Should -Be 'InstallForUser'
        $manifest.LogonRequirementType | Should -Be 'OnlyWhenUserLoggedOn'
        $det = $manifest.Detection
        $det.Type | Should -Be 'RegistryKeyValue'
        $det.Hive | Should -Be 'CurrentUser'
        $det.RegistryKeyRelative | Should -Be "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$Key"
        $det.ValueName | Should -Be 'DisplayVersion'
        $det.PropertyType | Should -Be 'Version'
        $det.Operator | Should -Be 'GreaterEquals'
        $det.ExpectedValue | Should -Be '12.10.0'
        $det.PSObject.Properties.Name | Should -Not -Contain 'FilePath'

        $rules = @(ConvertTo-IntuneWin32Rules -Manifest $manifest)
        $rules.Count | Should -Be 1
        $rules[0].keyPath | Should -Be "HKEY_CURRENT_USER\$($det.RegistryKeyRelative)"
        $rules[0].operationType | Should -Be 'version'
        $rules[0].operator | Should -Be 'greaterThanOrEqual'
        $rules[0].comparisonValue | Should -Be '12.10.0'
    }
}

Describe 'ConfigMgr registry detection hive routing' {
    InModuleScope AppPackagerCommon {
        BeforeAll {
            function New-CMDetectionClauseRegistryKeyValue {
                param($Hive, $KeyName, $ValueName, $PropertyType, $Value,
                    $ExpressionOperator, $ExpectedValue, $Is64Bit)
                return $PSBoundParameters
            }
            function New-CMDetectionClauseRegistryKey {
                param($Hive, $KeyName, $Existence, $Is64Bit)
                return $PSBoundParameters
            }
        }

        It 'preserves <Hive> as <ExpectedHive> for <Type>' -ForEach @(
            @{ Type = 'RegistryKeyValue'; Hive = 'CurrentUser'; ExpectedHive = 'CurrentUser' }
            @{ Type = 'RegistryKeyValue'; Hive = 'HKCU'; ExpectedHive = 'CurrentUser' }
            @{ Type = 'RegistryKeyValue'; Hive = $null; ExpectedHive = 'LocalMachine' }
            @{ Type = 'RegistryKey'; Hive = 'CurrentUser'; ExpectedHive = 'CurrentUser' }
            @{ Type = 'RegistryKey'; Hive = 'HKCU'; ExpectedHive = 'CurrentUser' }
            @{ Type = 'RegistryKey'; Hive = $null; ExpectedHive = 'LocalMachine' }
        ) {
            $clause = New-SingleDetectionClause -Det ([pscustomobject]@{
                Type = $Type; Hive = $Hive
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Postman'
                ValueName = 'DisplayVersion'; PropertyType = 'Version'
                Operator = 'GreaterEquals'; ExpectedValue = '12.10.0'; Is64Bit = $false
            })
            $clause.Hive | Should -Be $ExpectedHive
            $clause.KeyName | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Postman'
            if ($Type -eq 'RegistryKeyValue') {
                $clause.PropertyType | Should -Be 'Version'
                $clause.ExpressionOperator | Should -Be 'GreaterEquals'
                $clause.ExpectedValue | Should -Be '12.10.0'
            }
            else { $clause.Existence | Should -BeTrue }
        }
    }
}
