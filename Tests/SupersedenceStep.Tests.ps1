#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The lookup of the old application, the pairing of deployment types and the supersedence step.

.DESCRIPTION
    The cases are the cases of the tests of App Packager Desktop (SupersedenceStepTests). The lookup and the pairing are plain
    functions over objects. The step runs against mocked ConfigMgr cmdlets: it names applications and deployment types and not
    identifiers, reads the relations before it writes, and never fails the run.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function New-TestDt { param([string]$Id, [string]$Name) [pscustomobject]@{ CiId = $Id; Name = $Name } }
    function New-TestApp {
        param([string]$Id, [string]$Title, [string]$Version, [bool]$Expired = $false, [object[]]$Dts = $null)
        if (-not $Dts) { $Dts = @(New-TestDt ($Id + '1') ($Title + ' - x64')) }
        [pscustomobject]@{ CiId = $Id; Title = $Title; SoftwareVersion = $Version; IsExpired = $Expired; DeploymentTypes = $Dts }
    }
}

Describe 'The lookup of the old application' {
    It 'chooses the newest older application and names the others' {
        $apps = @((New-TestApp '1' 'Vendor App 1.0' '1.0.0'), (New-TestApp '2' 'Vendor App 1.5' '1.5.0'), (New-TestApp '3' 'Vendor App 2.0' '2.0.0'))
        $result = Select-SupersedenceOldApplication -Applications $apps -NewTitle 'Vendor App 3.0' -NewVersion '3.0.0'
        $result.Old.CiId | Should -Be '3'
        @($result.Notes | Where-Object { $_ -match 'Vendor App 1\.5' }).Count | Should -BeGreaterThan 0
    }

    It 'names a line that already holds a higher version and does not supersede it' {
        $apps = @((New-TestApp '1' 'Vendor App 1.0' '1.0.0'), (New-TestApp '2' 'Vendor App 2.0' '2.0.0'))
        $result = Select-SupersedenceOldApplication -Applications $apps -NewTitle 'Vendor App 1.5' -NewVersion '1.5.0'
        $result.Old.CiId | Should -Be '1'
        @($result.Notes | Where-Object { $_ -match 'Vendor App 2\.0' -and $_ -match 'not superseded' }).Count | Should -BeGreaterThan 0
    }

    It 'never returns an application with the title of the new one' {
        $result = Select-SupersedenceOldApplication -Applications @(New-TestApp '1' 'Vendor App' '1.0.0') -NewTitle 'vendor app' -NewVersion '2.0.0'
        $result.Old | Should -BeNullOrEmpty
        $result.Summary | Should -BeLike 'No older version found*'
    }

    It 'does not take version <Version> as an old application of 2.0.0' -ForEach @(@{ Version = '2.0.0' }, @{ Version = '3.0.0' }) {
        (Select-SupersedenceOldApplication -Applications @(New-TestApp '1' 'Vendor App 2' $Version) -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0').Old | Should -BeNullOrEmpty
    }

    It 'skips a retired application and takes the next lower one' {
        $apps = @((New-TestApp '1' 'Vendor App 1.0' '1.0.0'), (New-TestApp '2' 'Vendor App 1.5' '1.5.0' $true))
        $result = Select-SupersedenceOldApplication -Applications $apps -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0'
        $result.Old.CiId | Should -Be '1'
        @($result.Notes | Where-Object { $_ -match 'retired' }).Count | Should -BeGreaterThan 0
    }

    It 'skips a version that cannot be parsed and names it' {
        $apps = @((New-TestApp '1' 'Vendor App old' 'banana'), (New-TestApp '2' 'Vendor App 1.0' '1.0.0'))
        $result = Select-SupersedenceOldApplication -Applications $apps -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0'
        $result.Old.CiId | Should -Be '2'
        @($result.Notes | Where-Object { $_ -match 'banana' }).Count | Should -BeGreaterThan 0
    }

    It 'gives no old application for a new version that cannot be parsed' {
        $result = Select-SupersedenceOldApplication -Applications @(New-TestApp '1' 'Vendor App 1.0' '1.0.0') -NewTitle 'Vendor App' -NewVersion 'not-a-version'
        $result.Old | Should -BeNullOrEmpty
        $result.Summary | Should -Match 'cannot be read'
    }
}

Describe 'The pairing of deployment types' {
    It 'pairs one and one' {
        $pairing = Get-SupersedenceDeploymentTypePair -NewApplication (New-TestApp '9' 'New' '2.0') -OldApplication (New-TestApp '1' 'Old' '1.0')
        @($pairing.Pairs).Count | Should -Be 1
        $pairing.Pairs[0].New.CiId | Should -Be '91'
        $pairing.Pairs[0].Old.CiId | Should -Be '11'
        @($pairing.Warnings).Count | Should -Be 0
    }

    It 'pairs many by suffix and warns for an unmatched type, with no relation' {
        $newApp = New-TestApp '9' 'App 2' '2.0' $false @((New-TestDt '91' 'App 2 - x64'), (New-TestDt '92' 'App 2 - x86'), (New-TestDt '93' 'App 2 - ARM64'))
        $oldApp = New-TestApp '1' 'App 1' '1.0' $false @((New-TestDt '11' 'App 1 - x64'), (New-TestDt '12' 'App 1 - x86'))
        $pairing = Get-SupersedenceDeploymentTypePair -NewApplication $newApp -OldApplication $oldApp
        (@($pairing.Pairs | ForEach-Object { $_.New.CiId + '>' + $_.Old.CiId }) -join ',') | Should -Be '91>11,92>12'
        @($pairing.Warnings | Where-Object { $_ -match 'ARM64' -and $_ -match 'no partner' }).Count | Should -BeGreaterThan 0
    }

    It 'never cross links architectures when the suffixes differ' {
        $newApp = New-TestApp '9' 'App 2' '2.0' $false @((New-TestDt '91' 'App 2 - x64'), (New-TestDt '92' 'App 2 - x86'))
        $oldApp = New-TestApp '1' 'App 1' '1.0' $false @((New-TestDt '11' 'App 1 - ARM64'), (New-TestDt '12' 'App 1 - arm'))
        @((Get-SupersedenceDeploymentTypePair -NewApplication $newApp -OldApplication $oldApp).Pairs).Count | Should -Be 0
    }

    It 'supersedes nothing for an old application with no deployment type, and says so' {
        $old = [pscustomobject]@{ CiId = '1'; Title = 'Old'; SoftwareVersion = '1.0'; IsExpired = $false; DeploymentTypes = @() }
        $pairing = Get-SupersedenceDeploymentTypePair -NewApplication (New-TestApp '9' 'New' '2.0') -OldApplication $old
        @($pairing.Pairs).Count | Should -Be 0
        @($pairing.Warnings)[0] | Should -Match 'no deployment type'
    }
}

Describe 'The version compare' {
    It 'reads <Text> as <Expected>' -ForEach @(
        @{ Text = '1.2.3'; Expected = '1.2.3' }, @{ Text = '31'; Expected = '31.0' }, @{ Text = '2024.1-beta+5'; Expected = '2024.1' }
        @{ Text = '140.17.0esr'; Expected = '140.17.0' }, @{ Text = '1.2.3.4.5'; Expected = '1.2.3.4' }
    ) {
        ConvertTo-SupersedenceVersionText $Text | Should -Be $Expected
    }
    It 'compares over all four parts' {
        Compare-SupersedenceVersion '1.0.0' '1.0.0.1' | Should -Be -1
        Compare-SupersedenceVersion '2.0' '1.9.9.9' | Should -Be 1
        Compare-SupersedenceVersion '1.0' '1.0.0' | Should -Be 0
        Compare-SupersedenceVersion 'banana' '1.0' | Should -Be 0
    }
}

Describe 'The supersedence step' {
    BeforeAll {
        # The ConfigMgr cmdlets exist only with the console, so the module scope gets stubs that the cases mock.
        InModuleScope AppPackagerCommon {
            function script:Get-CMApplication { param($Name, [switch]$DisableWildcardHandling, $ErrorAction) }
            function script:Get-CMDeploymentType { param($ApplicationName, $DeploymentTypeName, [switch]$DisableWildcardHandling, $ErrorAction) }
            function script:Get-CMDeploymentTypeSupersedence { param($InputObject, $ErrorAction) }
            function script:Set-CMApplicationSupersedence { param($Name, $CurrentDeploymentTypeName, $SupersededApplicationName, $OldDeploymentTypeName, $IsUninstall, [switch]$DisableWildcardHandling, $ErrorAction) }
        }
        function New-TestCmApp { param($Id, $Title, $Version, [bool]$Expired = $false, $Xml = '<x>APD-Line:m</x>') [pscustomobject]@{ CI_ID = $Id; LocalizedDisplayName = $Title; SoftwareVersion = $Version; IsExpired = $Expired; SDMPackageXML = $Xml } }
        function New-TestCmDt { param($Id, $Name) [pscustomobject]@{ CI_ID = $Id; LocalizedDisplayName = $Name } }
        $script:Marker = 'APD-Line:catalog:vendor-app|default|x64'
    }
    AfterAll {
        InModuleScope AppPackagerCommon {
            foreach ($n in 'Get-CMApplication', 'Get-CMDeploymentType', 'Get-CMDeploymentTypeSupersedence', 'Set-CMApplicationSupersedence') { Remove-Item -LiteralPath "function:script:$n" -ErrorAction SilentlyContinue }
        }
    }
    BeforeEach {
        $m = '>' + $script:Marker + '<'
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            @(
                [pscustomobject]@{ CI_ID = 1; LocalizedDisplayName = 'Vendor App 1.0'; SoftwareVersion = '1.0.0'; IsExpired = $false; SDMPackageXML = '<k>APD-Line:catalog:vendor-app|default|x64</k>' }
                [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'Vendor App 2.0'; SoftwareVersion = '2.0.0'; IsExpired = $false; SDMPackageXML = '<none/>' }
                [pscustomobject]@{ CI_ID = 3; LocalizedDisplayName = 'Unrelated'; SoftwareVersion = '0.1'; IsExpired = $false; SDMPackageXML = '<k>other</k>' }
            )
        }
        Mock Get-CMDeploymentType -ModuleName AppPackagerCommon {
            [pscustomobject]@{ CI_ID = 99; LocalizedDisplayName = ($ApplicationName + ' - x64') }
        }
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @() }
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon {}
    }

    It 'reads the line by the marker and always lists the new application' {
        $apps = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -Marker 'APD-Line:catalog:vendor-app|default|x64' -NewApplicationName 'Vendor App 2.0' }
        (@($apps | ForEach-Object { $_.Title }) -join '|') | Should -Be 'Vendor App 1.0|Vendor App 2.0'
        $apps[0].DeploymentTypes[0].Name | Should -Be 'Vendor App 1.0 - x64'
    }

    It 'matches the marker as XML text with the characters escaped' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            [pscustomobject]@{ CI_ID = 1; LocalizedDisplayName = 'Old'; SoftwareVersion = '1.0'; IsExpired = $false; SDMPackageXML = '<Tag>APD-Line:catalog:o''neil&amp;co|default|x64</Tag>' }
        }
        $apps = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -Marker "APD-Line:catalog:o'neil&co|default|x64" -NewApplicationName 'New' }
        @($apps).Count | Should -Be 1
    }

    It 'sets one relation for each pair by name, with the uninstall option of the mode' {
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'ReplaceAndUninstall' }
        $step.Status | Should -Be 'Applied'
        $step.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0 and uninstalls it'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Name -eq 'Vendor App 2.0' -and $CurrentDeploymentTypeName -eq 'Vendor App 2.0 - x64' -and $SupersededApplicationName -eq 'Vendor App 1.0' -and
            $OldDeploymentTypeName -eq 'Vendor App 1.0 - x64' -and $IsUninstall -eq $true -and $DisableWildcardHandling
        }
    }

    It 'passes IsUninstall false for Replace' {
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $IsUninstall -eq $false }
    }

    It 'reads the relations first and does not write a relation that exists' {
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0 - x64' }) }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'counts a relation that is there after a failed write as done' {
        $script:Reads = 0
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { $script:Reads++; if ($script:Reads -ge 2) { @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0 - x64' }) } else { @() } }
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon { throw 'a text of the provider' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied'
    }

    It 'reports a failed write as not applied with the reason of the first failed relation' {
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon { throw 'access denied' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Failure | Should -Be 'Vendor App 2.0 - x64 supersedes Vendor App 1.0 - x64: access denied'
    }

    It 'reports a failed lookup as not applied' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { throw 'provider down' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Message | Should -Match 'lookup of the older version failed: provider down'
    }

    It 'asks nothing of the site for a mode of Off or a missing marker' {
        $off = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'App' -NewVersion '1.0' -Marker 'm' -Mode 'Off' }
        $none = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'App' -NewVersion '1.0' -Marker '' -Mode 'Replace' }
        $off.Status | Should -Be 'NotRun'
        $none.Status | Should -Be 'NotRun'
        Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'says no older version found when the line holds only the new application' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'Vendor App 2.0'; SoftwareVersion = '2.0.0'; IsExpired = $false; SDMPackageXML = '<none/>' }
        }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'NoOlderVersion'
        $step.Message | Should -BeLike 'No older version found*'
        $step.Failure | Should -BeNullOrEmpty
    }

    It 'applies without a publish only when the application of this version is on the site' {
        $missing = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Other App' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' -RequirePublished }
        $missing.Status | Should -Be 'NotApplied'
        $missing.Message | Should -Match 'is not on the site'
        $wrong = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '3.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' -RequirePublished }
        $wrong.Message | Should -Match 'not 3\.0\.0'
    }

    It 'keeps a name with quotes and wildcards as data' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            @(
                [pscustomobject]@{ CI_ID = 1; LocalizedDisplayName = 'O"ld [x64] 1.0'; SoftwareVersion = '1.0'; IsExpired = $false; SDMPackageXML = '<k>APD-Line:catalog:vendor-app|default|x64</k>' }
                [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'N"ew [x64] 2.0'; SoftwareVersion = '2.0'; IsExpired = $false; SDMPackageXML = '<none/>' }
            )
        }
        Mock Get-CMDeploymentType -ModuleName AppPackagerCommon { [pscustomobject]@{ CI_ID = 9; LocalizedDisplayName = $ApplicationName + ' - x64' } }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'N"ew [x64] 2.0' -NewVersion '2.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Name -eq 'N"ew [x64] 2.0' -and $SupersededApplicationName -eq 'O"ld [x64] 1.0' }
        Should -Invoke Get-CMDeploymentType -ModuleName AppPackagerCommon -ParameterFilter { $DisableWildcardHandling }
    }
}

Describe 'The step after a publish' {
    It 'runs when the application is complete and never fails the run' {
        $publish = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $publish | Should -Match 'Invoke-SupersedenceStepAndLog -ApplicationName \$appName'
        $publish.IndexOf('Start-CMContentDistribution') | Should -BeLessThan $publish.LastIndexOf('Invoke-SupersedenceStepAndLog')
        $publish.LastIndexOf('Invoke-SupersedenceStepAndLog') | Should -BeLessThan $publish.IndexOf('Created ConfigMgr application')
    }

    It 'warns with the separate outcome text and does not throw' {
        $text = (Get-Command Invoke-SupersedenceStepAndLog).ScriptBlock.ToString()
        $text | Should -Match 'Published, supersedence not applied: \{0\}'
        $text | Should -Match '-Level WARN'
    }

    It 'refuses a plan before the first write to the site' {
        $publish = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $publish.IndexOf('if ($supersedencePlan.Refusal) { throw') | Should -BeLessThan $publish.IndexOf('New-CMApplication @cmAppParams')
        $publish.IndexOf('if ($supersedencePlan.Refusal) { throw') | Should -BeLessThan $publish.IndexOf('Get-ManifestDeploymentTypeSpecs')
    }
}
