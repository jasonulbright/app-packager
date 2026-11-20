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
        $result.Summary | Should -BeLike 'Supersedence: no older version of*'
    }

    It 'does not take version <Version> as an old application of 2.0.0' -ForEach @(@{ Version = '2.0.0' }, @{ Version = '3.0.0' }) {
        (Select-SupersedenceOldApplication -Applications @(New-TestApp '1' 'Vendor App 2' $Version) -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0').Old | Should -BeNullOrEmpty
    }

    It 'takes only the highest older version: when it is retired there is no older version, and n-2 is not taken' {
        $apps = @((New-TestApp '1' 'Vendor App 1.0' '1.0.0'), (New-TestApp '2' 'Vendor App 1.5' '1.5.0' $true))
        $result = Select-SupersedenceOldApplication -Applications $apps -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0'
        $result.Old | Should -BeNullOrEmpty
        $result.Summary | Should -BeLike 'Supersedence: no older version of*'
        @($result.Notes | Where-Object { $_ -match "1\.5\.0 is the highest older version and it is retired" }).Count | Should -Be 1
    }

    It 'reads a pre-release as older than its release' {
        $result = Select-SupersedenceOldApplication -Applications @(New-TestApp '1' 'Vendor App 1.2.3-beta' '1.2.3-beta') -NewTitle 'Vendor App 1.2.3' -NewVersion '1.2.3'
        $result.Old.CiId | Should -Be '1'
        Compare-SupersedenceVersion -A '1.2.3-beta' -B '1.2.3' | Should -Be -1
        Compare-SupersedenceVersion -A '1.2.3' -B '1.2.3-beta' | Should -Be 1
        Compare-SupersedenceVersion -A '1.2.3' -B '1.2.3' | Should -Be 0
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
        function New-TestSiteApp { param($Id, $Title, $Version, $Vendor = 'Vendor', [bool]$Expired = $false) [pscustomobject]@{ CI_ID = $Id; LocalizedDisplayName = $Title; SoftwareVersion = $Version; IsExpired = $Expired; Manufacturer = $Vendor } }
        $script:Site = @(
            (New-TestSiteApp 1 'Vendor App 1.0' '1.0.0'),
            (New-TestSiteApp 2 'Vendor App 2.0' '2.0.0'),
            (New-TestSiteApp 3 'Unrelated' '0.1' 'Other'),
            (New-TestSiteApp 4 'Vendor App 0.5' '0.5.0' 'Another Vendor')
        )
    }
    AfterAll {
        InModuleScope AppPackagerCommon {
            foreach ($n in 'Get-CMApplication', 'Get-CMDeploymentType', 'Get-CMDeploymentTypeSupersedence', 'Set-CMApplicationSupersedence') { Remove-Item -LiteralPath "function:script:$n" -ErrorAction SilentlyContinue }
        }
    }
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Get-CMApplication -ModuleName AppPackagerCommon { @($script:Site | Where-Object { $_.LocalizedDisplayName -like $Name }) }
        Mock Get-CMDeploymentType -ModuleName AppPackagerCommon {
            [pscustomobject]@{ CI_ID = 99; LocalizedDisplayName = ($ApplicationName + ' - x64') }
        }
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @() }
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon {}
    }

    It 'reads the applications of the same title and the same vendor and always lists the new application' {
        $apps = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -BaseTitle 'Vendor App' -Vendor 'Vendor' -NewApplicationName 'Vendor App 2.0' }
        (@($apps | Where-Object { $_.VendorMatch } | ForEach-Object { $_.Title }) -join '|') | Should -Be 'Vendor App 1.0|Vendor App 2.0'
        $apps[0].DeploymentTypes[0].Name | Should -Be 'Vendor App 1.0 - x64'
    }

    It 'flags an application of another vendor and leaves out another title' {
        $apps = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -BaseTitle 'Vendor App' -Vendor 'Vendor' -NewApplicationName 'Vendor App 2.0' }
        @($apps | ForEach-Object { $_.Title }) | Should -Not -Contain 'Unrelated'
        ($apps | Where-Object { $_.Title -eq 'Vendor App 0.5' }).VendorMatch | Should -BeFalse
        ($apps | Where-Object { $_.Title -eq 'Vendor App 1.0' }).VendorMatch | Should -BeTrue
    }

    It 'does not take an application of another vendor, and the log line names both vendor strings' {
        $found = InModuleScope AppPackagerCommon { Find-SupersedenceOldApplication -BaseTitle 'Vendor App' -Vendor 'Vendor' -NewApplicationName 'Vendor App 0.9' -NewVersion '0.9.0' }
        $found.Old | Should -BeNullOrEmpty
        $found.Summary | Should -BeLike 'Supersedence: no older version of Vendor App 0.9 found; nothing to supersede.*'
        $found.Summary | Should -Match "vendor 'Another Vendor'"
        $found.Summary | Should -Match "vendor 'Vendor'"
    }

    It 'compares the vendor without regard to letter case, and matches nothing when the new application names no vendor' {
        $same = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -BaseTitle 'Vendor App' -Vendor 'VENDOR' -NewApplicationName 'Vendor App 2.0' }
        ($same | Where-Object { $_.Title -eq 'Vendor App 1.0' }).VendorMatch | Should -BeTrue
        $none = InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -BaseTitle 'Vendor App' -Vendor '' -NewApplicationName 'Vendor App 2.0' }
        ($none | Where-Object { $_.Title -eq 'Vendor App 1.0' }).VendorMatch | Should -BeFalse
        ($none | Where-Object { $_.Title -eq 'Vendor App 0.5' }).VendorMatch | Should -BeFalse
    }

    It 'finds no older version and says why when the new application names no vendor' {
        $found = InModuleScope AppPackagerCommon { Find-SupersedenceOldApplication -BaseTitle 'Vendor App' -Vendor '' -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' }
        $found.Old | Should -BeNullOrEmpty
        $found.Summary | Should -Match 'names no vendor'
    }

    It 'finds the highest version below the new one' {
        $found = InModuleScope AppPackagerCommon { Find-SupersedenceOldApplication -BaseTitle 'Vendor App' -Vendor 'Vendor' -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' }
        $found.Old.Title | Should -Be 'Vendor App 1.0'
    }

    It 'asks the site with the wildcard of the base title and never reads another title' {
        [void](InModuleScope AppPackagerCommon { Get-SupersedenceLineApplication -BaseTitle 'Vendor App' -Vendor 'Vendor' -NewApplicationName 'Vendor App 2.0' })
        Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Name -eq 'Vendor*App*' }
        Should -Invoke Get-CMDeploymentType -ModuleName AppPackagerCommon -Times 0 -Exactly -ParameterFilter { $ApplicationName -eq 'Unrelated' }
    }

    It 'sets one relation for each pair by name, with the uninstall option of the mode' {
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'ReplaceAndUninstall' }
        $step.Status | Should -Be 'Applied'
        $step.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0 and uninstalls it'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Name -eq 'Vendor App 2.0' -and $CurrentDeploymentTypeName -eq 'Vendor App 2.0 - x64' -and $SupersededApplicationName -eq 'Vendor App 1.0' -and
            $OldDeploymentTypeName -eq 'Vendor App 1.0 - x64' -and $IsUninstall -eq $true -and $DisableWildcardHandling
        }
    }

    It 'passes IsUninstall false for Replace' {
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $IsUninstall -eq $false }
    }

    It 'changes nothing on the older application' {
        [void](InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' })
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Name -eq 'Vendor App 2.0' }
    }

    It 'reads the relations first and does not write a relation that exists' {
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0 - x64' }) }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied'
        Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'counts a relation that is there after a failed write as done' {
        $global:SupWritten = $false
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { if ($global:SupWritten) { @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0 - x64' }) } else { @() } }
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon { $global:SupWritten = $true; throw 'a text of the provider' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied' -Because $step.Message
    }

    It 'reports a failed write as not applied with the reason of the first failed relation' {
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon { throw 'access denied' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Failure | Should -Be 'Vendor App 2.0 - x64 supersedes Vendor App 1.0 - x64: access denied'
    }

    It 'reports a failed lookup as not applied' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { throw 'provider down' }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Message | Should -Match 'lookup of the older version failed: provider down'
    }

    It 'asks nothing of the site for a mode of Off or an empty title' {
        $off = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'App' -NewVersion '1.0' -BaseTitle 'App' -Mode 'Off' }
        $none = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'App' -NewVersion '1.0' -BaseTitle '' -Mode 'Replace' }
        $off.Status | Should -Be 'NotRun'
        $none.Status | Should -Be 'NotRun'
        Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'says no older version found when only the new application is there' {
        Mock Get-CMApplication -ModuleName AppPackagerCommon { $script:Site | Where-Object { $_.CI_ID -eq 2 } }
        $step = InModuleScope AppPackagerCommon { Invoke-SupersedenceStep -NewApplicationName 'Vendor App 2.0' -NewVersion '2.0.0' -BaseTitle 'Vendor App' -Vendor 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NoOlderVersion'
        $step.Message | Should -BeLike 'Supersedence: no older version of*'
        $step.Failure | Should -BeNullOrEmpty
    }
}

Describe 'The title lock and the answer' {
    It 'knows whether the title holds the version of this release' {
        InModuleScope AppPackagerCommon {
            Test-SupersedenceVersionInTitle -Title 'Vendor App - 2.0.0' -Version '2.0.0' | Should -BeTrue
            Test-SupersedenceVersionInTitle -Title 'Vendor App' -Version '2.0.0' | Should -BeFalse
            Test-SupersedenceVersionInTitle -Title 'Vendor App 12.0.0' -Version '2.0.0' | Should -BeFalse
        }
    }
    It 'takes the base title from the title without the version' {
        InModuleScope AppPackagerCommon { Get-SupersedenceBaseTitle -AppName 'Vendor App - 2.0.0' -Version '2.0.0' | Should -Be 'Vendor App' }
    }
    It 'reads Yes only when it was given, and Skip for every other value and for none' {
        InModuleScope AppPackagerCommon {
            $saved = $env:APP_PACKAGER_SUPERSEDENCE_ANSWER
            try {
                $env:APP_PACKAGER_SUPERSEDENCE_ANSWER = 'Yes'; Get-SupersedenceAnswer | Should -Be 'Yes'
                $env:APP_PACKAGER_SUPERSEDENCE_ANSWER = 'yes'; Get-SupersedenceAnswer | Should -Be 'Yes'
                $env:APP_PACKAGER_SUPERSEDENCE_ANSWER = 'Skip'; Get-SupersedenceAnswer | Should -Be 'Skip'
                $env:APP_PACKAGER_SUPERSEDENCE_ANSWER = 'banana'; Get-SupersedenceAnswer | Should -Be 'Skip'
                Remove-Item Env:\APP_PACKAGER_SUPERSEDENCE_ANSWER; Get-SupersedenceAnswer | Should -Be 'Skip'
            }
            finally { if ($null -ne $saved) { $env:APP_PACKAGER_SUPERSEDENCE_ANSWER = $saved } else { Remove-Item Env:\APP_PACKAGER_SUPERSEDENCE_ANSWER -ErrorAction SilentlyContinue } }
        }
    }
}

Describe 'The bulk action' {
    BeforeAll {
        $script:Own = [ordered]@{ 'catalog:a' = ''; 'catalog:b' = 'Replace'; 'catalog:c' = 'Off' }
    }
    It 'All applications sets the default and removes the values of single applications' {
        $r = Set-SupersedenceBulkValue -Scope All -Value 'Replace' -Default 'Off' -Overrides $script:Own
        $r.Default | Should -Be 'Replace'
        (@($r.Overrides.Values | Where-Object { $_ }) ).Count | Should -Be 0
        $r.Changed | Should -Be 2
    }
    It 'Selected applications writes the value on each selected application only' {
        $r = Set-SupersedenceBulkValue -Scope Selected -Value 'Replace and uninstall' -Default 'Off' -Overrides $script:Own -Selected @('catalog:a')
        $r.Default | Should -Be 'Off'
        $r.Overrides['catalog:a'] | Should -Be 'Replace and uninstall'
        $r.Overrides['catalog:b'] | Should -Be 'Replace'
        $r.Changed | Should -Be 1
    }
    It 'No applications removes every value and sets the default to Off' {
        $r = Set-SupersedenceBulkValue -Scope None -Value 'Replace' -Default 'Replace' -Overrides $script:Own
        $r.Default | Should -Be 'Off'
        (@($r.Overrides.Values | Where-Object { $_ })).Count | Should -Be 0
    }
    It 'does not change the map it was given' {
        [void](Set-SupersedenceBulkValue -Scope None -Value 'Replace' -Default 'Replace' -Overrides $script:Own)
        $script:Own['catalog:b'] | Should -Be 'Replace'
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

Describe 'The Intune supersedence' {
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -like '*isof*') {
                return [pscustomobject]@{ value = @(
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'old1'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'new2'; displayName = 'Vendor App - 2.0.0'; displayVersion = '2.0.0'; publisher = 'Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'oth3'; displayName = 'Vendor App - 0.5.0'; displayVersion = '0.5.0'; publisher = 'Another Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'unr4'; displayName = 'Unrelated'; displayVersion = '0.1'; publisher = 'Vendor' }
                    ); '@odata.nextLink' = $null }
            }
            if ($Method -eq 'GET') { return [pscustomobject]@{ value = @() } }
            return $null
        }
    }

    It 'finds the older app by display name and publisher' {
        $apps = InModuleScope AppPackagerCommon { Get-IntuneSupersedenceApplication -Token 't' -BaseTitle 'Vendor App' -Publisher 'Vendor' }
        (@($apps | ForEach-Object { $_.Id }) -join '|') | Should -Be 'old1|new2'
    }

    It 'sends one supersedence of type update for Replace, to the highest older version' {
        $step = InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'Applied'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'POST' -and $Uri -like '*/mobileApps/new2/updateRelationships' -and
            $Body.relationships.Count -eq 1 -and $Body.relationships[0].targetId -eq 'old1' -and $Body.relationships[0].supersedenceType -eq 'update'
        }
    }

    It 'sends the type replace for Replace and uninstall' {
        [void](InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'ReplaceAndUninstall' })
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Body.relationships[0].supersedenceType -eq 'replace' }
    }

    It 'sends the relationships that the app already has again, because updateRelationships replaces them all' {
        $body = Get-IntuneSupersedenceBody -OldAppId 'old1' -Mode 'Replace' -Existing @(
            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppDependency'; targetId = 'dep9'; targetType = 'child'; dependencyType = 'detect' }
            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; targetId = 'old0'; targetType = 'child'; supersedenceType = 'update' }
            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; targetId = 'parent'; targetType = 'parent'; supersedenceType = 'update' }
        )
        (@($body.relationships | ForEach-Object { $_.targetId }) -join '|') | Should -Be 'dep9|old0|old1'
    }

    It 'writes nothing when the relationship exists, and says no older version when there is none' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Uri -like '*isof*') { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'old1'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Vendor' }, [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'new2'; displayName = 'Vendor App - 2.0.0'; displayVersion = '2.0.0'; publisher = 'Vendor' }) } }
            if ($Uri -like '*relationships') { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; targetId = 'old1'; targetType = 'child'; supersedenceType = 'update' }) } }
        }
        (InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }).Status | Should -Be 'Applied'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
        (InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 0.1.0' -NewVersion '0.1.0' -Publisher 'Vendor' -Mode 'Replace' }).Status | Should -Be 'NoOlderVersion'
    }

    It 'reports a failed Graph call as not applied and never throws' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon { throw 'Graph is down' }
        $step = InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Message | Should -Match 'Graph is down'
    }

    It 'matches nothing when the new app names no publisher, and says why' {
        $step = InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher '' -Mode 'Replace' }
        $step.Status | Should -Be 'NoOlderVersion'
        $step.Message | Should -Match 'names no publisher'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'does nothing for the mode Off' {
        (InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'x' -NewVersion '1' -Mode 'Off' }).Status | Should -Be 'NotRun'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 0 -Exactly
    }

    It 'uses beta only for the relationship read and write, and v1.0 for the lookup' {
        [void](InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' })
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'https://graph.microsoft.com/v1.0/*isof*' }
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $Uri -eq 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/new2/relationships' }
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/new2/updateRelationships' }
    }

    It 'names the beta base once, as a constant, and only the two relationship calls use it' {
        $source = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psm1') -Raw
        ([regex]::Matches($source, 'graph\.microsoft\.com/beta')).Count | Should -Be 1
        ([regex]::Matches($source, '\$RelationshipGraphBase/')).Count | Should -Be 2
        $source | Should -Match "IntuneRelationshipGraphBase = 'https://graph\.microsoft\.com/beta'"
    }

    It 'sends the dependencies and the older supersedences of the app again in the same call' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Uri -like '*isof*') { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'old1'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Vendor' }) } }
            if ($Method -eq 'GET') {
                return [pscustomobject]@{ value = @(
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppDependency'; targetId = 'dep9'; targetType = 'child'; dependencyType = 'autoInstall' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; targetId = 'old0'; targetType = 'child'; supersedenceType = 'replace' }
                    ) }
            }
            return $null
        }
        (InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }).Status | Should -Be 'Applied'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'POST' -and $Body.relationships.Count -eq 3 -and
            $Body.relationships[0].targetId -eq 'dep9' -and $Body.relationships[0].dependencyType -eq 'autoInstall' -and
            $Body.relationships[1].targetId -eq 'old0' -and $Body.relationships[1].supersedenceType -eq 'replace' -and
            $Body.relationships[2].targetId -eq 'old1' -and $Body.relationships[2].supersedenceType -eq 'update'
        }
    }

    It 'takes only the highest older version: n-2, a newer version, another vendor and another title stay untouched' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -like '*isof*') {
                return [pscustomobject]@{ value = @(
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'n2'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'n1'; displayName = 'Vendor App - 1.5.0'; displayVersion = '1.5.0'; publisher = 'Vendor Inc.' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'cur'; displayName = 'Vendor App - 2.0.0'; displayVersion = '2.0.0'; publisher = 'Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'nxt'; displayName = 'Vendor App - 3.0.0'; displayVersion = '3.0.0'; publisher = 'Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'wrv'; displayName = 'Vendor App - 1.9.0'; displayVersion = '1.9.0'; publisher = 'Another Vendor' }
                        [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'wrt'; displayName = 'Vendor Tool - 1.9.0'; displayVersion = '1.9.0'; publisher = 'Vendor' }
                    ) }
            }
            if ($Method -eq 'GET') { return [pscustomobject]@{ value = @() } }
            return $null
        }
        (InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'cur' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }).Status | Should -Be 'Applied'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Body.relationships.Count -eq 1 -and $Body.relationships[0].targetId -eq 'n1' }
    }

    It 'makes no relation when the only other app has another vendor' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -like '*isof*') { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'wrv'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Another Vendor' }) } }
            return [pscustomobject]@{ value = @() }
        }
        $step = InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NoOlderVersion'
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
    }

    It 'reports a failed relationship write with a clear message, never throws, and changes nothing else' {
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -like '*isof*') { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'old1'; displayName = 'Vendor App - 1.0.0'; displayVersion = '1.0.0'; publisher = 'Vendor' }) } }
            if ($Method -eq 'GET') { return [pscustomobject]@{ value = @() } }
            throw 'Graph returned 404 (resource not found)'
        }
        $step = InModuleScope AppPackagerCommon { Invoke-IntuneSupersedenceStep -Token 't' -AppId 'new2' -NewTitle 'Vendor App - 2.0.0' -NewVersion '2.0.0' -Publisher 'Vendor' -Mode 'Replace' }
        $step.Status | Should -Be 'NotApplied'
        $step.Message | Should -Match 'The Intune app is published\. The supersedence relationship was not created: Graph returned 404'
        $step.Failure | Should -Be $step.Message
        Should -Invoke Invoke-GraphJson -ModuleName AppPackagerCommon -Times 0 -Exactly -ParameterFilter { $Method -in @('PATCH', 'DELETE') }
    }

    It 'runs after the publish, only for the answer Yes, and a failed relation leaves the app published' {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\start-apppackager.ps1') -Raw
        $text.IndexOf('Publish-IntuneWin32App -TenantId') | Should -BeLessThan $text.IndexOf('Invoke-IntuneSupersedenceStep -Token')
        $text | Should -Match "-ne 'Off' -and \`$SupersedenceAnswer -eq 'Yes'"
        $text | Should -Match '\$pubNote\.Ok = \$true'
    }
}

Describe 'The note on the title page' {
    It 'is empty while supersedence is off, and says so while it is on' {
        InModuleScope AppPackagerCommon {
            Get-SupersedenceTitleLockNote -Supersedence ([pscustomobject]@{ Mode = 'Off'; TitleOverrides = [pscustomobject]@{} }) | Should -BeNullOrEmpty
            Get-SupersedenceTitleLockNote -Supersedence $null | Should -BeNullOrEmpty
            Get-SupersedenceTitleLockNote -Supersedence ([pscustomobject]@{ Mode = 'Replace'; TitleOverrides = [pscustomobject]@{} }) | Should -Match 'on by default'
            $own = [pscustomobject]@{ Mode = 'Off'; TitleOverrides = [pscustomobject]@{ 'catalog:a' = [pscustomobject]@{ Mode = 'Replace' }; 'catalog:b' = [pscustomobject]@{ Mode = 'Off' } } }
            Get-SupersedenceTitleLockNote -Supersedence $own | Should -Match 'on for 1 application'
        }
    }
}

Describe 'The vendor compare' {
    It 'reads <A> and <B> as the same vendor: <Same>' -ForEach @(
        @{ A = 'Contoso'; B = 'CONTOSO'; Same = $true }
        @{ A = 'Contoso Inc.'; B = 'Contoso'; Same = $true }
        @{ A = 'Contoso, Inc'; B = 'contoso corporation'; Same = $true }
        @{ A = 'Contoso GmbH'; B = 'Contoso Ltd.'; Same = $true }
        @{ A = 'Contoso Software'; B = 'Contoso'; Same = $false }
        @{ A = 'Contoso'; B = 'Fabrikam'; Same = $false }
        @{ A = 'Inc'; B = 'Inc.'; Same = $true }
    ) {
        InModuleScope AppPackagerCommon -Parameters @{ A = $A; B = $B; Same = $Same } {
            ((Get-SupersedenceVendorKey -Vendor $A) -eq (Get-SupersedenceVendorKey -Vendor $B)) | Should -Be $Same
        }
    }
    It 'names the application whose title gets the version' {
        InModuleScope AppPackagerCommon {
            Get-SupersedenceTitleNotice -Title 'Vendor App' | Should -Be 'Vendor App: supersedence is on, the version will be added to the title, so a new application is created'
        }
    }
}
