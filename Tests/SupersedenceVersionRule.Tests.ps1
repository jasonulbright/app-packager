#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The version rule, the Exact guard, the settings resolution, the blocked combination and the identity marker.

.DESCRIPTION
    The cases are the cases of the tests of App Packager Desktop (SupersedenceVersionRuleTests). Minimum writes GreaterEquals,
    Exact writes IsEquals, and only when the Exact guard holds a value that was read from the installed product.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Version = '3.7.7'

    function New-TestFileClause {
        param([string]$Op = 'GreaterEquals', [string]$Value = '3.7.7', [string]$PropertyType = $null)
        [pscustomobject]@{ Type = 'File'; FilePath = '%ProgramFiles%\Vendor'; FileName = 'app.exe'; ExpectedValue = $Value; Operator = $Op; PropertyType = $PropertyType }
    }
    function New-TestRegistryClause {
        param([string]$Op = 'GreaterEquals', [string]$Value = '3.7.7', [string]$Key = 'SOFTWARE\Vendor\App')
        [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $Key; ValueName = 'DisplayVersion'; ExpectedValue = $Value; Operator = $Op; PropertyType = 'String' }
    }
    function New-TestManifest {
        param($Detection, [string]$InstallerType = 'EXE')
        [pscustomobject]@{
            AppName = 'Vendor App'; SoftwareVersion = '3.7.7'; InstallerType = $InstallerType; ApplicationId = 'catalog:vendor-app'
            ProfileId = 'default'; Architecture = 'x64'; ProductCode = '{11111111-2222-3333-4444-555555555555}'; InstallerFile = 'setup.msi'
            Detection = $Detection
        }
    }
    $script:Allow = { param($clause) [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null } }
    $script:MsiKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{11111111-2222-3333-4444-555555555555}'
    function New-TestOverride { param($Settings, [string]$Key, [string]$VersionRule, [string]$Mode)
        $o = [ordered]@{}
        if ($VersionRule) { $o['VersionRule'] = $VersionRule }
        if ($Mode) { $o['Mode'] = $Mode }
        $Settings.TitleOverrides[$Key] = [pscustomobject]$o
    }
}

Describe 'The operator that the publish writes' {
    It 'writes IsEquals for Exact and changes nothing for Minimum' {
        $exact = Set-DetectionVersionRule -Detection (New-TestFileClause) -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow
        $minimum = Set-DetectionVersionRule -Detection (New-TestFileClause) -PackagedVersion $script:Version -Rule Minimum -Guard $script:Allow
        $exact.Detection.Operator | Should -Be 'IsEquals'
        @($exact.Outcomes).Count | Should -Be 1
        $minimum.Detection.Operator | Should -Be 'GreaterEquals'
        @($minimum.Outcomes).Count | Should -Be 0
    }

    It 'treats a file clause with no operator as the default GreaterEquals' {
        $clause = New-TestFileClause -Op ''
        $clause.Operator = $null
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow).Detection.Operator | Should -Be 'IsEquals'
    }

    It 'changes an MSI registry version clause like any other version clause' {
        $clause = New-TestRegistryClause -Key $script:MsiKey
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow).Detection.Operator | Should -Be 'IsEquals'
    }

    It 'does not change a clause with operator <Op> and expected value <Value>' -ForEach @(
        @{ Op = 'IsEquals'; Value = '3.7.7' }, @{ Op = 'GreaterThan'; Value = '3.7.7' }, @{ Op = 'LessEquals'; Value = '3.7.7' }, @{ Op = 'GreaterEquals'; Value = '9.9.9' }
    ) {
        $result = Set-DetectionVersionRule -Detection (New-TestFileClause -Op $Op -Value $Value) -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow
        $result.Detection.Operator | Should -Be $Op
        @($result.Outcomes).Count | Should -Be 0
    }

    It 'does not change an existence clause or a script clause' {
        $existence = New-TestFileClause -PropertyType 'Existence'
        $script1 = [pscustomobject]@{ Type = 'Script'; ScriptText = 'Write-Output 1'; ExpectedValue = $script:Version; Operator = 'GreaterEquals' }
        (Set-DetectionVersionRule -Detection $existence -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow).Detection.Operator | Should -Be 'GreaterEquals'
        (Set-DetectionVersionRule -Detection $script1 -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow).Detection.Operator | Should -Be 'GreaterEquals'
    }

    It 'changes only the version clauses of a compound detection' {
        $compound = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; Clauses = @((New-TestFileClause), (New-TestFileClause -Op 'GreaterEquals' -Value 'other'), (New-TestRegistryClause -Op 'IsEquals')) }
        $result = Set-DetectionVersionRule -Detection $compound -PackagedVersion $script:Version -Rule Exact -Guard $script:Allow
        (@($result.Detection.Clauses | ForEach-Object { $_.Operator }) -join ',') | Should -Be 'IsEquals,GreaterEquals,IsEquals'
        @($result.Outcomes).Count | Should -Be 1
    }

    It 'leaves the manifest alone and covers the deployment type detections' {
        $manifest = New-TestManifest -Detection (New-TestFileClause)
        $manifest | Add-Member -NotePropertyName DeploymentTypes -NotePropertyValue @([pscustomobject]@{ NameSuffix = 'x64'; Detection = (New-TestRegistryClause) })
        $result = Set-ManifestVersionRule -Manifest $manifest -Rule Exact -Guard $script:Allow
        $manifest.Detection.Operator | Should -Be 'GreaterEquals'
        $manifest.DeploymentTypes[0].Detection.Operator | Should -Be 'GreaterEquals'
        $result.Manifest.Detection.Operator | Should -Be 'IsEquals'
        $result.Manifest.DeploymentTypes[0].Detection.Operator | Should -Be 'IsEquals'
        @($result.Outcomes).Count | Should -Be 2
        $result.Manifest.ApplicationId | Should -Be $manifest.ApplicationId
    }

    It 'gives the same manifest back for Minimum' {
        $manifest = New-TestManifest -Detection (New-TestFileClause)
        $result = Set-ManifestVersionRule -Manifest $manifest -Rule Minimum -Guard $script:Allow
        [object]::ReferenceEquals($result.Manifest, $manifest) | Should -BeTrue
        @($result.Outcomes).Count | Should -Be 0
    }
}

Describe 'The Exact guard' {
    It 'applies Exact only when the installed value equals the expected value' {
        $recorded = @{ (Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version) = '3.7.7' }
        (Get-ExactGuardVerdict -Manifest (New-TestManifest (New-TestFileClause)) -Clause (New-TestFileClause) -TitleKey 'catalog:vendor-app' -Recorded $recorded).Result | Should -Be 'ExactApplied'
    }

    It 'blocks Exact for an installed value with more parts' {
        $recorded = @{ (Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version) = '3.7.7.1234' }
        $verdict = Get-ExactGuardVerdict -Manifest (New-TestManifest (New-TestFileClause)) -Clause (New-TestFileClause) -TitleKey 'catalog:vendor-app' -Recorded $recorded
        $verdict.Result | Should -Be 'ExactBlocked'
        $verdict.Detail | Should -Match '3\.7\.7\.1234'
        $verdict.Detail | Should -Match '3\.7\.7'
    }

    It 'blocks Exact without an installed value and says so' {
        $verdict = Get-ExactGuardVerdict -Manifest (New-TestManifest (New-TestFileClause)) -Clause (New-TestFileClause) -TitleKey 'catalog:vendor-app' -Recorded @{}
        $verdict.Result | Should -Be 'ExactBlocked'
        $verdict.Detail | Should -Match 'No installed value'
    }

    It 'does not ignore zero parts but normalizes the separators' {
        Test-SupersedenceSameVersionText '3.7.7' '3.7.7.0' | Should -BeFalse
        Test-SupersedenceSameVersionText '3.7.7+1' '3.7.7.1' | Should -BeTrue
        Test-SupersedenceSameVersionText '3.7.7-2' '3.7.7.2' | Should -BeTrue
    }

    It 'takes the ProductVersion of the MSI as the installed value of the MSI clause' {
        $clause = New-TestRegistryClause -Key $script:MsiKey
        $manifest = New-TestManifest -Detection $clause -InstallerType 'MSI'
        (Get-ExactGuardVerdict -Manifest $manifest -Clause $clause -TitleKey 'catalog:vendor-app' -MsiProductVersion { '3.7.7' }).Result | Should -Be 'ExactApplied'
        (Get-ExactGuardVerdict -Manifest $manifest -Clause $clause -TitleKey 'catalog:vendor-app' -MsiProductVersion { '3.7.7.0' }).Result | Should -Be 'ExactBlocked'
        (Get-ExactGuardVerdict -Manifest $manifest -Clause $clause -TitleKey 'catalog:vendor-app' -MsiProductVersion { $null }).Result | Should -Be 'ExactBlocked'
    }

    It 'holds an accepted block for that version only and writes no Exact' {
        $accepted = @(Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version)
        (Get-ExactGuardVerdict -Manifest (New-TestManifest (New-TestFileClause)) -Clause (New-TestFileClause) -TitleKey 'catalog:vendor-app' -Recorded @{} -Accepted $accepted).Result | Should -Be 'ExactAcceptedBlock'
        $next = New-TestManifest (New-TestFileClause -Value '3.7.8')
        $next.SoftwareVersion = '3.7.8'
        (Get-ExactGuardVerdict -Manifest $next -Clause (New-TestFileClause -Value '3.7.8') -TitleKey 'catalog:vendor-app' -Recorded @{} -Accepted $accepted).Result | Should -Be 'ExactBlocked'
    }
}

Describe 'The settings and the blocked combination' {
    It 'ships Minimum, Minimum and Off' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule | Should -Be 'Minimum'
        $settings.IntuneVersionRule | Should -Be 'Minimum'
        $settings.Mode | Should -Be 'Off'
        $resolved = Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app'
        $resolved.ConfigMgrRule.Value | Should -Be 'Minimum'
        $resolved.Mode.Value | Should -Be 'Off'
        $resolved.Blocked | Should -BeNullOrEmpty
    }

    It 'resolves each setting by itself from the command line, then the title, then the site' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Minimum'; $settings.Mode = 'Replace'
        New-TestOverride $settings 'catalog:vendor-app' -VersionRule 'Exact'
        $title = Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app'
        $cli = Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app' -CommandLine ([pscustomobject]@{ VersionRule = 'minimum'; Mode = 'replace-uninstall' })
        $other = Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:other'
        $title.ConfigMgrRule.Value | Should -Be 'Exact';  $title.ConfigMgrRule.Source | Should -Be 'TitleOverride'
        $title.Mode.Value | Should -Be 'Replace';          $title.Mode.Source | Should -Be 'SiteDefault'
        $cli.ConfigMgrRule.Value | Should -Be 'Minimum';  $cli.ConfigMgrRule.Source | Should -Be 'CommandLine'
        $cli.Mode.Value | Should -Be 'ReplaceAndUninstall'; $cli.Mode.Source | Should -Be 'CommandLine'
        $other.ConfigMgrRule.Value | Should -Be 'Minimum'; $other.ConfigMgrRule.Source | Should -Be 'SiteDefault'
    }

    It 'needs Exact for Replace and uninstall: <Rule> with <Mode>' -ForEach @(
        @{ Rule = 'Minimum'; Mode = 'ReplaceAndUninstall'; Blocked = $true }
        @{ Rule = 'Exact'; Mode = 'ReplaceAndUninstall'; Blocked = $false }
        @{ Rule = 'Minimum'; Mode = 'Replace'; Blocked = $false }
        @{ Rule = 'Minimum'; Mode = 'Off'; Blocked = $false }
    ) {
        $text = Get-SupersedenceBlockedCombination -ConfigMgrRule $Rule -Mode $Mode
        ($null -ne $text) | Should -Be $Blocked
        if ($Blocked) { $text | Should -BeLike 'Replace and uninstall needs the version rule Exact.*' }
    }

    It 'checks the block after resolution, so an override of the rule alone can cause it' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'; $settings.Mode = 'Replace and uninstall'
        New-TestOverride $settings 'catalog:self-updating' -VersionRule 'Minimum'
        (Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app').Blocked | Should -BeNullOrEmpty
        (Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:self-updating').Blocked | Should -Not -BeNullOrEmpty
    }

    It 'refuses a word that is not allowed and names its place' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Sometimes'
        $message = $null
        try { Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app' | Out-Null } catch { $message = $_.Exception.Message }
        $message | Should -Match 'Sometimes'
        $message | Should -Match 'site default'
    }

    It 'gives a title without a definition no override key' {
        Get-SupersedenceOverrideKey ' catalog:vendor-app ' | Should -Be 'catalog:vendor-app'
        Get-SupersedenceOverrideKey '' | Should -BeNullOrEmpty
        Get-SupersedenceOverrideKey $null | Should -BeNullOrEmpty
    }

    It 'matches an override key without regard to letter case, also from a file load' {
        $settings = New-SupersedenceSettings
        $settings.TitleOverrides = ('{"Catalog:Vendor-App":{"Mode":"Replace"}}' | ConvertFrom-Json)
        (Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:vendor-app').Mode.Value | Should -Be 'Replace'
    }
}

Describe 'The plan of one publish' {
    It 'publishes with the shipped defaults the same detection, and only the marker is new' {
        $manifest = New-TestManifest (New-TestFileClause)
        $plan = Get-SupersedencePlan -Manifest $manifest -Settings (New-SupersedenceSettings) -PublishesToConfigMgr $true -PublishesToIntune $true
        $plan.Refusal | Should -BeNullOrEmpty
        $plan.IntuneManifest.Detection.Operator | Should -Be 'GreaterEquals'
        $plan.Marker | Should -Be 'APD-Line:catalog:vendor-app|default|x64'
        ($plan.ConfigMgrManifest | ConvertTo-Json -Depth 8) | Should -Be ($manifest | ConvertTo-Json -Depth 8)
    }

    It 'takes the identifier from the run for a grid stage of the default profile' {
        $manifest = New-TestManifest (New-TestFileClause)
        $manifest.ApplicationId = ''
        $settings = New-SupersedenceSettings
        New-TestOverride $settings 'catalog:vendor-app' -Mode 'Replace'
        $without = Get-SupersedencePlan -Manifest $manifest -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false
        $with = Get-SupersedencePlan -Manifest $manifest -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false -FallbackApplicationId 'catalog:vendor-app'
        $without.Marker | Should -BeNullOrEmpty
        $with.Marker | Should -Be 'APD-Line:catalog:vendor-app|default|x64'
        @($with.Log | Where-Object { $_ -match 'Replace' -and $_ -match 'title override' }).Count | Should -BeGreaterThan 0
    }

    It 'blocks Replace and uninstall for a title with an accepted block' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'; $settings.Mode = 'Replace and uninstall'
        $settings.AcceptedExactBlocks = @(Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version)
        New-TestOverride $settings 'catalog:vendor-app' -Mode 'Replace and uninstall'
        (Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false).Refusal | Should -Match 'Replace and uninstall is blocked'
    }

    It 'does not stop the publish for a null override map from a hand-edited file' {
        $settings = New-SupersedenceSettings
        $settings.TitleOverrides = $null
        (Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false).Refusal | Should -BeNullOrEmpty
    }

    It 'refuses the row for a blocked combination and creates nothing' {
        $settings = New-SupersedenceSettings
        $settings.Mode = 'Replace and uninstall'
        $plan = Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false
        $plan.Refusal | Should -Be (Get-SupersedenceBlockedCombination -ConfigMgrRule 'Minimum' -Mode 'ReplaceAndUninstall')
        $plan.Marker | Should -BeNullOrEmpty
    }

    It 'does not block an Intune only run with the ConfigMgr combination' {
        $settings = New-SupersedenceSettings
        $settings.Mode = 'Replace and uninstall'
        (Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $false -PublishesToIntune $true).Refusal | Should -BeNullOrEmpty
    }

    It 'refuses Exact without an installed value and names the way out' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'
        $plan = Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false
        $plan.Refusal | Should -Match 'No installed value'
        $plan.Refusal | Should -Match 'accept the block'
    }

    It 'writes IsEquals to the target that asked for it when a value is recorded' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'; $settings.IntuneVersionRule = 'Minimum'
        $settings.InstalledValues[(Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version)] = '3.7.7'
        $manifest = New-TestManifest (New-TestFileClause)
        $plan = Get-SupersedencePlan -Manifest $manifest -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $true
        $plan.Refusal | Should -BeNullOrEmpty
        $plan.ConfigMgrManifest.Detection.Operator | Should -Be 'IsEquals'
        $plan.IntuneManifest.Detection.Operator | Should -Be 'GreaterEquals'
        $manifest.Detection.Operator | Should -Be 'GreaterEquals'
    }

    It 'publishes Minimum for an accepted block and logs it' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'
        $settings.AcceptedExactBlocks = @(Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version)
        $plan = Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false
        $plan.Refusal | Should -BeNullOrEmpty
        $plan.ConfigMgrManifest.Detection.Operator | Should -Be 'GreaterEquals'
        @($plan.Log | Where-Object { $_ -match 'block accepted' }).Count | Should -BeGreaterThan 0
    }

    It 'gives an MSI title Exact from the ProductVersion with no record' {
        $clause = New-TestRegistryClause -Key $script:MsiKey
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'
        $plan = Get-SupersedencePlan -Manifest (New-TestManifest $clause 'MSI') -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false -MsiProductVersion { '3.7.7' }
        $plan.Refusal | Should -BeNullOrEmpty
        $plan.ConfigMgrManifest.Detection.Operator | Should -Be 'IsEquals'
    }

    It 'refuses the row with the message for a bad word in the settings' {
        $settings = New-SupersedenceSettings
        $settings.Mode = 'Maybe'
        (Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false).Refusal | Should -Match 'Maybe'
    }

    It 'carries the resolved mode in the plan' {
        $settings = New-SupersedenceSettings
        $settings.ConfigMgrVersionRule = 'Exact'; $settings.Mode = 'Replace'
        $settings.InstalledValues[(Get-SupersedenceRecordKey 'catalog:vendor-app' $script:Version)] = '3.7.7'
        (Get-SupersedencePlan -Manifest (New-TestManifest (New-TestFileClause)) -Settings $settings -PublishesToConfigMgr $true -PublishesToIntune $false).Mode | Should -Be 'Replace'
    }
}

Describe 'The identity marker' {
    It 'holds the identity, the profile and the architecture' {
        $manifest = New-TestManifest (New-TestFileClause)
        $manifest.ProfileId = 'pilot'; $manifest.Architecture = 'ARM64'
        Get-SupersedenceIdentityMarker -Manifest $manifest | Should -Be 'APD-Line:catalog:vendor-app|pilot|ARM64'
    }

    It 'is absent for a title without an application identifier' {
        $manifest = New-TestManifest (New-TestFileClause)
        $manifest.ApplicationId = ''
        Get-SupersedenceIdentityMarker -Manifest $manifest | Should -BeNullOrEmpty
    }

    It 'is written after the application exists and never rewrites the keywords of the administrator' {
        $text = (Get-Command Set-CMApplicationIdentityMarker).ScriptBlock.ToString()
        $text | Should -Match 'Set-CMApplication -InputObject \$app -Keyword \$Marker'
        $text | Should -Match 'Identity marker not written'
        $text | Should -Match 'Get-CMApplication -Name \$ApplicationName -DisableWildcardHandling'
        $text | Should -Match 'if \(\$others\.Count -gt 0\) \{ throw'
        $text.IndexOf('$others') | Should -BeLessThan $text.IndexOf('Set-CMApplication -InputObject')
        $text.IndexOf('SccmSerializer') | Should -BeLessThan $text.IndexOf('Set-CMApplication -InputObject')
        $publish = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $publish.IndexOf('New-CMApplication @cmAppParams') | Should -BeLessThan $publish.IndexOf('Set-CMApplicationIdentityMarker -ApplicationName')
    }

    It 'leaves the publish as it was when the plan has no marker' {
        $publish = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $publish | Should -Match 'if \(-not \[string\]::IsNullOrWhiteSpace\(\[string\]\$supersedencePlan\.Marker\)\)'
    }
}
