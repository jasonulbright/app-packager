#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The ConfigMgr requirements of an application become Intune requirement rules, or a blocking gap.

.DESCRIPTION
    Ported from the cases of IntuneRequirementRuleTests in App Packager Desktop. A WMI value, the
    operating system language and a script condition become PowerShell requirement rules. A
    requirement with no Intune rule is a Blocking finding unless its drop is accepted in
    Intune.AcceptedRequirementDrops; then it is a Review finding.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:Defaults = Get-DefaultConditionTemplates

    function One {
        param($Spec, $Doc)
        if (-not $Doc) { $Doc = $script:Defaults }
        @(ConvertTo-IntuneRequirementRules -Specs @($Spec) -Templates $Doc)[0]
    }
    function ScriptOf($Translation) { [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$Translation.Rule['scriptContent'])) }
    function Spec { param([string]$ConditionId, $Value = $null, $Cultures = $null) [pscustomobject]@{ ConditionId = $ConditionId; Value = $Value; Cultures = $Cultures } }
    function Doc { param($Condition) [pscustomobject]@{ SchemaVersion = 1; Conditions = @($Condition) } }
    function Manifest {
        param([object[]]$Requirements = @())
        [pscustomobject]@{
            AppName = 'Contoso Tool'; Architecture = 'x64'; SetupFile = 'install.bat'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Contoso'; Is64Bit = $true }
            Requirements = $Requirements
        }
    }
    function Codes($Findings) { (@($Findings | Where-Object { $_.Code -like 'Requirement*' } | ForEach-Object { '{0}:{1}' -f $_.Severity, $_.Code })) -join ',' }
}

Describe 'Requirement rules' {
    It 'turns the CPU architecture condition into a script rule that compares the processor architecture' {
        $t = One (Spec 'cpu-arch' 'ARM64')

        $t.Translated | Should -BeTrue -Because $t.Gap
        $rule = $t.Rule
        $rule['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppPowerShellScriptRule'
        $rule['ruleType'] | Should -Be 'requirement'
        $rule['displayName'] | Should -Be 'APKG - CPU Architecture'
        $rule['operationType'] | Should -Be 'string'
        $rule['operator'] | Should -Be 'equal'
        $rule['comparisonValue'] | Should -Be 'True'
        $rule['runAsAccount'] | Should -Be 'system'
        $rule['runAs32Bit'] | Should -BeFalse
        $rule['enforceSignatureCheck'] | Should -BeFalse
        $text = ScriptOf $t
        $text | Should -Match ([regex]::Escape("-Namespace 'root\cimv2' -ClassName 'Win32_Processor'"))
        $text | Should -Match ([regex]::Escape('[string]$_.Architecture'))
        $text | Should -Match ([regex]::Escape("`$values -contains '12'"))
    }

    It 'maps x64 like the ConfigMgr rule' {
        ScriptOf (One (Spec 'cpu-arch' 'x64')) | Should -Match ([regex]::Escape("`$values -contains '9'"))
    }

    It 'lists the cultures of the operating system language condition' {
        $t = One (Spec 'os-language' $null @('de-DE', 'fr-FR'))

        $t.Translated | Should -BeTrue -Because $t.Gap
        $text = ScriptOf $t
        $text | Should -Match ([regex]::Escape("`$cultures = @('de-DE', 'fr-FR')"))
        $text | Should -Match 'InstalledUICulture\.Name'
        $text | Should -Match ([regex]::Escape('$cultures -contains $name'))
        $t.Rule['comparisonValue'] | Should -Be 'True'
    }

    It 'runs the script of the VPN condition and expects the value of the spec (<Value>)' -ForEach @(
        @{ Value = $true; Expected = 'True' }, @{ Value = $false; Expected = 'False' }
    ) {
        $t = One (Spec 'vpn-connected' $Value)

        $t.Translated | Should -BeTrue -Because $t.Gap
        $t.Rule['comparisonValue'] | Should -Be $Expected
        $template = $script:Defaults.Conditions | Where-Object { $_.Id -eq 'vpn-connected' }
        $expectedText = & (Get-Module AppPackagerCommon) { param($t) Get-ConditionTemplateScriptText -Template $t } $template
        ScriptOf $t | Should -Be $expectedText
    }

    It 'gives an unknown condition (<Id>) a gap that names the problem' -ForEach @(
        @{ Id = 'no-such-condition' }, @{ Id = '' }
    ) {
        $t = One (Spec $Id 'x')

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Match 'not in the condition templates'
    }

    It 'gives a value that the template does not map a gap' {
        $t = One (Spec 'cpu-arch' 'mips')

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Match 'mips'
    }

    It 'gives an empty culture list a gap' {
        (One (Spec 'os-language' $null @())).Translated | Should -BeFalse
    }

    It 'gives a boolean condition a gap when the value is not a boolean' {
        $t = One (Spec 'vpn-connected' 'maybe')

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Match 'boolean'
    }

    It 'gives a condition of the kind <Kind> with the rule type <RuleType> no script form' -ForEach @(
        @{ Kind = 'BuiltIn'; RuleType = 'CommonValue' }
        @{ Kind = 'Wql'; RuleType = 'Boolean' }
        @{ Kind = 'Script'; RuleType = 'CommonValue' }
        @{ Kind = 'Wql'; RuleType = 'OperatingSystemLanguage' }
    ) {
        $doc = Doc ([pscustomobject]@{ Id = 'custom'; GlobalConditionName = 'Custom'; Kind = $Kind; RuleType = $RuleType; Namespace = 'root\cimv2'; Class = 'C'; Property = 'P' })

        $t = One (Spec 'custom' $true @('de-DE')) $doc

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Match 'no Intune requirement rule'
    }

    It 'gives a rule type that is not known a gap' {
        $doc = Doc ([pscustomobject]@{ Id = 'odd'; GlobalConditionName = 'Odd'; Kind = 'Wql'; RuleType = 'Range' })

        (One (Spec 'odd' '1') $doc).Gap | Should -Match 'unsupported RuleType'
    }

    It 'gives a WMI property name that is not a plain name a gap and never writes it into a script' {
        $doc = Doc ([pscustomobject]@{ Id = 'evil'; GlobalConditionName = 'Evil'; Kind = 'Wql'; RuleType = 'CommonValue'; Namespace = 'root\cimv2'; Class = 'Win32_Processor'; Property = 'Architecture; calc' })

        $t = One (Spec 'evil' '1') $doc

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Match 'not a plain name'
        $t.Rule | Should -BeNullOrEmpty
    }

    It 'gives a WMI condition without a class a gap' {
        $doc = Doc ([pscustomobject]@{ Id = 'w'; GlobalConditionName = 'W'; Kind = 'Wql'; RuleType = 'CommonValue'; Namespace = 'root\cimv2'; Class = ''; Property = 'P' })

        (One (Spec 'w' '1') $doc).Gap | Should -Match 'names no WMI namespace, class and property'
    }

    It 'escapes a quote in the value' {
        $doc = Doc ([pscustomobject]@{ Id = 'q'; GlobalConditionName = 'Q'; Kind = 'Wql'; RuleType = 'CommonValue'; Namespace = 'root\cimv2'; Class = 'Win32_Processor'; Property = 'Name' })

        ScriptOf (One (Spec 'q' "it's") $doc) | Should -Match ([regex]::Escape("`$values -contains 'it''s'"))
    }

    It 'clips a long condition name to the length of a Graph display name' {
        $doc = Doc ([pscustomobject]@{ Id = 'l'; GlobalConditionName = ('N' * 300); Kind = 'Wql'; RuleType = 'CommonValue'; Namespace = 'root\cimv2'; Class = 'Win32_Processor'; Property = 'Architecture'; Values = [pscustomobject]@{ a = '1' } })

        (One (Spec 'l' 'a') $doc).Rule['displayName'].Length | Should -Be 100
    }

    It 'gives a culture code that does not exist a gap and not an exception (<Code>)' -ForEach @(
        @{ Code = 'german' }, @{ Code = 'de DE' }
    ) {
        $t = One (Spec 'os-language' $null @($Code))

        $t.Translated | Should -BeFalse
        $t.Gap | Should -Not -BeNullOrEmpty
    }

    It 'accepts the specs as hashtables, the form of the saved deployment conditions' {
        $t = One @{ ConditionId = 'cpu-arch'; Value = 'x64' }

        $t.Translated | Should -BeTrue -Because $t.Gap
    }

    It 'publishes a script condition unsigned and not enforced without a signing policy' {
        (One (Spec 'vpn-connected' $true)).Rule['enforceSignatureCheck'] | Should -BeFalse
    }

    It 'does not turn a required requirement signature that was not produced into a gap that a drop could accept' {
        Mock Get-SigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ RequireRequirements = $true } }

        { ConvertTo-IntuneRequirementRules -Specs @((Spec 'vpn-connected' $true)) -Templates $script:Defaults -SignScripts } | Should -Throw
    }

    It 'does not sign the scripts that the program writes itself under a required policy' {
        Mock Get-SigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ RequireRequirements = $true } }

        $t = @(ConvertTo-IntuneRequirementRules -Specs @((Spec 'cpu-arch' 'x64')) -Templates $script:Defaults -SignScripts)[0]

        $t.Translated | Should -BeTrue -Because $t.Gap
        $t.Rule['enforceSignatureCheck'] | Should -BeFalse
    }
}

Describe 'Requirement specs of the Intune publish' {
    It 'reads the specs of the manifest and the specs that the run applies beside it' {
        $m = Manifest @((Spec 'cpu-arch' 'x64'))

        $specs = @(Get-IntuneRequirementSpecs -Manifest $m -Extra @((Spec 'vpn-connected' $true)))

        (@($specs | ForEach-Object { $_.ConditionId }) -join ',') | Should -Be 'cpu-arch,vpn-connected'
    }

    It 'applies a spec that both sources hold once' {
        $m = Manifest @((Spec 'cpu-arch' 'x64'))

        @(Get-IntuneRequirementSpecs -Manifest $m -Extra @((Spec 'cpu-arch' 'x64'))).Count | Should -Be 1
    }

    It 'reads the Rules of the requirements JSON of a run' {
        $json = '{"SchemaVersion":1,"Rules":[{"ConditionId":"cpu-arch","Value":"ARM64"},{"ConditionId":"os-language","Cultures":["de-DE"]}]}'

        $specs = @(ConvertFrom-RequirementsJson -Json $json)

        $specs.Count | Should -Be 2
        $specs[0].Value | Should -Be 'ARM64'
        @(ConvertFrom-RequirementsJson -Json '').Count | Should -Be 0
    }

    It 'refuses requirements JSON that is not valid instead of publishing without the rules' {
        { ConvertFrom-RequirementsJson -Json '{ not json' } | Should -Throw '*not valid JSON*'
    }
}

Describe 'Requirement findings' {
    It 'reports a translated requirement as information and blocks nothing' {
        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'cpu-arch' 'x64')))

        Codes $f | Should -Be 'Info:RequirementTranslated'
        @($f | Where-Object Severity -eq 'Blocking').Count | Should -Be 0
    }

    It 'translates a requirement that the run applies beside the manifest' {
        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest) -Requirements @((Spec 'vpn-connected' $true))

        Codes $f | Should -Be 'Info:RequirementTranslated'
    }

    It 'blocks the publish for a requirement with no Intune rule' {
        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'no-such-condition')))

        Codes $f | Should -Be 'Blocking:RequirementNotTranslatable'
        ($f | Where-Object Code -eq 'RequirementNotTranslatable').Message | Should -Match 'no-such-condition'
    }

    It 'turns the block into a review finding when the drop of that rule for that application is accepted' {
        $key = Get-IntuneRequirementDropKey -Manifest (Manifest) -ConditionId 'no-such-condition'

        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'no-such-condition'))) -AcceptedDrops @($key)

        Codes $f | Should -Be 'Review:RequirementDropped'
        @($f | Where-Object Severity -eq 'Blocking').Count | Should -Be 0
    }

    It 'does not count the accepted drop of another rule or another application (<Key>)' -ForEach @(
        @{ Key = 'AppPackager:legacy:Other-Tool/default|no-such-condition' }
        @{ Key = 'AppPackager:legacy:Contoso-Tool/default|other-condition' }
        @{ Key = 'Contoso Tool|no-such-condition' }
        @{ Key = 'no-such-condition' }
    ) {
        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'no-such-condition'))) -AcceptedDrops @($Key)

        Codes $f | Should -Be 'Blocking:RequirementNotTranslatable'
    }

    It 'compares the accepted drop without case' {
        $key = (Get-IntuneRequirementDropKey -Manifest (Manifest) -ConditionId 'no-such-condition').ToUpperInvariant()

        Codes (Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'no-such-condition'))) -AcceptedDrops @($key)) | Should -Be 'Review:RequirementDropped'
    }

    It 'prints the exact entry to accept and the name of the setting' {
        $m = Manifest @((Spec 'no-such-condition'))
        $blocking = Get-IntuneCompatibilityFindings -Manifest $m | Where-Object Code -eq 'RequirementNotTranslatable'

        $blocking.Message | Should -Match ([regex]::Escape((Get-IntuneRequirementDropKey -Manifest $m -ConditionId 'no-such-condition')))
        $blocking.Message | Should -Match 'Intune\.AcceptedRequirementDrops'
    }

    It 'carries the key that the accept step stores' {
        $m = Manifest @((Spec 'no-such-condition'))
        $blocking = Get-IntuneCompatibilityFindings -Manifest $m | Where-Object Code -eq 'RequirementNotTranslatable'

        $blocking.DropKey | Should -Be (Get-IntuneRequirementDropKey -Manifest $m -ConditionId 'no-such-condition')
        $blocking.CanAcceptDrop | Should -BeTrue
        $accepted = Get-IntuneCompatibilityFindings -Manifest $m -AcceptedDrops @($blocking.DropKey)
        Codes $accepted | Should -Be 'Review:RequirementDropped'
        @($accepted | ForEach-Object { $_.CanAcceptDrop }) | Should -Not -Contain $true
    }

    It 'keeps the key through a new version and changes it with the profile' {
        $a = Manifest; $a | Add-Member -NotePropertyName ApplicationId -NotePropertyValue 'app-1'; $a.AppName = 'Contoso Tool - 1.0 (x64)'
        $b = Manifest; $b | Add-Member -NotePropertyName ApplicationId -NotePropertyValue 'app-1'; $b.AppName = 'Contoso Tool - 2.0 (x64)'

        Get-IntuneRequirementDropKey -Manifest $a -ConditionId 'vpn-connected' | Should -Be (Get-IntuneRequirementDropKey -Manifest $b -ConditionId 'vpn-connected')
        $b | Add-Member -NotePropertyName ProfileId -NotePropertyValue 'pilot'
        Get-IntuneRequirementDropKey -Manifest $a -ConditionId 'vpn-connected' | Should -Not -Be (Get-IntuneRequirementDropKey -Manifest $b -ConditionId 'vpn-connected')
    }

    It 'blocks a culture code that does not exist' {
        $f = Get-IntuneCompatibilityFindings -Manifest (Manifest @((Spec 'os-language' $null @('de DE'))))

        Codes $f | Should -Be 'Blocking:RequirementNotTranslatable'
    }

    It 'does not let one accepted drop hide a second rule that has no rule' {
        $m = Manifest @((Spec 'no-such-condition'), (Spec 'another-missing'))
        $key = Get-IntuneRequirementDropKey -Manifest (Manifest) -ConditionId 'no-such-condition'

        Codes (Get-IntuneCompatibilityFindings -Manifest $m -AcceptedDrops @($key)) | Should -Be 'Review:RequirementDropped,Blocking:RequirementNotTranslatable'
    }

    It 'leaves a manifest with finished Graph rules alone' {
        $m = Manifest @((Spec 'no-such-condition'))
        $m | Add-Member -NotePropertyName IntuneRules -NotePropertyValue @(@{ ruleType = 'detection'; '@odata.type' = '#microsoft.graph.win32LobAppRegistryRule' })

        Codes (Get-IntuneCompatibilityFindings -Manifest $m) | Should -BeNullOrEmpty
    }

    It 'reports no requirement finding for an application without requirements' {
        Codes (Get-IntuneCompatibilityFindings -Manifest (Manifest)) | Should -BeNullOrEmpty
    }
}
