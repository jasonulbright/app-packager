#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The chain check, the recorded installed values and the stored detection of a real deployment type.

.DESCRIPTION
    The cases are the cases of the tests of App Packager Desktop (SupersedenceChainCheckTests, InstalledValueRecorderTests and
    SupersedenceRealDetectionXmlTests). The chain check warns and never blocks. The fixtures hold the stored XML of a deployment
    type as ConfigMgr 5.00.9146 wrote it.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Fixtures = Join-Path $PSScriptRoot 'Fixtures\Supersedence'
    $script:RealFileXml = Get-Content -LiteralPath (Join-Path $script:Fixtures 'file-detection.xml') -Raw -Encoding UTF8

    $script:Method = @'
<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/AppMgmtSDK/AppMgmtSDK">
  <Settings><File LogicalName="File_1" Path="C:\Vendor" Filter="app.exe" /></Settings>
  <Rule><Expression><Operator>{OP}</Operator><Operands>
    <SettingReference SettingLogicalName="File_1" DataType="Version" Method="Value" />
    <ConstantValue Value="1.0.0" DataType="Version" />
  </Operands></Expression></Rule>
</EnhancedDetectionMethod>
'@
    function Get-TestDeploymentTypeXml {
        # The XML of a deployment type as the site stores it: the detection method is XML text inside an argument.
        param([string]$Op)
        $escaped = [System.Security.SecurityElement]::Escape($script:Method.Replace('{OP}', $Op))
        '<AppMgmtDigest><DeploymentType><Installer><DetectAction><Args><Arg Name="MethodBody" Type="String">' + $escaped + '</Arg></Args></DetectAction></Installer></DeploymentType></AppMgmtDigest>'
    }
    function New-TestManifest {
        param([string]$Op, [string]$PropertyType = 'Version')
        [pscustomobject]@{ AppName = 'New'; SoftwareVersion = '2.0.0'; Detection = [pscustomobject]@{ Type = 'File'; FilePath = 'p'; FileName = 'f'; ExpectedValue = '2.0.0'; Operator = $Op; PropertyType = $PropertyType } }
    }
    function New-TestOld { param([string[]]$Operators) [pscustomobject]@{ Title = 'Old'; Clauses = @($Operators | ForEach-Object { [pscustomobject]@{ DeploymentType = 'Old - x64'; Operator = $_ } }) } }
}

Describe 'Reading the site' {
    It 'reads the <Op> operator from the detection method text inside the deployment type XML' -ForEach @(@{ Op = 'GreaterEquals' }, @{ Op = 'Equals' }) {
        (@(Get-DetectionVersionOperator (Get-TestDeploymentTypeXml $Op)) -join ',') | Should -Be $Op
    }

    It 'gives no clause for a document that cannot be read or has no version clause' {
        @(Get-DetectionVersionOperator $null).Count | Should -Be 0
        @(Get-DetectionVersionOperator '<<<not xml').Count | Should -Be 0
        @(Get-DetectionVersionOperator ((Get-TestDeploymentTypeXml 'Equals').Replace('DataType=&quot;Version&quot;', 'DataType=&quot;String&quot;'))).Count | Should -Be 0
    }

    It 'reads the version operator of a real deployment type once' {
        (@(Get-DetectionVersionOperator $script:RealFileXml) -join ',') | Should -Be 'GreaterEquals'
    }
}

Describe 'The warnings' {
    It 'warns for Replace and uninstall with Exact when the old application still has GreaterEquals' {
        $warnings = @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode ReplaceAndUninstall -NewManifest (New-TestManifest 'IsEquals') -OldDetection (New-TestOld 'GreaterEquals'))
        $warnings.Count | Should -Be 1
        $warnings[0] | Should -Match "'Old'"
        $warnings[0] | Should -Match 'Old - x64'
        $warnings[0] | Should -Match 'uninstall will still be reported as failed'
    }

    It 'gives no warning for Replace and uninstall when the old application is already Exact' {
        @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode ReplaceAndUninstall -NewManifest (New-TestManifest 'IsEquals') -OldDetection (New-TestOld 'Equals')).Count | Should -Be 0
    }

    It 'calls an Exact old clause against a new Minimum inconsistent' {
        $warnings = @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Minimum -Mode Replace -NewManifest (New-TestManifest 'GreaterEquals') -OldDetection (New-TestOld 'Equals'))
        @($warnings | Where-Object { $_ -match 'inconsistent' -and $_ -match 'Exact version clause' }).Count | Should -BeGreaterThan 0
    }

    It 'calls a Minimum old clause against a new Exact inconsistent when nothing is uninstalled' {
        $warnings = @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode Replace -NewManifest (New-TestManifest 'IsEquals') -OldDetection (New-TestOld 'GreaterEquals'))
        @($warnings | Where-Object { $_ -match 'inconsistent' -and $_ -match 'GreaterEquals' }).Count | Should -BeGreaterThan 0
    }

    It 'warns for a new application that keeps GreaterEquals under Replace and uninstall only' {
        $kept = New-TestManifest 'GreaterEquals'
        $under = @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode ReplaceAndUninstall -NewManifest $kept -OldDetection (New-TestOld 'Equals'))
        $replace = @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode Replace -NewManifest $kept -OldDetection (New-TestOld 'Equals'))
        @($under | Where-Object { $_ -match 'keeps a GreaterEquals version clause' }).Count | Should -BeGreaterThan 0
        @($replace | Where-Object { $_ -match 'keeps a GreaterEquals' }).Count | Should -Be 0
    }

    It 'does not take an existence clause for a version clause' {
        $existence = New-TestManifest 'GreaterEquals' 'Existence'
        @(Get-SupersedenceMinimumClause -Manifest $existence).Count | Should -Be 0
        @(Get-SupersedenceChainWarning -NewTitle 'New' -NewRule Exact -Mode ReplaceAndUninstall -NewManifest $existence -OldDetection (New-TestOld)).Count | Should -Be 0
    }

    It 'walks a compound detection and the deployment type detections too' {
        $manifest = New-TestManifest 'IsEquals'
        $manifest.Detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; Clauses = @([pscustomobject]@{ Type = 'File'; FilePath = 'p'; FileName = 'a'; Operator = 'GreaterEquals' }) }
        $manifest | Add-Member -NotePropertyName DeploymentTypes -NotePropertyValue @([pscustomobject]@{ NameSuffix = 'x86'; Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'k'; ValueName = 'DisplayVersion'; Operator = 'GreaterEquals' } })
        @(Get-SupersedenceMinimumClause -Manifest $manifest).Count | Should -Be 2
    }
}

Describe 'Reading the deployment types of an old application' {
    BeforeAll {
        InModuleScope AppPackagerCommon { function script:Get-CMApplication { param($Name, [switch]$DisableWildcardHandling, $ErrorAction) } }
    }
    AfterAll {
        InModuleScope AppPackagerCommon { Remove-Item -LiteralPath 'function:script:Get-CMApplication' -ErrorAction SilentlyContinue }
    }
    It 'reads the XML of each deployment type from the XML of the application, which holds them all' {
        $xml = '<AppMgmtDigest><Application><Title>Old</Title></Application>' +
        '<DeploymentType LogicalName="a"><Title>Old - x64</Title><Marker>one</Marker></DeploymentType>' +
        '<DeploymentType LogicalName="b"><Title>Old - x86</Title><Marker>two</Marker></DeploymentType></AppMgmtDigest>'
        $body = { [pscustomobject]@{ SDMPackageXML = $xml } }.GetNewClosure()
        Mock Get-CMApplication -ModuleName AppPackagerCommon -MockWith $body
        $items = InModuleScope AppPackagerCommon { Get-SupersedenceDeploymentTypeXml -ApplicationName 'Old' }
        (@($items | ForEach-Object { $_.Name }) -join '|') | Should -Be 'Old - x64|Old - x86'
        $items[1].Xml | Should -Match 'two'
        $items[1].Xml | Should -Not -Match 'one'
    }
}

Describe 'Recording an installed value and accepting a block' {
    It 'accepts only a version text as a value: <Text>' -ForEach @(
        @{ Text = '1'; Valid = $true }, @{ Text = '1.2.3.4'; Valid = $true }, @{ Text = '2024.1-beta+5'; Valid = $true }
        @{ Text = ''; Valid = $false }, @{ Text = 'v1.2'; Valid = $false }, @{ Text = '1.2 3'; Valid = $false }, @{ Text = "1.2`n"; Valid = $false }, @{ Text = '1;calc'; Valid = $false }
    ) {
        Test-SupersedenceValidValue $Text | Should -Be $Valid
    }

    It 'stores the value under the key that the guard reads, once' {
        $settings = New-SupersedenceSettings
        $key = Set-SupersedenceInstalledValue -Settings $settings -TitleKey 'catalog:abc' -Version '1.2.3' -InstalledValue '1.2.3.4'
        $key | Should -Be (Get-SupersedenceRecordKey 'catalog:abc' '1.2.3')
        $settings.InstalledValues.$key | Should -Be '1.2.3.4'
        [void](Set-SupersedenceInstalledValue -Settings $settings -TitleKey 'catalog:abc' -Version '1.2.3' -InstalledValue '1.2.3.5')
        @($settings.InstalledValues.PSObject.Properties).Count | Should -Be 1
        $settings.InstalledValues.$key | Should -Be '1.2.3.5'
    }

    It 'replaces a value recorded in another letter case after a JSON round trip' {
        $settings = New-SupersedenceSettings
        $settings.InstalledValues = ('{"catalog:7zip|24.08":"24.08"}' | ConvertFrom-Json)
        [void](Set-SupersedenceInstalledValue -Settings $settings -TitleKey 'catalog:7Zip' -Version '24.08' -InstalledValue '24.08.0')
        $props = @($settings.InstalledValues.PSObject.Properties)
        $props.Count | Should -Be 1
        $props[0].Value | Should -Be '24.08.0'
        # A hash literal is not case sensitive, so the two keys that differ by case are added one by one.
        $recorded = [System.Collections.Hashtable]::new()
        $recorded['a|1'] = '1'
        $recorded['A|1'] = '2'
        { Get-ExactGuardVerdict -Manifest ([pscustomobject]@{ SoftwareVersion = '1' }) -Clause ([pscustomobject]@{ ExpectedValue = '1' }) -TitleKey 'a' -Recorded $recorded } | Should -Not -Throw
    }

    It 'stores an acceptance once and per version' {
        $settings = New-SupersedenceSettings
        [void](Add-SupersedenceAcceptedBlock -Settings $settings -TitleKey 'catalog:abc' -Version '1.2.3')
        [void](Add-SupersedenceAcceptedBlock -Settings $settings -TitleKey 'CATALOG:ABC' -Version '1.2.3')
        [void](Add-SupersedenceAcceptedBlock -Settings $settings -TitleKey 'catalog:abc' -Version '1.2.4')
        @($settings.AcceptedExactBlocks).Count | Should -Be 2
    }

    It 'refuses a bad version, value or title and stores nothing' {
        $settings = New-SupersedenceSettings
        { Set-SupersedenceInstalledValue -Settings $settings -TitleKey 'catalog:abc' -Version 'a|b' -InstalledValue '1.0' } | Should -Throw
        { Set-SupersedenceInstalledValue -Settings $settings -TitleKey 'catalog:abc' -Version '1.0' -InstalledValue 'newest' } | Should -Throw
        { Set-SupersedenceInstalledValue -Settings $settings -TitleKey ' ' -Version '1.0' -InstalledValue '1.0' } | Should -Throw
        { Add-SupersedenceAcceptedBlock -Settings $settings -TitleKey 'catalog:abc' -Version ' 1.0' } | Should -Throw
        @($settings.InstalledValues.Keys).Count | Should -Be 0
        @($settings.AcceptedExactBlocks).Count | Should -Be 0
    }
}
