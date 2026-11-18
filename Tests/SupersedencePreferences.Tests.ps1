#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The supersedence settings are saved preferences with Options fields, and the publish consumers read them.

.DESCRIPTION
    Supersedence.Mode defaults to Off, ConfigMgrVersionRule and IntuneVersionRule to Minimum. A word that is not allowed keeps the
    default. TitleOverrides, InstalledValues and AcceptedExactBlocks have no field: the settings file and the command line edit
    them. The ConfigMgr publish and the Intune publish read the settings, and the command line outranks them.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-OptionsSectionDefinition') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Read-PrefsWith {
        param($Supersedence)
        $script:PrefsFile = Join-Path $TestDrive ('p-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ Supersedence = $Supersedence } | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).Supersedence
    }
}

Describe 'The Supersedence preferences' {
    It 'default to Off, Minimum and Minimum when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        $s = (Read-Preferences).Supersedence
        $s.Mode | Should -Be 'Off'
        $s.ConfigMgrVersionRule | Should -Be 'Minimum'
        $s.IntuneVersionRule | Should -Be 'Minimum'
        @($s.TitleOverrides.PSObject.Properties).Count | Should -Be 0
        @($s.InstalledValues.PSObject.Properties).Count | Should -Be 0
        @($s.AcceptedExactBlocks).Count | Should -Be 0
    }

    It 'keep the defaults in a file from before the keys' {
        $script:PrefsFile = Join-Path $TestDrive 'old.json'
        ([pscustomobject]@{ SiteCode = 'ABC' } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).Supersedence.Mode | Should -Be 'Off'
    }

    It 'read the mode <Written> as <Expected>' -ForEach @(
        @{ Written = 'Replace'; Expected = 'Replace' }, @{ Written = 'replace-uninstall'; Expected = 'Replace and uninstall' }
        @{ Written = 'Replace and uninstall'; Expected = 'Replace and uninstall' }, @{ Written = 'ReplaceAndUninstall'; Expected = 'Replace and uninstall' }, @{ Written = 'OFF'; Expected = 'Off' }
    ) {
        (Read-PrefsWith @{ Mode = $Written }).Mode | Should -Be $Expected
    }

    It 'read the rule <Written> as <Expected>' -ForEach @(@{ Written = 'exact'; Expected = 'Exact' }, @{ Written = 'MINIMUM'; Expected = 'Minimum' }) {
        $s = Read-PrefsWith @{ ConfigMgrVersionRule = $Written; IntuneVersionRule = $Written }
        $s.ConfigMgrVersionRule | Should -Be $Expected
        $s.IntuneVersionRule | Should -Be $Expected
    }

    It 'keep the default for a word that is not allowed: <Written>' -ForEach @(@{ Written = '' }, @{ Written = 'Sometimes' }, @{ Written = 'Exact; drop' }) {
        $s = Read-PrefsWith @{ Mode = $Written; ConfigMgrVersionRule = $Written; IntuneVersionRule = $Written }
        $s.Mode | Should -Be 'Off'
        $s.ConfigMgrVersionRule | Should -Be 'Minimum'
        $s.IntuneVersionRule | Should -Be 'Minimum'
    }

    It 'keep the overrides with allowed words only' {
        $s = Read-PrefsWith @{ TitleOverrides = @{ 'catalog:package-a' = @{ VersionRule = 'Exact'; Mode = 'Replace' }; 'catalog:package-b' = @{ VersionRule = 'nope' }; 'catalog:package-c' = @{ Mode = 'replace-uninstall' } } }
        $s.TitleOverrides.'catalog:package-a'.VersionRule | Should -Be 'Exact'
        $s.TitleOverrides.'catalog:package-a'.Mode | Should -Be 'Replace'
        $s.TitleOverrides.PSObject.Properties['catalog:package-b'] | Should -BeNullOrEmpty
        $s.TitleOverrides.'catalog:package-c'.Mode | Should -Be 'Replace and uninstall'
    }

    It 'keep the installed values and the accepted blocks of the right shape' {
        $s = Read-PrefsWith @{ InstalledValues = @{ 'catalog:package-a|1.2.3' = '1.2.3.4'; 'no-separator' = '1.0'; 'catalog:package-b|1.0' = 'newest' }; AcceptedExactBlocks = @('catalog:package-a|1.2.3', 'bad', 'catalog:package-a|1.2.3') }
        @($s.InstalledValues.PSObject.Properties).Count | Should -Be 1
        $s.InstalledValues.'catalog:package-a|1.2.3' | Should -Be '1.2.3.4'
        @($s.AcceptedExactBlocks).Count | Should -Be 1
    }

    It 'are read by the publish from the preferences file with the words as written' {
        $path = Join-Path $TestDrive 'publish.json'
        ([pscustomobject]@{ Supersedence = [pscustomobject]@{ Mode = 'Maybe'; ConfigMgrVersionRule = 'Exact' } } | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding ASCII
        $s = Get-SupersedenceSettings -PreferencesPath $path
        $s.Mode | Should -Be 'Maybe'
        $s.ConfigMgrVersionRule | Should -Be 'Exact'
        $s.IntuneVersionRule | Should -Be 'Minimum'
        (Get-SupersedenceSettings -PreferencesPath (Join-Path $TestDrive 'absent.json')).Mode | Should -Be 'Off'
    }
}

Describe 'The command line level' {
    It 'is empty without the variables and outranks the other levels with them' {
        $ruleBefore = $env:APP_PACKAGER_VERSION_RULE; $modeBefore = $env:APP_PACKAGER_SUPERSEDENCE
        try {
            $env:APP_PACKAGER_VERSION_RULE = ''; $env:APP_PACKAGER_SUPERSEDENCE = ''
            Get-SupersedenceCommandLine | Should -BeNullOrEmpty
            $env:APP_PACKAGER_VERSION_RULE = 'Exact'; $env:APP_PACKAGER_SUPERSEDENCE = 'Replace'
            $cli = Get-SupersedenceCommandLine
            $settings = New-SupersedenceSettings
            $resolved = Resolve-Supersedence -Settings $settings -ApplicationId 'catalog:package-a' -CommandLine $cli
            $resolved.ConfigMgrRule.Source | Should -Be 'CommandLine'
            $resolved.Mode.Value | Should -Be 'Replace'
        }
        finally { $env:APP_PACKAGER_VERSION_RULE = $ruleBefore; $env:APP_PACKAGER_SUPERSEDENCE = $modeBefore }
    }

    It 'has the parameters on Invoke-AppPackagerBuild.ps1' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'Invoke-AppPackagerBuild.ps1'), [ref]$null, [ref]$null)
        $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($n in 'VersionRule', 'Supersedence', 'RewriteOldVersions', 'Yes', 'RecordInstalledValue', 'AcceptExactBlock', 'Architecture') { $names | Should -Contain $n }
    }

    It 'sets the environment of the child from the parameters' {
        $text = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Invoke-AppPackagerBuild.ps1') -Raw
        $text | Should -Match "APP_PACKAGER_VERSION_RULE'\] = \`$VersionRule"
        $text | Should -Match "APP_PACKAGER_SUPERSEDENCE'\] = \`$supersedenceModeToken"
    }
}

Describe 'Options' {
    It 'has the three fields in Options, Packaging, Supersedence' {
        $section = Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'supersedence' }
        $section.Title | Should -Be 'Supersedence'
        $section.Parent | Should -Be 'packaging'
        $prefs = @($section.Fields | ForEach-Object { $_.Pref })
        $prefs | Should -Contain 'Supersedence.ConfigMgrVersionRule'
        $prefs | Should -Contain 'Supersedence.IntuneVersionRule'
        $prefs | Should -Contain 'Supersedence.Mode'
        @(($section.Fields | Where-Object { $_.Pref -eq 'Supersedence.Mode' }).Items | ForEach-Object { $_.Value }) | Should -Be @('Off', 'Replace', 'Replace and uninstall')
    }
}

Describe 'The publish consumers' {
    It 'reads the plan in the ConfigMgr publish before the specs are resolved and passes the manifest of the plan on' {
        $text = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $text | Should -Match '\$Manifest = \$supersedencePlan\.ConfigMgrManifest'
        $text.IndexOf('$Manifest = $supersedencePlan.ConfigMgrManifest') | Should -BeLessThan $text.IndexOf('Get-ManifestDeploymentTypeSpecs')
        $text | Should -Match '-PublishesToConfigMgr \$true -PublishesToIntune \$false'
    }

    It 'runs the chain check before the publish only when the mode is on and the marker exists' {
        $text = (Get-Command New-MECMApplicationFromManifest).ScriptBlock.ToString()
        $text | Should -Match "\`$supersedencePlan\.Mode -ne 'Off' -and -not \[string\]::IsNullOrWhiteSpace\(\[string\]\`$supersedencePlan\.Marker\)"
        $text.IndexOf('Invoke-SupersedenceChainCheck') | Should -BeLessThan $text.IndexOf('Get-CMApplication -Name $appName -DisableWildcardHandling')
    }

    It 'reads the Intune plan before the Intune publish and refuses with the text of the plan' {
        $text = (Get-LauncherFunction 'Invoke-PackagerIntunePublish').Extent.Text
        $text | Should -Match '-PublishesToConfigMgr \$false -PublishesToIntune \$true'
        $text | Should -Match 'if \(\$supersedencePlan\.Refusal\) \{ throw \$supersedencePlan\.Refusal \}'
        $text | Should -Match '\$manifest = \$supersedencePlan\.IntuneManifest'
        $text.IndexOf('$manifest = $supersedencePlan.IntuneManifest') | Should -BeLessThan $text.IndexOf('Publish-IntuneWin32App')
    }

    It 'maps the Intune operator that the version rule writes' {
        $manifest = [pscustomobject]@{ Detection = [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Vendor'; FileName = 'app.exe'; PropertyType = 'Version'; ExpectedValue = '1.0.0'; Operator = 'IsEquals'; Is64Bit = $true } }
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest $manifest -Profile $null)
        $rules.Count | Should -Be 1
        $rules[0].operator | Should -Be 'equal'
        $rules[0].operationType | Should -Be 'version'
    }
}
