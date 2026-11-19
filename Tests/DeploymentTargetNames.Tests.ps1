#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The deployment target names ConfigMgr, ConfigMgrAndIntune and ConfigMgrAndWSUS are read and written; the old names still load.

.DESCRIPTION
    The preferences file, the run snapshot and the publication record carry the current names. A file from an
    earlier build, or a file that the desktop app wrote, loads with either spelling.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-DeploymentTargetNames', 'Get-KnownDeploymentTarget', 'ConvertTo-DeploymentTargetName',
        'Test-DeploymentTargetSkipsSite', 'Test-DeploymentTargetPublishesToWsus', 'Test-DeploymentTargetPublishesToIntune', 'Get-OptionsSectionDefinition') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Read-PrefsFromJson {
        param([string]$Json)
        $script:PrefsFile = Join-Path $TestDrive ('prefs-' + [guid]::NewGuid().ToString('N') + '.json')
        Set-Content -LiteralPath $script:PrefsFile -Value $Json -Encoding ASCII
        Read-Preferences
    }
}

Describe 'Deployment target in the preferences file' {
    It 'defaults to ConfigMgr when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).Intune.DeploymentTarget | Should -BeExactly 'ConfigMgr'
    }

    It 'loads <Saved> as <Expected>' -TestCases @(
        @{ Saved = 'ConfigMgr'; Expected = 'ConfigMgr' }
        @{ Saved = 'ConfigMgrAndIntune'; Expected = 'ConfigMgrAndIntune' }
        @{ Saved = 'IntuneOnly'; Expected = 'IntuneOnly' }
        @{ Saved = 'ConfigMgrAndWSUS'; Expected = 'ConfigMgrAndWSUS' }
        @{ Saved = 'WSUSOnly'; Expected = 'WSUSOnly' }
        @{ Saved = 'IntuneAndWSUS'; Expected = 'IntuneAndWSUS' }
        @{ Saved = 'ConfigMgrAndIntuneAndWSUS'; Expected = 'ConfigMgrAndIntuneAndWSUS' }
        @{ Saved = 'MECM'; Expected = 'ConfigMgr' }
        @{ Saved = 'MECMAndIntune'; Expected = 'ConfigMgrAndIntune' }
        @{ Saved = 'MECMAndWSUS'; Expected = 'ConfigMgrAndWSUS' }
        @{ Saved = 'mecmandwsus'; Expected = 'ConfigMgrAndWSUS' }
    ) {
        param($Saved, $Expected)
        $prefs = Read-PrefsFromJson ('{ "Intune": { "DeploymentTarget": "' + $Saved + '" } }')
        $prefs.Intune.DeploymentTarget | Should -BeExactly $Expected
    }

    It 'follows the saved default destinations for a name that is no target' {
        $prefs = Read-PrefsFromJson '{ "AppFlow": { "DefaultDestinations": { "ConfigMgr": false, "WSUS": true, "Intune": false } }, "Intune": { "DeploymentTarget": "Nonsense" } }'
        $prefs.Intune.DeploymentTarget | Should -BeExactly 'WSUSOnly'
    }

    It 'does not select ConfigMgr for an empty or unknown name when no destination is saved' {
        foreach ($value in '', 'Nonsense') {
            $prefs = Read-PrefsFromJson ('{ "AppFlow": { "DefaultDestinations": { "ConfigMgr": false, "WSUS": false, "Intune": false } }, "Intune": { "DeploymentTarget": "' + $value + '" } }')
            $prefs.Intune.DeploymentTarget | Should -BeExactly ''
        }
    }

    It 'selects no default destination for a name that is no target in a file without default destinations' {
        $prefs = Read-PrefsFromJson '{ "AppFlow": { "Tracked": [] }, "Intune": { "DeploymentTarget": "Nonsense" } }'
        $prefs.AppFlow.DefaultDestinations.ConfigMgr | Should -BeFalse
        $prefs.AppFlow.DefaultDestinations.WSUS | Should -BeFalse
        $prefs.AppFlow.DefaultDestinations.Intune | Should -BeFalse
        $prefs.Intune.DeploymentTarget | Should -BeExactly ''
    }

    It 'maps the old toggle PublishToIntune to ConfigMgrAndIntune' {
        (Read-PrefsFromJson '{ "Intune": { "PublishToIntune": true } }').Intune.DeploymentTarget | Should -BeExactly 'ConfigMgrAndIntune'
    }

    It 'writes the current name when a file with an old name is saved again' {
        $prefs = Read-PrefsFromJson '{ "Intune": { "DeploymentTarget": "MECMAndWSUS" } }'
        $json = $prefs | ConvertTo-Json -Depth 5
        $json | Should -Match '"DeploymentTarget":\s*"ConfigMgrAndWSUS"'
        $json | Should -Not -Match 'MECMAnd'
    }

    It 'derives the destinations of a file without default destinations from <Saved>' -TestCases @(
        @{ Saved = 'ConfigMgr'; Cm = $true; Wsus = $false; Intune = $false }
        @{ Saved = 'ConfigMgrAndIntune'; Cm = $true; Wsus = $false; Intune = $true }
        @{ Saved = 'IntuneOnly'; Cm = $false; Wsus = $false; Intune = $true }
        @{ Saved = 'ConfigMgrAndWSUS'; Cm = $true; Wsus = $true; Intune = $false }
        @{ Saved = 'WSUSOnly'; Cm = $false; Wsus = $true; Intune = $false }
        @{ Saved = 'IntuneAndWSUS'; Cm = $false; Wsus = $true; Intune = $true }
        @{ Saved = 'ConfigMgrAndIntuneAndWSUS'; Cm = $true; Wsus = $true; Intune = $true }
        @{ Saved = 'MECM'; Cm = $true; Wsus = $false; Intune = $false }
        @{ Saved = 'MECMAndIntune'; Cm = $true; Wsus = $false; Intune = $true }
        @{ Saved = 'MECMAndWSUS'; Cm = $true; Wsus = $true; Intune = $false }
    ) {
        param($Saved, $Cm, $Wsus, $Intune)
        $prefs = Read-PrefsFromJson ('{ "AppFlow": { "Tracked": [] }, "Intune": { "DeploymentTarget": "' + $Saved + '" } }')
        $prefs.AppFlow.DefaultDestinations.ConfigMgr | Should -Be $Cm
        $prefs.AppFlow.DefaultDestinations.WSUS | Should -Be $Wsus
        $prefs.AppFlow.DefaultDestinations.Intune | Should -Be $Intune
    }

    It 'loads a settings file that the desktop app wrote with the new names' {
        $apd = '{ "SiteCode": "XYZ", "Intune": { "DeploymentTarget": "ConfigMgrAndWSUS", "TenantId": "", "ClientId": "", "ClientSecretProtected": "" }, ' +
            '"AppFlow": { "DefaultDestinations": { "ConfigMgr": true, "WSUS": true, "Intune": false } } }'
        $prefs = Read-PrefsFromJson $apd
        $prefs.SiteCode | Should -Be 'XYZ'
        $prefs.Intune.DeploymentTarget | Should -BeExactly 'ConfigMgrAndWSUS'
        $prefs.Intune.PublishToIntune | Should -BeFalse
        $prefs.AppFlow.DefaultDestinations.ConfigMgr | Should -BeTrue
        $prefs.AppFlow.DefaultDestinations.WSUS | Should -BeTrue
    }

    It 'does not publish to Intune from a WSUS target with a stale Intune toggle, for either spelling' {
        foreach ($name in 'ConfigMgrAndWSUS', 'MECMAndWSUS') {
            (Read-PrefsFromJson ('{ "Intune": { "DeploymentTarget": "' + $name + '", "PublishToIntune": true } }')).Intune.PublishToIntune | Should -BeFalse
        }
    }
}

Describe 'Deployment target helpers' {
    It 'lists the current names' {
        @(Get-DeploymentTargetNames) | Should -Be @('ConfigMgr', 'ConfigMgrAndIntune', 'IntuneOnly', 'ConfigMgrAndWSUS', 'WSUSOnly', 'IntuneAndWSUS', 'ConfigMgrAndIntuneAndWSUS')
    }

    It 'maps <Value> to <Expected> for a run' -TestCases @(
        @{ Value = 'MECM'; Expected = 'ConfigMgr' }
        @{ Value = 'MECMAndIntune'; Expected = 'ConfigMgrAndIntune' }
        @{ Value = 'MECMAndWSUS'; Expected = 'ConfigMgrAndWSUS' }
        @{ Value = 'ConfigMgrAndWSUS'; Expected = 'ConfigMgrAndWSUS' }
        @{ Value = 'WSUSOnly'; Expected = 'WSUSOnly' }
        @{ Value = 'IntuneAndWSUS'; Expected = 'IntuneAndWSUS' }
    ) {
        param($Value, $Expected)
        Get-KnownDeploymentTarget -Value $Value | Should -BeExactly $Expected
    }

    It 'stops a run for the value <Value> instead of choosing ConfigMgr' -TestCases @(@{ Value = '' }, @{ Value = 'ContentOnly' }, @{ Value = 'x' }) {
        param($Value)
        { Get-KnownDeploymentTarget -Value $Value } | Should -Throw '*is not valid. Use one of: ConfigMgr, ConfigMgrAndIntune*'
    }

    It 'names the nearest single target with the current names' {
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $true; WSUS = $false; Intune = $false }) | Should -BeExactly 'ConfigMgr'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $true; WSUS = $false; Intune = $true }) | Should -BeExactly 'ConfigMgrAndIntune'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $false }) | Should -BeExactly 'ConfigMgrAndWSUS'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $false; WSUS = $true; Intune = $false }) | Should -BeExactly 'WSUSOnly'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $true }) | Should -BeExactly 'IntuneOnly'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $false; WSUS = $true; Intune = $true }) | Should -BeExactly 'IntuneAndWSUS'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $true }) | Should -BeExactly 'ConfigMgrAndIntuneAndWSUS'
        ConvertTo-DeploymentTargetName -Destinations ([pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $false }) | Should -BeExactly ''
    }

    It 'treats <Value> as skipping the site: <Skips>, publishing to WSUS: <Wsus>' -TestCases @(
        @{ Value = 'ConfigMgr'; Skips = $false; Wsus = $false }
        @{ Value = 'MECM'; Skips = $false; Wsus = $false }
        @{ Value = 'ConfigMgrAndWSUS'; Skips = $false; Wsus = $true }
        @{ Value = 'MECMAndWSUS'; Skips = $false; Wsus = $true }
        @{ Value = 'ConfigMgrAndIntune'; Skips = $false; Wsus = $false }
        @{ Value = 'MECMAndIntune'; Skips = $false; Wsus = $false }
        @{ Value = 'IntuneOnly'; Skips = $true; Wsus = $false }
        @{ Value = 'WSUSOnly'; Skips = $true; Wsus = $true }
        @{ Value = 'IntuneAndWSUS'; Skips = $true; Wsus = $true }
        @{ Value = 'ConfigMgrAndIntuneAndWSUS'; Skips = $false; Wsus = $true }
    ) {
        param($Value, $Skips, $Wsus)
        Test-DeploymentTargetSkipsSite -DeploymentTarget $Value | Should -Be $Skips
        Test-DeploymentTargetPublishesToWsus -DeploymentTarget $Value | Should -Be $Wsus
    }
}
