#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The minimum Windows release of an Intune app is a saved preference with an Options field and a consumer.

.DESCRIPTION
    Intune.MinimumWindowsRelease defaults to Windows10_21H2, accepts the shape Windows10_21H2 or
    Windows11_23H2, is offered in Options, Connections, Intune, and reaches the publisher
    through the publish configuration.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-IntunePublishConfigForContext', 'Test-DeploymentTargetPublishesToWsus', 'Get-OptionsSectionDefinition') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Read-PrefsWithIntune {
        param([hashtable]$Intune)
        $script:PrefsFile = Join-Path $TestDrive ('prefs-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ Intune = [pscustomobject]$Intune } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        Read-Preferences
    }
}

Describe 'Intune.MinimumWindowsRelease' {
    It 'defaults to Windows10_21H2 when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).Intune.MinimumWindowsRelease | Should -Be 'Windows10_21H2'
    }

    It 'reads <Value> from the file' -ForEach @(
        @{ Value = 'Windows11_23H2' }, @{ Value = 'Windows10_22H2' }, @{ Value = 'Windows11_25H2' }
    ) {
        (Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; MinimumWindowsRelease = $Value }).Intune.MinimumWindowsRelease | Should -Be $Value
    }

    It 'keeps the default for <Value>' -ForEach @(
        @{ Value = '' }, @{ Value = 'Windows9_1' }, @{ Value = 'Windows11_' }, @{ Value = 'latest' }, @{ Value = 'Windows10_21H2; drop' }
    ) {
        (Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; MinimumWindowsRelease = $Value }).Intune.MinimumWindowsRelease | Should -Be 'Windows10_21H2'
    }

    It 'keeps the default in a file from before the key' {
        (Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; TenantId = 'contoso.onmicrosoft.com' }).Intune.MinimumWindowsRelease | Should -Be 'Windows10_21H2'
    }

    It 'reaches the publish configuration' {
        try { $secret = ConvertTo-SecureString -String 'placeholder' -AsPlainText -Force | ConvertFrom-SecureString }
        catch { Set-ItResult -Skipped -Because 'DPAPI is not available in this session.'; return }
        $prefs = [pscustomobject]@{ Intune = [pscustomobject]@{
                DeploymentTarget = 'IntuneOnly'; PublishToIntune = $true; TenantId = 'contoso.onmicrosoft.com'
                ClientId = '11111111-2222-3333-4444-555555555555'; ClientSecretProtected = $secret; MinimumWindowsRelease = 'Windows11_23H2' } }
        $config = Get-IntunePublishConfigForContext -Prefs $prefs -Target 'IntuneOnly'
        $config.MinimumWindowsRelease | Should -Be 'Windows11_23H2'
    }

    It 'hands the release to Publish-IntuneWin32App' {
        (Get-LauncherFunction 'Invoke-PackagerIntunePublish').Extent.Text | Should -Match 'MinimumSupportedWindowsRelease\s*=\s*\[string\]\$Config\.MinimumWindowsRelease'
    }

    It 'has a field in Options, Connections, Intune' {
        $section = Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'intune' }
        $field = $section.Fields | Where-Object { $_.Id -eq 'int.minWin' }
        $field.Pref | Should -Be 'Intune.MinimumWindowsRelease'
        $field.Title | Should -Be 'Minimum Windows release'
        @($field.Items | ForEach-Object { $_.Value }) | Should -Contain 'Windows10_21H2'
    }
}
