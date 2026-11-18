#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    CheckForUpdatesAtStartup is a saved preference (default on), has an Options field and stops the launch lookup when off.
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
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Start-UpdateCheck', 'Get-OptionsSectionDefinition') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Get-SuiteInstallInfo { param($Path) $script:SuiteLookups++; $null }
    function Get-AppVersion { $script:VersionLookups++; '' }
}

Describe 'CheckForUpdatesAtStartup' {
    It 'defaults to on when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        (Read-Preferences).CheckForUpdatesAtStartup | Should -BeTrue
    }

    It 'reads <Value> from the file' -ForEach @(@{ Value = $true }, @{ Value = $false }) {
        $script:PrefsFile = Join-Path $TestDrive ('p-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ CheckForUpdatesAtStartup = $Value } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).CheckForUpdatesAtStartup | Should -Be $Value
    }

    It 'keeps on in a file from before the key' {
        $script:PrefsFile = Join-Path $TestDrive 'old.json'
        ([pscustomobject]@{ SiteCode = 'ABC' } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).CheckForUpdatesAtStartup | Should -BeTrue
    }

    It 'skips the lookup when off' {
        $script:Prefs = [pscustomobject]@{ CheckForUpdatesAtStartup = $false }
        $script:SuiteLookups = 0; $script:VersionLookups = 0
        Start-UpdateCheck
        $script:SuiteLookups | Should -Be 0
        $script:VersionLookups | Should -Be 0
    }

    It 'runs the lookup when on' {
        $script:Prefs = [pscustomobject]@{ CheckForUpdatesAtStartup = $true }
        $script:SuiteLookups = 0; $script:VersionLookups = 0
        Start-UpdateCheck
        ($script:SuiteLookups + $script:VersionLookups) | Should -BeGreaterThan 0
    }

    It 'has a field in Options, Tools' {
        $field = (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'tools' }).Fields | Where-Object { $_.Id -eq 'upd.check' }
        $field.Pref | Should -Be 'CheckForUpdatesAtStartup'
        $field.Kind | Should -Be 'Check'
    }
}
