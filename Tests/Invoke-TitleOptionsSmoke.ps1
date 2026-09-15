#Requires -Version 5.1
# Legacy per-app title choices stored in the preferences file must still
# reach the background context map the packager children consume.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$t = $null; $e = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$t, [ref]$e)
if ($e) { throw ($e.Message -join '; ') }
foreach ($name in @('Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-TitleModesMapForContext', 'Get-DefaultTitleModeForContext')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Get-Packagers {
    @(
        [pscustomobject]@{ Script = 'package-chrome.ps1'; Application = 'Chrome'; Vendor = 'Google'; SupportsVariants = @(); SupportsInstallModes = @() }
        [pscustomobject]@{ Script = 'package-firefox.ps1'; Application = 'Firefox'; Vendor = 'Mozilla'; SupportsVariants = @(); SupportsInstallModes = @() }
    )
}
$PackagersRoot = Join-Path $root 'Packagers'

# A preferences file carrying legacy per-app title choices.
$fixtureDir = Join-Path $env:TEMP ('ap-title-options-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir)
$script:fixturePath = Join-Path $fixtureDir 'preferences.json'
function Get-PreferencesPath { $script:fixturePath }
try {
    $seed = [pscustomobject]@{
        DeploymentConditions = [pscustomobject]@{
            Apps = [pscustomobject]@{
                'package-chrome'  = [pscustomobject]@{ Architecture = 'Any'; Languages = @(); Network = 'Any'; Split = 'None'; InstallMode = ''; TitleMode = 'NoVersion' }
                'package-firefox' = [pscustomobject]@{ Architecture = 'Any'; Languages = @(); Network = 'Any'; Split = 'None'; InstallMode = ''; TitleMode = 'IncludeVersion' }
            }
        }
        CommandOverrides = [pscustomobject]@{ Apps = [pscustomobject]@{} }
        IncludeVersionInTitle = $true
    }
    Set-Content -LiteralPath $script:fixturePath -Value ($seed | ConvertTo-Json -Depth 8) -Encoding UTF8

    $script:Prefs = Read-Preferences
    $map = Get-TitleModesMapForContext
    if ($map['package-chrome'] -ne 'NoVersion' -or $map['package-firefox'] -ne 'IncludeVersion') {
        throw 'Stored title choices did not reach the background context map'
    }
    if ((Get-DefaultTitleModeForContext) -ne 'IncludeVersion') {
        throw 'The stored include-version preference did not load'
    }

    Set-Content -LiteralPath $script:fixturePath -Value '{}' -Encoding UTF8
    $script:Prefs = Read-Preferences
    if ($script:Prefs.IncludeVersionInTitle -ne $false -or (Get-DefaultTitleModeForContext) -ne '') {
        throw 'A preferences file without the setting did not default to off'
    }

    'PASS: the stored title policy reaches the background context'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
