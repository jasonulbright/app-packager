#Requires -Version 5.1
# The Deployment Conditions panel is a read-only view of the legacy
# per-app maps; the Application Workbench owns per-app authoring. This
# probe checks that stored legacy values still display here, that the
# panel no longer writes them, and that the background context map still
# reads the title policy the packager children consume.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
$t = $null; $e = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$t, [ref]$e)
if ($e) { throw ($e.Message -join '; ') }
foreach ($name in @('New-DeploymentConditionsPanel', 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-TitleModesMapForContext')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Get-ConditionTemplates { [pscustomobject]@{ Conditions = @() } }
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
    }
    Set-Content -LiteralPath $script:fixturePath -Value ($seed | ConvertTo-Json -Depth 8) -Encoding UTF8

    $script:Prefs = Read-Preferences
    $map = Get-TitleModesMapForContext
    if ($map['package-chrome'] -ne 'NoVersion' -or $map['package-firefox'] -ne 'IncludeVersion') {
        throw 'Stored title choices did not reach the background context map'
    }

    $panel = New-DeploymentConditionsPanel
    $grid = $panel.Element.FindName('dgCondApps')
    if (-not $grid.IsReadOnly) { throw 'The per-app grid is still editable' }
    if (@($grid.Columns | Where-Object { $_.Header -eq 'Application title' }).Count -ne 1) { throw 'Missing title column' }
    if (-not $panel.Element.FindName('btnOpenWorkbench')) { throw 'Missing the workbench entry point' }
    $rows = @($grid.ItemsSource)
    $chrome = @($rows | Where-Object { $_.Packager -eq 'package-chrome' })
    $firefox = @($rows | Where-Object { $_.Packager -eq 'package-firefox' })
    if ($chrome[0].TitleModeDisplay -ne 'No version' -or $firefox[0].TitleModeDisplay -ne 'Include version') {
        throw 'The migrated view does not show the stored title choices'
    }

    # Committing the panel touches only the site-level condition templates.
    $before = $script:Prefs.DeploymentConditions.Apps | ConvertTo-Json -Depth 8
    & $panel.Commit
    $after = $script:Prefs.DeploymentConditions.Apps | ConvertTo-Json -Depth 8
    if ($before -ne $after) { throw 'The read-only panel rewrote the per-app map' }

    'PASS: per-app rules display read-only, the panel writes none of them, and the title policy still reaches the background context'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
