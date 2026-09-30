#Requires -Version 5.1
# Headless probe for the Options Packager Preferences panel: the Adobe
# Acrobat Reader block enables its language boxes only for the MUI edition,
# All languages replaces the list, and Commit writes the pair.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
# Every function of the shell, published globally as the shell does for its
# closure handlers.
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    Set-Item -Path ('function:global:' + $fn.Name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$fixtureDir = Join-Path $env:TEMP ('ap-pkgprefs-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir)
$script:fixturePath = Join-Path $fixtureDir 'preferences.json'
function global:Get-PreferencesPath { $script:fixturePath }
# Files the panel reads stay in the fixture folder, never in the repository.
function global:Get-CwaSwitchesPath { Join-Path $fixtureDir 'citrix-workspace-switches.json' }
function global:Get-TvHostConfigPath { Join-Path $fixtureDir 'teamviewer-host-config.json' }
function global:Save-Preferences { param($Prefs) }
function global:Add-LogLine { param([string]$Message) }
$PackagersRoot = Join-Path $root 'Packagers'

$assert = { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw ('FAIL: ' + $Message) } }
$script:PassedChecks = 0
$ok = { param([string]$m) $script:PassedChecks += 1; Write-Host ('  ok  ' + $m) }

try {
    Set-Content -LiteralPath $script:fixturePath -Value '{ "AdobeReaderInstallOptions": { "Edition": "MUI", "Languages": ["de_DE", "fr_FR"] } }' -Encoding UTF8
    $script:Prefs = Read-Preferences
    $panel = New-PackagerPreferencesPanel
    $content = $panel.Element.FindName('panelContent')
    $edition = @($content.Children | Where-Object { $_ -is [System.Windows.Controls.StackPanel] } | ForEach-Object { $_.Children } | Where-Object { $_ -is [System.Windows.Controls.ComboBox] -and $_.Items -contains 'Multilingual (MUI)' })[0]
    $all = @($content.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Content -eq 'All languages' })[0]
    $wrap = @($content.Children | Where-Object { $_ -is [System.Windows.Controls.WrapPanel] })[0]
    & $assert ($null -ne $edition -and $null -ne $all -and $null -ne $wrap) 'the Adobe block builds its edition list, All box and language list'
    & $ok 'block present'
    & $assert ($edition.SelectedIndex -eq 1 -and $wrap.IsEnabled -and $all.IsEnabled) 'a stored MUI selection enables the language list'
    # A string Content renders de_DE as deDE (access-key underscore).
    & $assert (@($wrap.Children | Where-Object { $_.Content -isnot [System.Windows.Controls.TextBlock] }).Count -eq 0) 'every language box shows its code through a TextBlock'
    $de = @($wrap.Children | Where-Object { $_.Content.Text -eq 'de_DE' })[0]
    $fr = @($wrap.Children | Where-Object { $_.Content.Text -eq 'fr_FR' })[0]
    $ja = @($wrap.Children | Where-Object { $_.Content.Text -eq 'ja_JP' })[0]
    & $assert ($de.IsChecked -and $fr.IsChecked -and -not $ja.IsChecked) 'the stored languages are checked'
    & $ok 'stored languages shown'

    $edition.SelectedIndex = 0
    & $assert (-not $wrap.IsEnabled -and -not $all.IsEnabled) 'the English edition disables the language controls'
    $edition.SelectedIndex = 1
    $all.IsChecked = $true
    & $assert (-not $wrap.IsEnabled) 'All languages disables the list'
    & $ok 'enable rules'

    & $panel.Commit
    & $assert ($script:Prefs.AdobeReaderInstallOptions.Edition -eq 'MUI' -and @($script:Prefs.AdobeReaderInstallOptions.Languages) -eq @('All')) 'Commit saves MUI with All'
    $all.IsChecked = $false
    $ja.IsChecked = $true
    & $panel.Commit
    & $assert ((@($script:Prefs.AdobeReaderInstallOptions.Languages) -join ',') -eq 'de_DE,fr_FR,ja_JP') ('Commit saves the checked codes: ' + (@($script:Prefs.AdobeReaderInstallOptions.Languages) -join ','))
    $edition.SelectedIndex = 0
    & $panel.Commit
    & $assert ($script:Prefs.AdobeReaderInstallOptions.Edition -eq 'English') 'Commit saves the English edition'
    & $ok 'commit'

    Write-Host ''
    Write-Host ('PASS: Packager Preferences probe, {0} checks' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
