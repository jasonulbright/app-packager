#Requires -Version 5.1
# Headless probe for the Options Adobe Acrobat Reader section: the language
# boxes are enabled only for the MUI edition, All languages replaces the list,
# and Commit writes the edition, the languages and the property options.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$script:SharedStyleRootOverride = $root

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
    . (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir
    Set-Content -LiteralPath $global:FixturePrefsPath -Value '{ "AdobeReaderInstallOptions": { "Edition": "MUI", "Languages": ["de_DE", "fr_FR"] } }' -Encoding UTF8
    $script:Prefs = Read-Preferences
    $page = New-OptionsTestPage -Id 'adobe'
    $edition = $page.Editors['adobe.edition'].Focus
    $languages = $page.Editors['adobe.languages']
    $all = $languages.Boxes['All']
    & $assert ($null -ne $edition -and $null -ne $all -and $languages.Boxes.Count -eq 27) 'the Adobe section builds its edition list, All box and 26 language boxes'
    & $ok 'section present'
    & $assert ($edition.SelectedItem.Tag -eq 'MUI' -and $all.IsEnabled -and $languages.Boxes['de_DE'].IsEnabled) 'a stored MUI selection enables the language list'
    # A string Content renders de_DE as deDE (access-key underscore).
    & $assert (@($languages.Boxes.Values | Where-Object { $_.Content -isnot [System.Windows.Controls.TextBlock] }).Count -eq 0) 'every language box shows its code through a TextBlock'
    & $assert ($languages.Boxes['de_DE'].IsChecked -and $languages.Boxes['fr_FR'].IsChecked -and -not $languages.Boxes['ja_JP'].IsChecked) 'the stored languages are checked'
    & $ok 'stored languages shown'

    $edition.SelectedIndex = 0
    & $assert (-not $all.IsEnabled -and -not $languages.Boxes['de_DE'].IsEnabled) 'the English edition disables the language controls'
    $edition.SelectedIndex = 1
    $all.IsChecked = $true
    & $assert (-not $languages.Boxes['de_DE'].IsEnabled) 'All languages disables the list'
    & $ok 'enable rules'

    & $page.Commit
    & $assert ($script:Prefs.AdobeReaderInstallOptions.Edition -eq 'MUI' -and @($script:Prefs.AdobeReaderInstallOptions.Languages) -eq @('All')) 'Commit saves MUI with All'
    $all.IsChecked = $false
    $languages.Boxes['ja_JP'].IsChecked = $true
    & $page.Commit
    & $assert ((@($script:Prefs.AdobeReaderInstallOptions.Languages) -join ',') -eq 'de_DE,fr_FR,ja_JP') ('Commit saves the checked codes: ' + (@($script:Prefs.AdobeReaderInstallOptions.Languages) -join ','))
    $edition.SelectedIndex = 0
    & $page.Commit
    & $assert ($script:Prefs.AdobeReaderInstallOptions.Edition -eq 'English') 'Commit saves the English edition'
    & $ok 'commit'

    # The Adobe property options reach the preferences.
    & $page.Editors['adobe.eula'].Set $true
    & $page.Editors['adobe.shortcut'].Set $true
    & $page.Editors['adobe.update'].Set '3'
    & $page.Commit
    $stored = $script:Prefs.AdobeReaderInstallOptions
    & $assert ($stored.AcceptEula -and $stored.DisableDesktopShortcut -and $stored.UpdateMode -eq '3' -and -not $stored.DisableChromeExtension) 'Commit saves the property options'
    & $ok 'property options'

    Write-Host ''
    Write-Host ('PASS: Adobe Acrobat Reader section probe, {0} checks' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
