#Requires -Version 5.1
# Window probe of the supersedence setting in Options, All Applications: the bulk buttons move the default, and the title page shows a note
# when supersedence is on and never writes the include-version choice of the operator.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-sup-options-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir

$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    Write-Host ('  ok  ' + $Message)
}

function Find-OptionsButton {
    param([Parameter(Mandatory)]$Element, [Parameter(Mandatory)][string]$Text)
    foreach ($child in @([System.Windows.LogicalTreeHelper]::GetChildren($Element))) {
        if ($child -is [System.Windows.Controls.Button] -and [string]$child.Content -eq $Text) { return $child }
        if ($child -is [System.Windows.DependencyObject]) { $hit = Find-OptionsButton -Element $child -Text $Text; if ($hit) { return $hit } }
    }
}
function Invoke-OptionsButton {
    param($Button)
    $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
}
function Set-ProbePrefs {
    param([string]$Json)
    Set-Content -LiteralPath $global:FixturePrefsPath -Value $Json -Encoding UTF8
    $script:Prefs = Read-Preferences
}

try {
    # The bulk buttons write the default through the default combo of the page.
    Set-ProbePrefs '{}'
    $page = New-OptionsTestPage -Id 'allapps'
    $mode = $page.Editors['sup.mode']
    & $assert ((& $mode.Get) -eq 'Off') 'the default is Off before the bulk action'
    $all = Find-OptionsButton -Element $page.Editors['sup.apps'].Element -Text 'All applications'
    & $assert ($null -ne $all) 'the button All applications exists'
    Invoke-OptionsButton $all
    & $assert ((& $mode.Get) -ne 'Off') 'All applications sets the default to the chosen value'
    & $page.Commit
    & $assert ([string]$script:Prefs.Supersedence.Mode -ne 'Off') 'the saved default is on after All applications'
    $none = Find-OptionsButton -Element $page.Editors['sup.apps'].Element -Text 'No applications'
    Invoke-OptionsButton $none
    & $assert ((& $mode.Get) -eq 'Off') 'No applications sets the default to Off'

    # The title page: a note while supersedence is on, and the choice of the operator is never written by it.
    foreach ($stored in $false, $true) {
        Set-ProbePrefs (('{{ "IncludeVersionInTitle": {0}, "Supersedence": {{ "Mode": "Replace" }} }}' -f $stored.ToString().ToLowerInvariant()))
        $title = New-OptionsTestPage -Id 'title'
        $inc = $title.Editors['title.include'].Focus
        $note = $title.Editors['title.lock'].Text
        & $assert ($note.Visibility -eq [System.Windows.Visibility]::Visible -and $note.Text -match 'Supersedence is on') ('the title page shows the note while supersedence is on (stored include {0})' -f $stored)
        & $assert ($inc.IsEnabled -and ($inc.IsChecked -eq $stored)) 'the include box keeps the stored choice and stays enabled'
        & $title.Commit
        & $assert ([bool]$script:Prefs.IncludeVersionInTitle -eq $stored) 'saving with supersedence on keeps the stored include choice'
        # Supersedence off again: the page returns to the same stored choice, with no note.
        $script:Prefs.Supersedence.Mode = 'Off'
        $title = New-OptionsTestPage -Id 'title'
        & $assert ($title.Editors['title.lock'].Text.Visibility -eq [System.Windows.Visibility]::Collapsed) 'the note is gone when supersedence is off'
        & $assert ($title.Editors['title.include'].Focus.IsChecked -eq $stored) 'the include box shows the choice the operator had'
        & $title.Commit
        & $assert ([bool]$script:Prefs.IncludeVersionInTitle -eq $stored) 'the stored include choice is the same after the round trip'
    }
    'PASS: the supersedence setting in Options'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
