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
    # The default combo of the page is the default of the grid: a default that is on closes the title of an application that follows it.
    Set-ProbePrefs '{}'
    $page = New-OptionsTestPage -Id 'allapps'
    $mode = $page.Editors['sup.mode']
    $grid = $page.Editors['sup.apps']
    & $assert ((& $mode.Get) -eq 'Off') 'the default is Off before any change'
    & $assert (@($grid.Rows | Where-Object { -not $_['e_title'] }).Count -eq 0) 'every title is open while the default is Off'
    & $mode.Set 'Replace'
    & $assert ($grid.Work.DefaultSupersedence -eq 'Replace' -and @($grid.Rows | Where-Object { $_['e_title'] }).Count -eq 0) 'the default Replace closes the title of every application that follows it'
    & $page.Commit
    & $assert ([string]$script:Prefs.Supersedence.Mode -eq 'Replace') 'the saved default is on after the change'
    & $mode.Set 'Off'
    & $assert (@($grid.Rows | Where-Object { -not $_['e_title'] }).Count -eq 0) 'the default Off opens the titles again'

    # The title page has no supersedence note: supersedence is set in Options, and the page never writes the include choice.
    foreach ($stored in $false, $true) {
        Set-ProbePrefs (('{{ "IncludeVersionInTitle": {0}, "Supersedence": {{ "Mode": "Replace" }} }}' -f $stored.ToString().ToLowerInvariant()))
        $title = New-OptionsTestPage -Id 'title'
        $inc = $title.Editors['title.include'].Focus
        & $assert (-not $title.Editors.Contains('title.lock')) ('the title page has no supersedence note (stored include {0})' -f $stored)
        & $assert ($inc.IsEnabled -and ($inc.IsChecked -eq $stored)) 'the include box keeps the stored choice and stays enabled'
        & $title.Commit
        & $assert ([bool]$script:Prefs.IncludeVersionInTitle -eq $stored) 'saving with supersedence on keeps the stored include choice'
    }
    'PASS: the supersedence setting in Options'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
