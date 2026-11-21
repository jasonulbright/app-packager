#Requires -Version 5.1
# Window probe of the grid of all applications in Options (Applications, All Applications): the page lists every packager,
# the combo boxes of a row write the row through their bindings, a combo box is open only where the packager offers the choice,
# the title locks while supersedence is on, the bulk buttons change the column that is chosen, and a save and a reload keep the grid.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-appgrid-options-' + [guid]::NewGuid().ToString('N'))
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
# The element of a realized row of the grid, by the automation id of the control in the column.
function Find-RowControl {
    param($Grid, $Row, [string]$Id)
    $Grid.ScrollIntoView($Row)
    $Grid.UpdateLayout()
    $container = $Grid.ItemContainerGenerator.ContainerFromItem($Row)
    if (-not $container) { throw ('FAIL: the grid did not realize the row ' + $Row.Application) }
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue($container)
    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        if ($node -is [System.Windows.FrameworkElement] -and [System.Windows.Automation.AutomationProperties]::GetAutomationId($node) -eq $Id) { return $node }
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
        for ($i = 0; $i -lt $count; $i++) { $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($node, $i)) }
    }
    throw ('FAIL: the row ' + $Row.Application + ' has no control ' + $Id)
}

$window = $null
try {
    Set-ProbePrefs '{ "DeploymentConditions": { "Apps": { "package-7zip": { "Split": "Architecture", "Network": "VpnOnly" } } }, "Supersedence": { "TitleOverrides": { "catalog:package-7zip": { "VersionRule": "Exact" } } } }'
    $page = New-OptionsTestPage -Id 'allapps'
    $editor = $page.Editors['sup.apps']
    $rows = @($editor.Rows)
    $catalog = @(Get-Packagers -Root $global:PackagersRoot)
    & $assert ($rows.Count -eq $catalog.Count -and $rows.Count -gt 300) ('the grid lists every application of the catalog (' + $rows.Count + ')')
    & $assert ($page.Editors.Contains('sup.mode') -and $page.Editors.Contains('sup.cmRule') -and $page.Editors.Contains('sup.inRule')) 'the three defaults are on the same page'

    # The page sits in a window, so that the grid realizes rows and the bindings work.
    $window = New-Object System.Windows.Window
    $window.Width = 1000; $window.Height = 900
    $window.Left = -3000; $window.Top = 0
    $window.ShowActivated = $false
    $window.Content = $page.Element
    $window.Show()
    $window.UpdateLayout()
    $grid = $editor.Grid
    & $assert ($grid.Columns.Count -eq 8) 'the grid has eight columns'
    & $assert ((@($grid.Columns | ForEach-Object { [string]$_.Header }) -join '|') -eq '|Application|Vendor|Source Type|Install Scope|Supersedence|Version in Title|Version Rule') 'the columns are in the order of the design'

    $zip = @($rows | Where-Object { $_.Base -eq 'package-7zip' })[0]
    & $assert ($zip.Arch -eq 'x64 and ARM64' -and $zip.Rule -eq 'Exact') 'a stored split and a stored version rule show in the row'
    $scopeRow = @($rows | Where-Object { $_.ScopeEnabled })[0]
    $plain = @($rows | Where-Object { -not $_.ScopeEnabled -and -not $_.ArchEnabled })[0]
    $scopeControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.scope'
    & $assert ($scopeControl.IsEnabled) ('the install scope is open for ' + $scopeRow.Application)
    $plainScope = Find-RowControl -Grid $grid -Row $plain -Id 'appgrid.scope'
    $plainArch = Find-RowControl -Grid $grid -Row $plain -Id 'appgrid.arch'
    & $assert (-not $plainScope.IsEnabled -and -not $plainArch.IsEnabled) ('the install scope and the package source type are closed for ' + $plain.Application)
    & $assert (-not [string]::IsNullOrWhiteSpace([string]$plainScope.ToolTip)) 'a closed combo box still explains itself'

    # A combo box writes its row.
    $scopeControl.SelectedItem = 'User'
    & $assert ($scopeRow.Scope -eq 'User') 'the install scope combo box writes the row'
    $supControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.supersedence'
    $titleControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.title'
    & $assert ($titleControl.IsEnabled) 'the title is open while supersedence follows the default Off'
    $titleControl.SelectedItem = 'No version'
    & $assert ($scopeRow.Title -eq 'No version') 'the title combo box writes the row'
    $supControl.SelectedItem = 'Replace'
    & $assert ($scopeRow.Supersedence -eq 'Replace') 'the supersedence combo box writes the row'
    $window.UpdateLayout()
    & $assert (-not $titleControl.IsEnabled -and [string]$titleControl.SelectedItem -eq 'Include version') 'the title is locked to Include version while supersedence is on'
    & $assert ($scopeRow.Title -eq 'No version') 'the lock leaves the stored title choice as it was'
    & $assert ([string]$titleControl.ToolTip -match 'separate applications') 'the locked combo box says why'
    & $assert ($editor.Element.Children.Count -gt 0 -and (@($editor.Element.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text -match 'Supersedence is on for 1 application' }).Count -eq 1)) 'the note counts the locked applications'
    $supControl.SelectedItem = 'Off'
    & $assert ($titleControl.IsEnabled -and [string]$titleControl.SelectedItem -eq 'No version') 'with supersedence Off the title shows the stored choice again'

    # The bulk bar.
    $column = @($editor.Element.Children[0].Children | Where-Object { $_ -is [System.Windows.Controls.ComboBox] })[0]
    $value = @($editor.Element.Children[0].Children | Where-Object { $_ -is [System.Windows.Controls.ComboBox] })[1]
    & $assert ([string]$column.SelectedItem -eq 'Supersedence') 'the bulk bar starts on the column Supersedence'
    $none = Find-OptionsButton -Element $editor.Element -Text 'No applications'
    $all = Find-OptionsButton -Element $editor.Element -Text 'All applications'
    $checked = Find-OptionsButton -Element $editor.Element -Text 'Checked applications'
    $checkAll = Find-OptionsButton -Element $editor.Element -Text 'Check all'
    $checkNone = Find-OptionsButton -Element $editor.Element -Text 'Check none'
    foreach ($b in $none, $all, $checked, $checkAll, $checkNone) { & $assert ($b -and -not [string]::IsNullOrWhiteSpace([string]$b.ToolTip)) ('the button ' + $b.Content + ' has a tooltip') }
    Invoke-OptionsButton $checked
    & $assert ($editor.Summary.Text -match 'Check at least one') 'Checked applications with nothing checked says so'
    Invoke-OptionsButton $checkAll
    & $assert (@($rows | Where-Object { $_.Checked }).Count -eq $rows.Count) 'Check all checks every row'
    Invoke-OptionsButton $checkNone
    & $assert (@($rows | Where-Object { $_.Checked }).Count -eq 0) 'Check none clears every row'

    $column.SelectedItem = 'Install Scope'
    & $assert ([string]$value.SelectedItem -eq 'System') 'the value list follows the column'
    $value.SelectedItem = 'User'
    $scopeRow.Scope = 'Packager default'
    $zip.Checked = $true; $scopeRow.Checked = $true
    Invoke-OptionsButton $checked
    & $assert ($scopeRow.Scope -eq 'User' -and @($rows | Where-Object { $_.Scope -eq 'User' }).Count -eq 1) 'Checked applications writes the scope of the checked rows that offer it'
    Invoke-OptionsButton $all
    & $assert (@($rows | Where-Object { $_.ScopeEnabled -and $_.Scope -ne 'User' }).Count -eq 0 -and @($rows | Where-Object { -not $_.ScopeEnabled -and $_.Scope -ne 'Packager default' }).Count -eq 0) 'All applications writes the scope of every row that offers it'
    Invoke-OptionsButton $none
    & $assert (@($rows | Where-Object { $_.Scope -ne 'Packager default' }).Count -eq 0) 'No applications resets the column'

    $column.SelectedItem = 'Supersedence'
    $value.SelectedItem = 'Replace'
    Invoke-OptionsButton $all
    & $assert ([string](& $page.Editors['sup.mode'].Get) -eq 'Replace') 'All applications sets the default of supersedence on the page'
    & $assert (@($rows | Where-Object { $_.Supersedence -ne 'Follow the default' }).Count -eq 0) 'All applications removes the own values'
    & $assert (@($rows | Where-Object { -not $_.TitleLocked }).Count -eq 0) 'every title is locked while the default is on'
    Invoke-OptionsButton $none
    & $assert ([string](& $page.Editors['sup.mode'].Get) -eq 'Off' -and @($rows | Where-Object { $_.TitleLocked }).Count -eq 0) 'No applications sets the default back to Off and unlocks the titles'

    $column.SelectedItem = 'Version Rule'
    $value.SelectedItem = 'Exact'
    Invoke-OptionsButton $all
    & $assert ([string](& $page.Editors['sup.cmRule'].Get) -eq 'Exact' -and [string](& $page.Editors['sup.inRule'].Get) -eq 'Exact') 'All applications sets both defaults of the version rule'
    Invoke-OptionsButton $none
    & $assert ([string](& $page.Editors['sup.cmRule'].Get) -eq 'Minimum') 'No applications sets the version rule back to Minimum'

    # Save and reload.
    $zip.Arch = 'Packager default'
    $scopeRow.Scope = 'System'; $scopeRow.Supersedence = 'Replace and uninstall'; $scopeRow.Rule = 'Exact'; $scopeRow.Title = 'Include version'
    & $page.Commit
    Save-Preferences $script:Prefs
    $script:Prefs = Read-Preferences
    $entry = $script:Prefs.DeploymentConditions.Apps.($scopeRow.Base)
    & $assert ($entry.InstallMode -eq 'AllUsers' -and $entry.TitleMode -eq 'IncludeVersion') 'the saved file holds the scope and the title of the row'
    $own = $script:Prefs.Supersedence.TitleOverrides.($scopeRow.Id)
    & $assert ($own.Mode -eq 'Replace and uninstall' -and $own.VersionRule -eq 'Exact') 'the saved file holds the supersedence and the version rule of the row'
    $zipEntry = $script:Prefs.DeploymentConditions.Apps.'package-7zip'
    & $assert ($zipEntry.Split -eq 'None' -and $zipEntry.Network -eq 'VpnOnly') 'the saved file keeps the condition that the grid does not own'
    $again = New-OptionsTestPage -Id 'allapps'
    $back = @($again.Editors['sup.apps'].Rows | Where-Object { $_.Base -eq $scopeRow.Base })[0]
    & $assert ($back.Scope -eq 'System' -and $back.Supersedence -eq 'Replace and uninstall' -and $back.Rule -eq 'Exact' -and $back.Title -eq 'Include version') 'a new page shows what was saved'
    'PASS: the grid of all applications in Options'
}
finally {
    if ($window) { $window.Close() }
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
