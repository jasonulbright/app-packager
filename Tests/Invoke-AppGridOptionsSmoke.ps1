#Requires -Version 5.1
# Window probe of the grid of all applications in Options (Applications, All Applications): the page lists every packager, the column
# chooser shows and hides columns, a cell writes the working copy through its binding, a closed cell says why, the title locks while
# supersedence is on, the search and the bulk bar change the column and the scope that are chosen, and a save and a reload keep the grid.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-appgrid-options-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir

$env:APP_PACKAGER_WORKBENCH_ROOT = Join-Path $fixtureDir 'workbench'
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    Write-Host ('  ok  ' + $Message)
}

function Find-ById {
    param([Parameter(Mandatory)]$Element, [Parameter(Mandatory)][string]$Id)
    foreach ($child in @([System.Windows.LogicalTreeHelper]::GetChildren($Element))) {
        if ($child -is [System.Windows.DependencyObject]) {
            if ([System.Windows.Automation.AutomationProperties]::GetAutomationId($child) -eq $Id) { return $child }
            $hit = Find-ById -Element $child -Id $Id
            if ($hit) { return $hit }
        }
    }
}
function Invoke-OptionsButton {
    param($Button)
    $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
}
function Invoke-MenuItem {
    param($Item, [bool]$Checked)
    $Item.IsChecked = $Checked
    $Item.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.MenuItem]::ClickEvent)))
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
function Update-TextBinding {
    param($Box)
    [System.Windows.Data.BindingOperations]::GetBindingExpression($Box, [System.Windows.Controls.TextBox]::TextProperty).UpdateSource()
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

    $window = New-Object System.Windows.Window
    $window.Width = 1400; $window.Height = 1000
    $window.Left = -3000; $window.Top = 0
    $window.ShowActivated = $false
    $window.Content = $page.Element
    $window.Show()
    $window.UpdateLayout()
    $grid = $editor.Grid
    & $assert ($grid.Columns.Count -eq 26) 'the grid has three fixed columns and one column for each of the 23 settings'
    $visible = @($grid.Columns | Where-Object { $_.Visibility -eq 'Visible' } | ForEach-Object { [string]$_.Header })
    & $assert (($visible -join '|') -eq '|Application|Vendor|Package Source|Install Scope|Supersedence|Version in Title|Version Rule') 'the default columns are the five settings of the design'
    foreach ($c in $grid.Columns) { if ($c.Header -and $c.Header -notin 'Application', 'Vendor') { & $assert ($c.MinWidth -ge 10 * ([string]$c.Header).Length) ('the column ' + $c.Header + ' holds its header') } }

    # The column chooser.
    $columnsButton = Find-ById -Element $editor.Element -Id 'sup.apps.columns'
    & $assert ($columnsButton -and $editor.Menu.Items.Count -eq 23) 'the Columns button has a menu with one item for each setting'
    $networkItem = @($editor.Menu.Items | Where-Object { $_.Header -eq 'Network' })[0]
    & $assert (-not $networkItem.IsChecked -and -not [string]::IsNullOrWhiteSpace([string]$networkItem.ToolTip)) 'a column that is not shown is not checked, and its item has a tooltip'
    Invoke-MenuItem $networkItem $true
    & $assert (@($grid.Columns | Where-Object { $_.Header -eq 'Network' })[0].Visibility -eq 'Visible') 'checking an item shows the column'
    Invoke-MenuItem $networkItem $false
    & $assert (@($grid.Columns | Where-Object { $_.Header -eq 'Network' })[0].Visibility -eq 'Collapsed') 'clearing an item hides the column'

    $zip = @($rows | Where-Object { $_.Base -eq 'package-7zip' })[0]
    & $assert ($zip['v_source'] -eq 'x64 and ARM64' -and $zip['v_rule'] -eq 'Exact') 'a stored split and a stored version rule show in the row'
    $scopeRow = @($rows | Where-Object { $_['e_scope'] })[0]
    $plain = @($rows | Where-Object { -not $_['e_scope'] -and -not $_['e_source'] })[0]
    $scopeControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.scope'
    & $assert ($scopeControl.IsEnabled) ('the install scope is open for ' + $scopeRow.Application)
    $plainScope = Find-RowControl -Grid $grid -Row $plain -Id 'appgrid.scope'
    & $assert (-not $plainScope.IsEnabled -and [string]$plainScope.ToolTip -match 'one install scope') ('a closed cell is greyed and its tooltip gives the reason (' + $plain.Application + ')')

    # A combo box writes the working copy.
    $scopeControl.SelectedItem = 'User'
    & $assert ($scopeRow['v_scope'] -eq 'User' -and [string]$editor.Work.Apps[$scopeRow.Base]['InstallMode'] -eq 'CurrentUser') 'the install scope combo box writes the working copy'
    $titleControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.title'
    $supControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.supersedence'
    $titleControl.SelectedItem = 'No version'
    & $assert ($scopeRow['v_title'] -eq 'No version') 'the title combo box writes the row'
    $supControl.SelectedItem = 'Replace'
    $window.UpdateLayout()
    $titleControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.title'
    & $assert (-not $titleControl.IsEnabled -and [string]$titleControl.SelectedItem -eq 'Include version' -and [string]$titleControl.ToolTip -match 'separate applications') 'the title is closed and shows Include version while supersedence is on'
    & $assert ([string]$editor.Work.Apps[$scopeRow.Base]['TitleMode'] -eq 'NoVersion') 'the lock leaves the stored title choice as it was'
    & $assert (@($editor.Element.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text -match 'Supersedence is on for 1 application' }).Count -eq 1) 'the note counts the locked applications'
    $supControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.supersedence'
    $supControl.SelectedItem = 'Off'
    $titleControl = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.title'
    & $assert ($titleControl.IsEnabled -and [string]$titleControl.SelectedItem -eq 'No version') 'with supersedence Off the title shows the stored choice again'

    # A text cell writes on LostFocus.
    Invoke-MenuItem (@($editor.Menu.Items | Where-Object { $_.Header -eq 'Check Interval Days' })[0]) $true
    $window.UpdateLayout()
    $cadence = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.cadence'
    $cadence.Text = '45'
    Update-TextBinding $cadence
    & $assert ($editor.Work.Cadence[$scopeRow.Base] -eq 45) 'the interval text box writes the working copy'
    $cadence = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.cadence'
    $cadence.Text = 'abc'
    Update-TextBinding $cadence
    & $assert ($editor.Work.Cadence[$scopeRow.Base] -eq 45 -and $scopeRow['v_cadence'] -eq '45') 'a bad interval is not written and the cell shows the stored value again'

    # The search and the bulk bar.
    $search = Find-ById -Element $editor.Element -Id 'sup.apps.search'
    $column = Find-ById -Element $editor.Element -Id 'sup.apps.column'
    $value = Find-ById -Element $editor.Element -Id 'sup.apps.value'
    $scope = Find-ById -Element $editor.Element -Id 'sup.apps.scope'
    $apply = Find-ById -Element $editor.Element -Id 'sup.apps.bulk.Apply'
    $none = Find-ById -Element $editor.Element -Id 'sup.apps.bulk.None'
    foreach ($c in $search, $column, $value, $scope, $apply, $none) { & $assert ($c -and -not [string]::IsNullOrWhiteSpace([string]$c.ToolTip)) ('the bar control ' + [System.Windows.Automation.AutomationProperties]::GetAutomationId($c) + ' has a tooltip') }
    & $assert ([string]$column.SelectedItem -eq 'Supersedence') 'the bar starts on the column Supersedence'
    $search.Text = $scopeRow.Application
    $filtered = @($editor.View | ForEach-Object { $_ })
    & $assert ($filtered.Count -ge 1 -and $filtered.Count -lt 5 -and $filtered -contains $scopeRow) 'the search filters the rows'
    $column.SelectedItem = 'Version Rule'
    & $assert ($value.Items.Count -eq 3 -and -not $value.IsEditable) 'the value list follows the column'
    $value.SelectedItem = 'Exact'
    $scope.SelectedItem = 'Filtered'
    Invoke-OptionsButton $apply
    & $assert ($scopeRow['v_rule'] -eq 'Exact') 'Apply for Filtered writes the rows that pass the search'
    & $assert (@($rows | Where-Object { $_['v_rule'] -eq 'Exact' -and $_.Base -ne 'package-7zip' }).Count -eq $filtered.Count) 'only the rows that pass the search changed'
    $search.Text = ''
    $scope.SelectedItem = 'All'
    Invoke-OptionsButton $none
    & $assert (@($rows | Where-Object { $_['v_rule'] -ne 'Follow the default' }).Count -eq 0) 'None resets the column for every row'

    $column.SelectedItem = 'Install Scope'
    $value.SelectedItem = 'System'
    $scope.SelectedItem = 'All'
    Invoke-OptionsButton $apply
    & $assert (@($rows | Where-Object { $_['e_scope'] -and $_['v_scope'] -ne 'System' }).Count -eq 0 -and @($rows | Where-Object { -not $_['e_scope'] -and $_['v_scope'] -ne 'Packager default' }).Count -eq 0) 'All writes the cells that apply and skips the cells that do not'
    & $assert ($editor.Summary.Text -match 'skipped') 'the summary says how many cells were skipped'
    Invoke-OptionsButton (Find-ById -Element $editor.Element -Id 'sup.apps.select.all')
    & $assert (@($rows | Where-Object { $_.Checked }).Count -eq $rows.Count) 'Check all checks every row'
    Invoke-OptionsButton (Find-ById -Element $editor.Element -Id 'sup.apps.select.none')
    & $assert (@($rows | Where-Object { $_.Checked }).Count -eq 0) 'Check none clears every row'
    $scope.SelectedItem = 'Checked'
    Invoke-OptionsButton $apply
    & $assert ($editor.Summary.Text -match 'Check at least one') 'Checked with nothing checked says so'

    # The default of the page locks the titles of the applications that follow it.
    & $page.Editors['sup.mode'].Set 'Replace'
    & $assert (@($rows | Where-Object { $_['e_title'] -and $_.Base -ne $scopeRow.Base }).Count -eq 0 -and $scopeRow['e_title']) 'a default that is on closes the title of every application that follows it, and not the one with its own value Off'
    & $page.Editors['sup.mode'].Set 'Off'
    & $assert (@($rows | Where-Object { -not $_['e_title'] }).Count -eq 0) 'the default Off opens the titles again'

    # A Workbench profile column: a cell writes the profile of the application when the options are saved.
    Invoke-MenuItem (@($editor.Menu.Items | Where-Object { $_.Header -eq 'Reboot After Install' })[0]) $true
    $window.UpdateLayout()
    $reboot = Find-RowControl -Grid $grid -Row $scopeRow -Id 'appgrid.reboot'
    $reboot.SelectedItem = 'Soft reboot'
    & $assert ($scopeRow['v_reboot'] -eq 'Soft reboot') 'the reboot combo box writes the row'
    # Save and reload.
    $scopeRow = @($rows | Where-Object { $_.Base -eq $scopeRow.Base })[0]
    [void](Set-ApplicationGridCell -Setting (@($editor.Settings | Where-Object { $_.Id -eq 'oneclick' })[0]) -Work $editor.Work -Packager $scopeRow.Packager -Value 'Yes')
    [void](Set-ApplicationGridCell -Setting (@($editor.Settings | Where-Object { $_.Id -eq 'supersedence' })[0]) -Work $editor.Work -Packager $scopeRow.Packager -Value 'Replace and uninstall')
    & $page.Commit
    Save-Preferences $script:Prefs
    $script:Prefs = Read-Preferences
    $profileRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
    $stored = @(Get-Profiles -ApplicationId $scopeRow.Id -DataRoot $profileRoot | Where-Object { -not $_.IsDefault })
    & $assert ($stored.Count -eq 1 -and (Get-Profile -ApplicationId $scopeRow.Id -ProfileId $stored[0].ProfileId -DataRoot $profileRoot).Install.RebootPolicy -eq 'Soft reboot') 'the saved Workbench profile holds the reboot policy of the row'
    $entry = $script:Prefs.DeploymentConditions.Apps.($scopeRow.Base)
    & $assert ($entry.InstallMode -eq 'AllUsers') 'the saved file holds the scope of the row'
    & $assert ($script:Prefs.Supersedence.TitleOverrides.($scopeRow.Id).Mode -eq 'Replace and uninstall') 'the saved file holds the supersedence of the row'
    & $assert (@($script:Prefs.AppFlow.Tracked) -contains $scopeRow.Base -and [int]$script:Prefs.AppFlow.CadenceOverrides.($scopeRow.Base) -eq 45) 'the saved file holds the One Click keys of the row'
    $zipEntry = $script:Prefs.DeploymentConditions.Apps.'package-7zip'
    & $assert ($zipEntry.Split -eq 'Architecture' -and $zipEntry.Network -eq 'VpnOnly') 'the saved file keeps the conditions that were not changed'
    $again = New-OptionsTestPage -Id 'allapps'
    $back = @($again.Editors['sup.apps'].Rows | Where-Object { $_.Base -eq $scopeRow.Base })[0]
    & $assert ($back['v_scope'] -eq 'System' -and $back['v_supersedence'] -eq 'Replace and uninstall' -and $back['v_oneclick'] -eq 'Yes' -and $back['v_cadence'] -eq '45') 'a new page shows what was saved'
    'PASS: the grid of all applications in Options'
}
finally {
    if ($window) { $window.Close() }
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
