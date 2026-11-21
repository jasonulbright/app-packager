# Window proof of the grid of all applications in Options (Applications, All Applications). Runs in the interactive session of a lab account.
# Every action on the grid is real mouse input (SetCursorPos and mouse_event), not a UI Automation pattern: a click on a check box,
# a click on a combo box and on its item, a click on a bulk button. UI Automation only finds the elements and reads the result.
# The run then saves with OK, reads the preferences file, reopens Options in the light theme and captures both themes.
# It reads no site and publishes nothing. Images: options-applications-dark, options-applications-light (README folder screenshots),
# options-all-applications and light-options-all-applications (guide).
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
# The sidebar scrolls when the window is short; this height shows the Options button without a scroll.
Size-Win $win 1440 1040
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
$prefsPath = Join-Path $appDir 'AppPackager.preferences.json'
$cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
function WinList { (@($root.FindAll('Descendants', $cond)) | Where-Object { $_.Current.Name -notlike 'GlowWindow*' } | ForEach-Object { $_.Current.Name }) -join ' ; ' }
function Require([bool]$Condition, [string]$Message) { if (-not $Condition) { throw ('FAIL: ' + $Message) }; Log ('ok: ' + $Message) }
# A failure keeps a screenshot of the whole screen, so that a failed run shows what was on it.
function Guard([scriptblock]$Body) { try { & $Body } catch { Shot-Screen ('fail-' + (Get-Date -Format 'HHmmss')); throw } }
# The text that a combo box shows: its selected item (a combo box has no value pattern, and its name is the name of the row).
function Get-ComboText($Combo) {
    try { $sel = $Combo.GetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern).Current.GetSelection(); if ($sel.Count) { return [string]$sel[0].Current.Name } } catch { $null = $_ }
    $text = $Combo.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))
    if ($text) { return [string]$text.Current.Name }
    ''
}
function Toggle-State($El) { $El.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Current.ToggleState }

# The row of an application: the grid is virtualized, so the row is realized through the item container pattern first.
function Get-GridRow($o, [string]$Application) {
    $grid = Get-El $o 'sup.apps'
    $item = $grid.GetCurrentPattern([System.Windows.Automation.ItemContainerPattern]::Pattern).FindItemByProperty($null, $AE::NameProperty, $Application)
    if (-not $item) { throw "The grid has no row for $Application." }
    try { $item.GetCurrentPattern([System.Windows.Automation.VirtualizedItemPattern]::Pattern).Realize() } catch { $null = $_ }
    Start-Sleep -Milliseconds 700
    $rowCond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Application)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))
    $row = $grid.FindFirst('Descendants', $rowCond)
    if (-not $row) { throw "The row of $Application is not on screen." }
    $row
}
# A real click on a combo box, then a real click on the item in its list.
function Pick-Combo($Combo, [string]$Item) {
    Click-El $Combo; Start-Sleep -Milliseconds 700
    $itemCond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Item)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem)))
    $r = $Combo.Current.BoundingRectangle
    $hits = @($root.FindAll('Descendants', $itemCond) | Where-Object { $b = $_.Current.BoundingRectangle; $b.Width -gt 0 -and [Math]::Abs($b.X - $r.X) -lt 400 })
    if (-not $hits.Count) { throw "The list has no item $Item. Windows: $(WinList)" }
    Click-El $hits[0]; Start-Sleep -Milliseconds 700
}
function Open-AllApplications {
    # A real click on the Options button of the sidebar. The window comes to the front first, and a second click follows when the first one only activated it.
    Focus-Win $win
    $button = Get-El $win 'btnOptions'
    $rect = $button.Current.BoundingRectangle
    Log ('btnOptions: enabled {0}, rectangle {1},{2} {3}x{4}' -f $button.Current.IsEnabled, [int]$rect.X, [int]$rect.Y, [int]$rect.Width, [int]$rect.Height)
    $o = $null
    foreach ($try in 1..3) {
        Click-El $button; Start-Sleep -Seconds 2
        try { $o = Find-LastLike 'Options*' 6; break } catch { Log ('click ' + $try + ' did not open Options. Cursor ' + [System.Windows.Forms.Cursor]::Position + '. Windows: ' + (WinList)); Shot-Screen ('fail-click-' + $try) }
    }
    if (-not $o) { throw 'Options did not open after three real clicks.' }
    Start-Sleep -Seconds 1
    Size-Win $o 1400 1040
    $tree = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem))))
    $group = $tree | Where-Object { $_.Current.Name -eq 'Applications' } | Select-Object -First 1
    if (-not $group) { throw 'The Options tree has no Applications group.' }
    try { Expand-El $group } catch { $null = $_ }
    Start-Sleep -Milliseconds 600
    $node = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem)))) | Where-Object { $_.Current.Name -eq 'All Applications' } | Select-Object -First 1
    if (-not $node) { throw 'The Options tree has no All Applications page.' }
    Click-El $node; Start-Sleep -Milliseconds 1500
    $o
}

Section 'setup' {
    $s = Find-Last 'Setup' 20
    Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
    Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
}
Section 'the grid, dark theme' {
    Guard {
    $o = Open-AllApplications
    $grid = Get-El $o 'sup.apps'
    Require ($grid.Current.IsEnabled) 'the grid is on the page'

    # A real click on the check box of a row.
    $thonny = Get-GridRow $o 'Thonny'
    $box = Get-El $thonny 'appgrid.check'
    Click-El $box; Start-Sleep -Milliseconds 500
    Require ((Toggle-State $box) -eq 'On') 'a click on the check box of Thonny checks the row'

    # A real click on a combo box and on its item: Install scope User.
    $scope = Get-El $thonny 'appgrid.scope'
    Require ($scope.Current.IsEnabled) 'the install scope is open for Thonny (the packager offers both)'
    Pick-Combo $scope 'System'
    Require ((Get-ComboText $scope) -eq 'System') 'the combo box shows System after the click'

    # The bulk bar: the column and the value are combo boxes, the scope is a button. The checked row is Thonny.
    Pick-Combo (Get-El $o 'sup.apps.column') 'Install Scope'
    Pick-Combo (Get-El $o 'sup.apps.value') 'User'
    Click-El (Get-El $o 'sup.apps.bulk.Checked'); Start-Sleep -Milliseconds 800
    $thonny = Get-GridRow $o 'Thonny'
    Require ((Get-ComboText (Get-El $thonny 'appgrid.scope')) -eq 'User') 'Checked applications writes User on the checked row'
    Require ([bool](Get-El $o 'Install Scope: 1 application(s) changed. Choose OK to save.' -Quiet)) 'the summary line counts the application that the bulk action changed'

    # A closed combo box: 7-Zip offers one install scope.
    $plain = Get-GridRow $o '7-Zip'
    Require (-not (Get-El $plain 'appgrid.scope').Current.IsEnabled) 'the install scope is closed for 7-Zip (the packager offers one scope)'
    Require ((Get-El $plain 'appgrid.arch').Current.IsEnabled) 'the package source type is open for 7-Zip (the packager offers an architecture split)'

    # Package source type, supersedence and the title lock on Google Chrome.
    $chrome = Get-GridRow $o 'Google Chrome'
    Pick-Combo (Get-El $chrome 'appgrid.arch') 'x64 and ARM64'
    Require ((Get-ComboText (Get-El $chrome 'appgrid.arch')) -eq 'x64 and ARM64') 'the package source type of Google Chrome shows x64 and ARM64'
    Require ((Get-El $chrome 'appgrid.title').Current.IsEnabled) 'the title of Google Chrome is open while supersedence is off'
    Pick-Combo (Get-El $chrome 'appgrid.title') 'No version'
    Pick-Combo (Get-El $chrome 'appgrid.supersedence') 'Replace'
    $chrome = Get-GridRow $o 'Google Chrome'
    Require (-not (Get-El $chrome 'appgrid.title').Current.IsEnabled) 'the title of Google Chrome is locked while supersedence is on'
    Require ((Get-ComboText (Get-El $chrome 'appgrid.title')) -eq 'Include version') 'the locked title shows Include version'

    # Version rule on Obsidian.
    $obsidian = Get-GridRow $o 'Obsidian'
    Pick-Combo (Get-El $obsidian 'appgrid.rule') 'Exact'
    Require ((Get-ComboText (Get-El (Get-GridRow $o 'Obsidian') 'appgrid.rule')) -eq 'Exact') 'the version rule of Obsidian shows Exact'

    # A harmless bulk action: reset the title column for every application.
    Pick-Combo (Get-El $o 'sup.apps.column') 'Version in Title'
    Click-El (Get-El $o 'sup.apps.bulk.None'); Start-Sleep -Milliseconds 800
    Require ((Get-ComboText (Get-El (Get-GridRow $o 'Google Chrome') 'appgrid.title')) -eq 'Include version') 'No applications resets the stored title choice, and the lock still shows Include version'
    # Check none with the mouse.
    Click-El (Get-El $o 'sup.apps.select.none'); Start-Sleep -Milliseconds 500
    Require ((Toggle-State (Get-El (Get-GridRow $o 'Thonny') 'appgrid.check')) -eq 'Off') 'Check none clears the box of Thonny'
    $null = Get-GridRow $o 'Thonny'
    Shot $o 'options-applications-dark'
    Shot $o 'options-all-applications'

    # Save with a real click on OK.
    Click-El (Get-El $o 'btnOK'); Start-Sleep -Seconds 3
    }
}
Section 'the preferences file' {
    Require (Test-Path -LiteralPath $prefsPath) 'the preferences file exists'
    $prefs = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
    Log ('DeploymentConditions.Apps: ' + (($prefs.DeploymentConditions.Apps | ConvertTo-Json -Depth 5 -Compress)))
    Log ('Supersedence.TitleOverrides: ' + (($prefs.Supersedence.TitleOverrides | ConvertTo-Json -Depth 5 -Compress)))
    Require ($prefs.DeploymentConditions.Apps.'package-thonny'.InstallMode -eq 'CurrentUser') 'the file holds InstallMode CurrentUser for Thonny'
    Require ($prefs.DeploymentConditions.Apps.'package-chrome'.Split -eq 'Architecture') 'the file holds Split Architecture for Google Chrome'
    Require ($prefs.Supersedence.TitleOverrides.'catalog:package-chrome'.Mode -eq 'Replace') 'the file holds Mode Replace for Google Chrome'
    Require ([string]$prefs.DeploymentConditions.Apps.'package-chrome'.TitleMode -eq '') 'the file keeps no stored title choice for Google Chrome after the reset'
    Require ($prefs.Supersedence.TitleOverrides.'catalog:package-obsidian'.VersionRule -eq 'Exact') 'the file holds VersionRule Exact for Obsidian'
    Require ($prefs.Supersedence.Mode -eq 'Off') 'the default of supersedence is still Off'
}
Section 'the grid, light theme' {
    Guard {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2
    $o = Open-AllApplications
    $thonny = Get-GridRow $o 'Thonny'
    Require ((Get-ComboText (Get-El $thonny 'appgrid.scope')) -eq 'User') 'a new Options window shows the saved scope of Thonny'
    Require ((Get-ComboText (Get-El (Get-GridRow $o 'Google Chrome') 'appgrid.supersedence')) -eq 'Replace') 'a new Options window shows the saved supersedence of Google Chrome'
    $null = Get-GridRow $o 'Thonny'
    Shot $o 'options-applications-light'
    Shot $o 'light-options-all-applications'
    Click-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
    }
}
Stop-App 'start-apppackager.ps1'
