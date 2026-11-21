# Window proof of the grid of all applications in Options (Applications, All Applications). Runs in the interactive session of a lab account.
# Every action on the grid is real input (SetCursorPos, mouse_event and SendKeys), not a UI Automation pattern: a click on a check box, on a
# combo box and on its item, on a bulk button, on the column chooser and on its items, and typed text in the search box. UI Automation only
# finds the elements and reads the result. The run then saves with OK, reads the preferences file, reopens Options in the light theme and
# captures both themes. It reads no site and publishes nothing. Images: options-applications-dark, options-applications-light (README
# folder screenshots), options-all-applications, options-all-applications-columns and light-options-all-applications (guide).
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
    $o = $null
    foreach ($try in 1..3) {
        Click-El $button; Start-Sleep -Seconds 2
        try { $o = Find-LastLike 'Options*' 6; break } catch { Log ('click ' + $try + ' did not open Options. Windows: ' + (WinList)); Shot-Screen ('fail-click-' + $try) }
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
# An application that the catalog marks WsusSupport: No, so that the proof has a closed cell with a reason.
function Get-WsusRefusedApplication {
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $appDir 'Packagers') -Filter 'package-*.ps1')) {
        $lines = @(Get-Content -LiteralPath $f.FullName -TotalCount 40)
        if (@($lines | Where-Object { $_ -match '^\s*WsusSupport\s*:\s*No' }).Count) { $app = @($lines | Where-Object { $_ -match '^\s*App\s*:' })[0]; if ($app) { return ($app -replace '^\s*App\s*:\s*', '').Trim() } }
    }
}
# A real click on the Columns button, then on the item of one column.
function Toggle-Column($o, [string]$Header) {
    Click-El (Get-El $o 'sup.apps.columns'); Start-Sleep -Milliseconds 700
    $itemCond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Header)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::MenuItem)))
    $item = $root.FindFirst('Descendants', $itemCond)
    if (-not $item) { throw "The column menu has no item $Header." }
    Click-El $item; Start-Sleep -Milliseconds 600
    Send-Keys '{ESC}'; Start-Sleep -Milliseconds 400
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
        Require ((Get-El $o 'sup.apps').Current.IsEnabled) 'the grid is on the page'

        # A real click on the check box of a row, and on a combo box and its item.
        $thonny = Get-GridRow $o 'Thonny'
        $box = Get-El $thonny 'appgrid.check'
        Click-El $box; Start-Sleep -Milliseconds 500
        Require ((Toggle-State $box) -eq 'On') 'a click on the check box of Thonny checks the row'
        $scope = Get-El $thonny 'appgrid.scope'
        Require ($scope.Current.IsEnabled) 'the install scope is open for Thonny (the packager offers both)'
        Pick-Combo $scope 'System'
        Require ((Get-ComboText $scope) -eq 'System') 'the combo box shows System after the click'

        # The bulk bar: Set [column] to [value] for [scope] [Apply].
        Pick-Combo (Get-El $o 'sup.apps.column') 'Install Scope'
        Pick-Combo (Get-El $o 'sup.apps.value') 'User'
        Pick-Combo (Get-El $o 'sup.apps.scope') 'Checked'
        Click-El (Get-El $o 'sup.apps.bulk.Apply'); Start-Sleep -Milliseconds 800
        Require ((Get-ComboText (Get-El (Get-GridRow $o 'Thonny') 'appgrid.scope')) -eq 'User') 'Apply for Checked writes User on the checked row'
        Require ([bool](Get-El $o 'Install Scope: 1 application(s) changed; 316 skipped because the column does not apply to them. Choose OK to save.' -Quiet) -or [bool](Get-El $o 'Install Scope: 1 application(s) changed. Choose OK to save.' -Quiet)) 'the summary line counts the application that the bulk action changed'

        # A closed cell: 7-Zip has one install scope.
        $plain = Get-GridRow $o '7-Zip'
        Require (-not (Get-El $plain 'appgrid.scope').Current.IsEnabled) 'the install scope is closed for 7-Zip (the packager offers one scope)'
        Require ((Get-El $plain 'appgrid.source').Current.IsEnabled) 'the package source is open for 7-Zip (the packager offers an architecture split)'

        # Google Chrome: package source, then the title and the lock.
        $chrome = Get-GridRow $o 'Google Chrome'
        Pick-Combo (Get-El $chrome 'appgrid.source') 'x64 and ARM64'
        Require ((Get-ComboText (Get-El $chrome 'appgrid.source')) -eq 'x64 and ARM64') 'the package source of Google Chrome shows x64 and ARM64'
        Require ((Get-El $chrome 'appgrid.title').Current.IsEnabled) 'the title of Google Chrome is open while supersedence is off'
        Pick-Combo (Get-El $chrome 'appgrid.title') 'No version'
        Pick-Combo (Get-El $chrome 'appgrid.supersedence') 'Replace'
        $chrome = Get-GridRow $o 'Google Chrome'
        Require (-not (Get-El $chrome 'appgrid.title').Current.IsEnabled) 'the title of Google Chrome is closed while supersedence is on'
        Require ((Get-ComboText (Get-El $chrome 'appgrid.title')) -eq 'Include version') 'the closed title shows Include version'
        Pick-Combo (Get-El (Get-GridRow $o 'Obsidian') 'appgrid.rule') 'Exact'
        Require ((Get-ComboText (Get-El (Get-GridRow $o 'Obsidian') 'appgrid.rule')) -eq 'Exact') 'the version rule of Obsidian shows Exact'
        Click-El (Get-El $o 'sup.apps.select.none'); Start-Sleep -Milliseconds 500
        Require ((Toggle-State (Get-El (Get-GridRow $o 'Thonny') 'appgrid.check')) -eq 'Off') 'Check none clears the box of Thonny'
        $null = Get-GridRow $o 'Thonny'
        Shot $o 'options-applications-dark'
        Shot $o 'options-all-applications'

        # The column chooser, a typed search with the scope Filtered, and a closed cell with a reason.
        foreach ($h in 'In One Click', 'Publish to WSUS') { Toggle-Column $o $h }
        # Three of the first five columns are hidden again, so that the new columns fit on the screen.
        foreach ($h in 'Install Scope', 'Version in Title', 'Version Rule') { Toggle-Column $o $h }
        $headers = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::HeaderItem)))) | ForEach-Object { $_.Current.Name }
        Log ('headers: ' + ($headers -join ' | '))
        Require (@($headers | Where-Object { $_ -ieq 'In One Click' }).Count -gt 0 -and @($headers | Where-Object { $_ -ieq 'Publish to WSUS' }).Count -gt 0) 'the chooser shows the columns In One Click and Publish to WSUS'
        Click-El (Get-El $o 'sup.apps.search'); Send-Keys 'Google Chrome'; Start-Sleep -Seconds 1
        Pick-Combo (Get-El $o 'sup.apps.column') 'In One Click'
        Pick-Combo (Get-El $o 'sup.apps.value') 'Yes'
        Pick-Combo (Get-El $o 'sup.apps.scope') 'Filtered'
        Click-El (Get-El $o 'sup.apps.bulk.Apply'); Start-Sleep -Milliseconds 800
        Require ((Get-ComboText (Get-El (Get-GridRow $o 'Google Chrome') 'appgrid.oneclick')) -eq 'Yes') 'Apply for Filtered tracks the application that the search shows'
        Click-El (Get-El $o 'sup.apps.search'); Send-Keys '^a'; Send-Keys '{DEL}'; Start-Sleep -Seconds 1
        $refused = Get-WsusRefusedApplication
        Log ('application that WSUS cannot carry: ' + $refused)
        Require ([bool]$refused) 'the catalog has an application that WSUS cannot carry'
        $wsusCell = Get-El (Get-GridRow $o $refused) 'appgrid.towsus'
        Require (-not $wsusCell.Current.IsEnabled) ('the Publish to WSUS cell of ' + $refused + ' is closed')
        Move-To $wsusCell; Start-Sleep -Milliseconds 1800
        Shot $o 'options-all-applications-columns' -Hover

        # Save with a real click on OK.
        Click-El (Get-El $o 'btnOK'); Start-Sleep -Seconds 3
    }
}
Section 'the preferences file' {
    Require (Test-Path -LiteralPath $prefsPath) 'the preferences file exists'
    $prefs = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
    Log ('DeploymentConditions.Apps: ' + (($prefs.DeploymentConditions.Apps | ConvertTo-Json -Depth 5 -Compress)))
    Log ('Supersedence.TitleOverrides: ' + (($prefs.Supersedence.TitleOverrides | ConvertTo-Json -Depth 5 -Compress)))
    Log ('AppFlow.Tracked: ' + (@($prefs.AppFlow.Tracked) -join ','))
    Require ($prefs.DeploymentConditions.Apps.'package-thonny'.InstallMode -eq 'CurrentUser') 'the file holds InstallMode CurrentUser for Thonny'
    Require ($prefs.DeploymentConditions.Apps.'package-chrome'.Split -eq 'Architecture') 'the file holds Split Architecture for Google Chrome'
    Require ([string]$prefs.DeploymentConditions.Apps.'package-chrome'.TitleMode -eq 'NoVersion') 'the file keeps the stored title choice of Google Chrome while the title is closed'
    Require ($prefs.Supersedence.TitleOverrides.'catalog:package-chrome'.Mode -eq 'Replace') 'the file holds Mode Replace for Google Chrome'
    Require ($prefs.Supersedence.TitleOverrides.'catalog:package-obsidian'.VersionRule -eq 'Exact') 'the file holds VersionRule Exact for Obsidian'
    Require (@($prefs.AppFlow.Tracked) -contains 'package-chrome') 'the file tracks Google Chrome in One Click'
    Require ($prefs.Supersedence.Mode -eq 'Off') 'the default of supersedence is still Off'
}
Section 'the grid, light theme' {
    Guard {
        Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2
        $o = Open-AllApplications
        Require ((Get-ComboText (Get-El (Get-GridRow $o 'Thonny') 'appgrid.scope')) -eq 'User') 'a new Options window shows the saved scope of Thonny'
        Require ((Get-ComboText (Get-El (Get-GridRow $o 'Google Chrome') 'appgrid.supersedence')) -eq 'Replace') 'a new Options window shows the saved supersedence of Google Chrome'
        $null = Get-GridRow $o 'Thonny'
        Shot $o 'options-applications-light'
        Shot $o 'light-options-all-applications'
        Click-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
        Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
    }
}
Stop-App 'start-apppackager.ps1'
