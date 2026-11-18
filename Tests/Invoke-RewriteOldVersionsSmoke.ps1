#Requires -Version 5.1
# Window probe of Rewrite Old Versions and Restore Old Versions: the dialog of the plan builds from the change plan, shows the notes,
# the clauses and the confirmation inside the window with enough contrast in the dark and in the light theme, keeps Confirm off until
# the box is checked, runs nothing before Confirm, shows the result rows after the run, offers Undo only for a rewrite, keeps the
# window open while a run works, passes a cancel to the run, and names a run that failed. The runs are stubs: the site is not read.
#   -ShotFolder   saves a maximized screenshot of each stage in each theme to that folder.
# SUITE_VISIBLE_WINDOWS=1 puts the window at 80,80 (the visible lab run).
param([string]$ShotFolder = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-rewrite-old-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir
. (Join-Path $PSScriptRoot 'WindowMeasure.ps1')
. $script:MeasureFunctions

$script:PassedChecks = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:PassedChecks += 1
    Write-Host ('  ok  ' + $Message)
}

function Sync-Layout { param($Window) $Window.UpdateLayout(); $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle) }
function Wait-Layout { param($Window, [int]$Passes = 4) for ($i = 0; $i -lt $Passes; $i++) { Sync-Layout $Window; Start-Sleep -Milliseconds 40 }; Sync-Layout $Window }
function Get-VisualDescendants {
    param([System.Windows.DependencyObject]$Root)
    $stack = New-Object System.Collections.Generic.Stack[System.Windows.DependencyObject]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $n = $stack.Pop()
        $n
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($n)
        for ($i = $count - 1; $i -ge 0; $i--) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($n, $i)) }
    }
}
function Test-VisibleInTree {
    param([System.Windows.UIElement]$Element)
    $d = $Element -as [System.Windows.DependencyObject]
    while ($d) {
        if ($d -is [System.Windows.UIElement] -and $d.Visibility -ne 'Visible') { return $false }
        $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
    }
    return $true
}
function Save-WindowPng {
    param([Parameter(Mandatory)]$Window, [Parameter(Mandatory)][string]$Path)
    $width = [int][Math]::Ceiling($Window.ActualWidth)
    $height = [int][Math]::Ceiling($Window.ActualHeight)
    $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($width, $height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($Window)
    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [System.IO.File]::Create($Path)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    return ('{0}x{1}' -f $width, $height)
}
function Open-Dialog {
    param($Dialog, [bool]$Dark, [bool]$Maximized)
    $w = $Dialog.Window
    $global:RewriteProbeDark = $Dark
    Initialize-SuiteTheme -Window $w -IsDarkGetter { [bool]$global:RewriteProbeDark }
    Set-DialogTheme -Dialog $w -IsDark $Dark
    $visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    $w.WindowStartupLocation = 'Manual'
    if ($Maximized) { $w.Left = 0; $w.Top = 0 }
    else { $w.Left = $(if ($visible) { 80 } else { -32000 }); $w.Top = $(if ($visible) { 80 } else { -32000 }) }
    $w.ShowInTaskbar = ($visible -or $Maximized)
    $w.Show()
    if ($Maximized) { $w.WindowState = 'Maximized' }
    Wait-Layout $w 6
    $rootGrid = Get-VisualDescendants $w | Where-Object { $_ -is [System.Windows.FrameworkElement] -and $_.Name -eq 'RootGrid' } | Select-Object -First 1
    for ($i = 0; $rootGrid -and $i -lt 100 -and $rootGrid.Opacity -lt 0.999; $i++) { Start-Sleep -Milliseconds 50; Sync-Layout $w }
}
function Invoke-ButtonClick {
    param($Dialog, $Button)
    $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    if ($Dialog.Window.IsVisible) { Wait-Layout $Dialog.Window 2 }
}

# A plan of three clauses with long names, a blocked application and two clauses that are left.
function New-ProbePlan {
    $long = 'Vendor Product With A Long Display Name 12.34.56.78 (x64) Enterprise Edition'
    $items = @(
        [pscustomobject]@{ Application = $long; Version = '12.34.56.78'; DeploymentType = ($long + ' - x64'); Clause = [pscustomobject]@{ Kind = 'File'; LogicalName = 'File_1'; Operator = 'GreaterEquals'; ExpectedValue = '12.34.56.78'; Path = 'C:\Program Files\Vendor'; FileName = 'product.exe'; Is64Bit = $true; Hive = ''; Key = ''; ValueName = '' }; Skip = $null; Rewrites = $true }
        [pscustomobject]@{ Application = 'Vendor Product 11.0'; Version = '11.0'; DeploymentType = 'Install'; Clause = [pscustomobject]@{ Kind = 'File'; LogicalName = 'File_2'; Operator = 'GreaterEquals'; ExpectedValue = '11.0'; Path = 'C:\Program Files\Vendor'; FileName = 'product.exe'; Is64Bit = $true; Hive = ''; Key = ''; ValueName = '' }; Skip = $null; Rewrites = $true }
        [pscustomobject]@{ Application = 'Vendor Product 10.0'; Version = '10.0'; DeploymentType = 'Install'; Clause = $null; Skip = 'the detection is not a single version clause'; Rewrites = $false }
        [pscustomobject]@{ Application = 'Vendor Product 9.0'; Version = '9.0'; DeploymentType = 'Install'; Clause = [pscustomobject]@{ Kind = 'File'; LogicalName = 'File_3'; Operator = 'IsEquals'; ExpectedValue = '9.0'; Path = 'C:\x'; FileName = 'a.exe'; Is64Bit = $false; Hive = ''; Key = ''; ValueName = '' }; Skip = 'the operator is IsEquals, not GreaterEquals'; Rewrites = $false }
    )
    return [pscustomobject]@{ Items = $items; Blocked = @("'Vendor Product 8.0' has a required deployment (All Users). Remove the required deployment, or leave this application out."); ClauseCount = 2; ApplicationCount = 2 }
}
function New-ProbeResult {
    param($Change, [int]$Failed = 0)
    $rows = @(); $n = 0
    foreach ($item in $Change.Items) { $n++; $rows += (New-SupersedenceChangeRow -Item $item -Succeeded ($n -gt $Failed) -Message $(if ($n -gt $Failed) { 'Written to the site and the log.' } else { 'access denied' })) }
    return (New-SupersedenceChangeResult -Rows $rows)
}

try {
    if ($ShotFolder) { [void](New-Item -ItemType Directory -Path $ShotFolder -Force) }
    $plan = New-ProbePlan
    $rewriteChange = New-SupersedenceRewriteChangePlan -Plan $plan -Restorable 2
    $restoreChange = New-SupersedenceRestoreChangePlan -Pending @(
        [pscustomobject]@{ At = [datetimeoffset]::UtcNow; Application = 'Vendor Product 11.0'; DeploymentType = 'Install'; OldOperator = 'GreaterEquals'; NewOperator = 'IsEquals'; LogicalName = 'File_2'; Restored = $false; Line = 'APD-Line:catalog:vendor|default|x64' })

    foreach ($theme in 'dark', 'light') {
        $dark = ($theme -eq 'dark')
        $maximized = [bool]$ShotFolder

        # ---- a rewrite: review, confirm, result, undo ----
        $label = "$theme / rewrite"
        $script:RunCalls = 0; $script:UndoCalls = 0; $script:SeenCancel = $null
        $dialog = New-RewriteOldVersionsDialog -Owner (New-OptionsTestOwner) -ChangePlan $rewriteChange `
            -Run { param($cancel) $script:RunCalls++; $script:SeenCancel = (& $cancel); New-ProbeResult -Change $rewriteChange -Failed 1 } `
            -Undo { param($cancel) $script:UndoCalls++; New-ProbeResult -Change $rewriteChange }
        $w = $dialog.Window
        Open-Dialog -Dialog $dialog -Dark $dark -Maximized $maximized
        try {
            & $assert ($w.Title -ceq 'Rewrite Old Versions') "[$label] the title is the action"
            & $assert ($dialog.Summary.Text -ceq '2 changes. Review the list, then confirm.') "[$label] the summary counts the changes: $($dialog.Summary.Text)"
            & $assert ($dialog.PlanGrid.Items.Count -eq 2 -and $dialog.PlanGrid.Visibility -eq 'Visible' -and $dialog.ResultGrid.Visibility -eq 'Collapsed') "[$label] the plan grid lists one row per clause"
            & $assert ($dialog.Notes.Visibility -eq 'Visible' -and $dialog.Notes.Text -match 'reads "not installed"' -and $dialog.Notes.Text -match 'Blocked: ' -and $dialog.Notes.Text -match 'Left: Vendor Product 10\.0' -and $dialog.Notes.Text -match '2 earlier rewrite') "[$label] the notes hold the effect, what is blocked, what is left and the pending restores"
            & $assert ($dialog.ConfirmButton.Content -ceq 'Rewrite 2 Clauses' -and -not $dialog.ConfirmButton.IsEnabled) "[$label] Confirm is named for the action and is off"
            & $assert ($dialog.ConfirmBox.Visibility -eq 'Visible' -and $dialog.ConfirmBox.IsChecked -ne $true) "[$label] the confirmation box starts unchecked"
            & $assert ($dialog.UndoButton.Visibility -eq 'Collapsed' -and $dialog.CloseButton.Visibility -eq 'Collapsed' -and $dialog.CancelButton.Visibility -eq 'Visible') "[$label] only Cancel and Confirm show before the run"
            Invoke-ButtonClick $dialog $dialog.ConfirmButton
            & $assert ($script:RunCalls -eq 0) "[$label] Confirm does nothing while it is off"

            # Everything is inside the window and readable.
            foreach ($control in @(Get-VisualDescendants $w | Where-Object { ($_ -is [System.Windows.Controls.TextBox] -or $_ -is [System.Windows.Controls.DataGrid] -or $_ -is [System.Windows.Controls.CheckBox]) -and $null -eq $_.TemplatedParent -and (Test-VisibleInTree $_) -and $_.ActualWidth -gt 0 })) {
                $box = $control.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $control.ActualWidth, $control.ActualHeight)))
                if ($box.Right -gt $w.ActualWidth + 0.5 -or $box.Bottom -gt $w.ActualHeight + 0.5) { throw ('FAIL: [{0}] {1} ends at {2},{3} in a window of {4}x{5}' -f $label, $control.GetType().Name, [int]$box.Right, [int]$box.Bottom, [int]$w.ActualWidth, [int]$w.ActualHeight) }
            }
            $confirmBox = $dialog.ConfirmButton.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $dialog.ConfirmButton.ActualWidth, $dialog.ConfirmButton.ActualHeight)))
            & $assert ($confirmBox.Bottom -le $w.ActualHeight + 0.5 -and $confirmBox.Right -le $w.ActualWidth + 0.5) "[$label] the buttons are inside the window"
            $weak = @(Get-TextContrast $w '' $w | Where-Object { $need = $(if ($_.Icon -or $_.Size -ge 24 -or ($_.Size -ge 18.66 -and $_.Bold)) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
            & $assert ($weak.Count -eq 0) ("[$label] every text has enough contrast " + (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
            $titleBar = Get-TitleBarContrast $w
            & $assert (($null -eq $titleBar.Active -or $titleBar.Active -ge 4.5) -and ($null -eq $titleBar.Inactive -or $titleBar.Inactive -ge 4.5)) "[$label] the title bar text has enough contrast"
            if ($ShotFolder) { $size = Save-WindowPng -Window $w -Path (Join-Path $ShotFolder ('rewrite-review-{0}.png' -f $theme)); Write-Host ('  shot  rewrite-review-{0}.png  {1}' -f $theme, $size) }

            $dialog.ConfirmBox.IsChecked = $true
            Wait-Layout $w 2
            & $assert ($dialog.ConfirmButton.IsEnabled) "[$label] Confirm turns on when the box is checked"
            $dialog.ConfirmBox.IsChecked = $false
            Wait-Layout $w 2
            & $assert (-not $dialog.ConfirmButton.IsEnabled) "[$label] Confirm turns off when the box is cleared"
            $dialog.ConfirmBox.IsChecked = $true
            Wait-Layout $w 2
            Invoke-ButtonClick $dialog $dialog.ConfirmButton
            & $assert ($script:RunCalls -eq 1 -and $script:SeenCancel -eq $false) "[$label] Confirm runs once and the cancel check starts false"
            & $assert ($dialog.State.Stage -eq 'Done' -and $dialog.ResultGrid.Visibility -eq 'Visible' -and $dialog.PlanGrid.Visibility -eq 'Collapsed') "[$label] the result grid replaces the plan"
            & $assert ($dialog.ResultGrid.Items.Count -eq 2 -and $dialog.Summary.Text -ceq '1 changed, 1 failed.') "[$label] the result rows and the summary follow the run: $($dialog.Summary.Text)"
            & $assert ($dialog.UndoButton.Visibility -eq 'Visible' -and $dialog.UndoButton.IsEnabled -and $dialog.CloseButton.Visibility -eq 'Visible' -and $dialog.CancelButton.Visibility -eq 'Collapsed' -and $dialog.ConfirmButton.Visibility -eq 'Collapsed') "[$label] Undo and Close show after the run"
            $weak = @(Get-TextContrast $w '' $w | Where-Object { $need = $(if ($_.Icon -or $_.Size -ge 24 -or ($_.Size -ge 18.66 -and $_.Bold)) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
            & $assert ($weak.Count -eq 0) ("[$label] every text of the result has enough contrast " + (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
            if ($ShotFolder) { $size = Save-WindowPng -Window $w -Path (Join-Path $ShotFolder ('rewrite-result-{0}.png' -f $theme)); Write-Host ('  shot  rewrite-result-{0}.png  {1}' -f $theme, $size) }

            Invoke-ButtonClick $dialog $dialog.UndoButton
            & $assert ($script:UndoCalls -eq 1 -and $dialog.State.Stage -eq 'Undone' -and $dialog.Summary.Text -ceq 'Undo: 2 changed, 0 failed.') "[$label] Undo runs the restore and names it: $($dialog.Summary.Text)"
            & $assert (-not $dialog.UndoButton.IsEnabled) "[$label] Undo cannot run twice"
            $outcome = & $dialog.Outcome
            & $assert ($outcome.Confirmed -and $outcome.Result.Summary -ceq '1 changed, 1 failed.' -and $outcome.UndoResult.Summary -ceq '2 changed, 0 failed.') "[$label] the outcome holds both results"
            Invoke-ButtonClick $dialog $dialog.CloseButton
            & $assert (-not $w.IsVisible) "[$label] Close closes the window"
        }
        finally { $w.Close() }

        # ---- Cancel before the confirmation changes nothing ----
        $label = "$theme / cancel"
        $script:RunCalls = 0
        $dialog = New-RewriteOldVersionsDialog -Owner (New-OptionsTestOwner) -ChangePlan $rewriteChange -Run { param($cancel) $script:RunCalls++; New-ProbeResult -Change $rewriteChange } -Undo { param($cancel) New-ProbeResult -Change $rewriteChange }
        Open-Dialog -Dialog $dialog -Dark $dark -Maximized $false
        try {
            Invoke-ButtonClick $dialog $dialog.CancelButton
            $outcome = & $dialog.Outcome
            & $assert (-not $dialog.Window.IsVisible -and -not $outcome.Confirmed -and $script:RunCalls -eq 0 -and $null -eq $outcome.Result) "[$label] Cancel closes the window and nothing ran"
        }
        finally { $dialog.Window.Close() }

        # ---- a restore: no Undo ----
        $label = "$theme / restore"
        $script:RunCalls = 0
        $dialog = New-RewriteOldVersionsDialog -Owner (New-OptionsTestOwner) -ChangePlan $restoreChange -Run { param($cancel) $script:RunCalls++; New-ProbeResult -Change $restoreChange }
        Open-Dialog -Dialog $dialog -Dark $dark -Maximized $maximized
        try {
            & $assert ($dialog.Window.Title -ceq 'Restore Old Versions' -and $dialog.Summary.Text -ceq '1 change. Review the list, then confirm.') "[$label] the title and the summary name the restore: $($dialog.Summary.Text)"
            & $assert ($dialog.Notes.Visibility -eq 'Collapsed' -and $dialog.ConfirmButton.Content -ceq 'Restore 1 Clause') "[$label] no notes, and Confirm is named for the action"
            & $assert ($dialog.ConfirmBox.Visibility -eq 'Visible' -and $dialog.ConfirmBox.IsChecked -ne $true) "[$label] the confirmation box starts unchecked"
            $weak = @(Get-TextContrast $dialog.Window '' $dialog.Window | Where-Object { $need = $(if ($_.Icon -or $_.Size -ge 24 -or ($_.Size -ge 18.66 -and $_.Bold)) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
            & $assert ($weak.Count -eq 0) ("[$label] every text has enough contrast " + (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
            if ($ShotFolder) { $size = Save-WindowPng -Window $dialog.Window -Path (Join-Path $ShotFolder ('restore-review-{0}.png' -f $theme)); Write-Host ('  shot  restore-review-{0}.png  {1}' -f $theme, $size) }
            $dialog.ConfirmBox.IsChecked = $true
            Wait-Layout $dialog.Window 2
            Invoke-ButtonClick $dialog $dialog.ConfirmButton
            & $assert ($script:RunCalls -eq 1 -and $dialog.UndoButton.Visibility -eq 'Collapsed' -and $dialog.CloseButton.Visibility -eq 'Visible') "[$label] the restore runs and offers no Undo"
        }
        finally { $dialog.Window.Close() }
    }

    # ---- a cancel while the run works: the window stays, the run sees the cancel ----
    $script:WindowHolder = $null
    $script:SawCancel = $null
    $script:StillOpen = $null
    $dialog = New-RewriteOldVersionsDialog -Owner (New-OptionsTestOwner) -ChangePlan $rewriteChange -Run {
        param($cancel)
        Invoke-ButtonClick $script:WindowHolder $script:WindowHolder.CancelButton
        $script:SawCancel = (& $cancel)
        $script:WindowHolder.Window.Close()
        $script:StillOpen = $script:WindowHolder.Window.IsVisible
        Get-SupersedenceStoppedResult -Items $rewriteChange.Items -Entries @() -Stop Cancelled
    } -Undo { param($cancel) New-ProbeResult -Change $rewriteChange }
    $script:WindowHolder = $dialog
    Open-Dialog -Dialog $dialog -Dark $true -Maximized $false
    try {
        $dialog.ConfirmBox.IsChecked = $true
        Wait-Layout $dialog.Window 2
        Invoke-ButtonClick $dialog $dialog.ConfirmButton
        & $assert ($script:SawCancel -eq $true) 'Cancel while the run works reaches the run'
        & $assert ($script:StillOpen -eq $true) 'the close box does not close the window while the run works'
        $outcome = & $dialog.Outcome
        & $assert ($outcome.Result.Stop -eq 'Cancelled' -and $outcome.Result.Summary -match '^The run was cancelled\.') 'the cancelled result is the outcome'
    }
    finally { $dialog.Window.Close() }

    # ---- a run that throws is named, and its rows are failed ----
    $dialog = New-RewriteOldVersionsDialog -Owner (New-OptionsTestOwner) -ChangePlan $rewriteChange -Run { param($cancel) throw 'the child could not start' } -Undo { param($cancel) New-ProbeResult -Change $rewriteChange }
    Open-Dialog -Dialog $dialog -Dark $false -Maximized $false
    try {
        $dialog.ConfirmBox.IsChecked = $true
        Wait-Layout $dialog.Window 2
        Invoke-ButtonClick $dialog $dialog.ConfirmButton
        & $assert ($dialog.Summary.Text -ceq 'The run failed. The log has the details.' -and $dialog.ResultGrid.Items.Count -eq 2 -and $dialog.UndoButton.Visibility -eq 'Collapsed') 'a run that throws is named and offers no Undo'
        & $assert (@($global:LogLines | Where-Object { $_ -like '*the child could not start*' }).Count -ge 1) 'the failure is written to the log'
    }
    finally { $dialog.Window.Close() }

    Write-Host ('PASS: Rewrite Old Versions probe ({0} checks)' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
