#Requires -Version 5.1
# Window probe of the review before staging: the dialog of each of the seven sections that have install options builds
# from the Options definitions, starts from the saved values, shows its fields inside the window and with enough contrast
# in the dark and in the light theme, returns nothing when nothing changed, returns the changed values for this run when a
# field changed, and never writes the preferences or the packager settings files. Cancel returns no run values.
#   -ShotFolder   saves a maximized screenshot of each dialog in each theme to that folder.
# SUITE_VISIBLE_WINDOWS=1 puts the window at 80,80 (the visible lab run).
param([string]$ShotFolder = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-stage-review-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir
. (Join-Path $PSScriptRoot 'WindowMeasure.ps1')
. $script:MeasureFunctions

# The probe shows the real error of a failed Continue instead of the short text for the window.
function global:Get-UserErrorText { param($ErrorRecord, [string]$Fallback) return ([string]$ErrorRecord.Exception.Message + ' at ' + [string]$ErrorRecord.InvocationInfo.PositionMessage) }
$global:Messages = @()
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

# One row per section: the application, its packager, a field to change and the run values that change must produce.
$cases = @(
    @{ Id = 'ssms'; App = 'SQL Server Management Studio'; Field = 'ssms.ui'; Value = 'Passive'; Check = { param($run) $run.Preferences.SSMSInstallOptions.UIMode -eq 'Passive' } }
    @{ Id = 'dbeaver'; App = 'DBeaver Community'; Field = 'dbv.scope'; Value = 'User'; Check = { param($run) $run.Preferences.DBeaverInstallOptions.InstallScope -eq 'User' } }
    @{ Id = 'adobe'; App = 'Adobe Acrobat Reader'; Field = 'adobe.eula'; Value = $true; Check = { param($run) $run.Preferences.AdobeReaderInstallOptions.AcceptEula -eq $true } }
    @{ Id = 'bc5'; App = 'Beyond Compare 5'; Field = 'bc.remote'; Value = $true; Check = { param($run) $run.Preferences.BeyondCompareOptions.DisableRemoteProfiles -eq $true } }
    @{ Id = 'm365'; App = 'Microsoft 365 Apps'; Field = 'm365.channel'; Value = 'Current'; Check = { param($run) $run.Preferences.M365Channel -eq 'Current' } }
    @{ Id = 'cwa'; App = 'Citrix Workspace App'; Field = 'cwa.storeUrl'; Value = 'https://store.example.com'; Check = { param($run) $run.Files.'citrix-workspace-switches.json'.Store.Url -eq 'https://store.example.com' } }
    @{ Id = 'teamviewer'; App = 'TeamViewer Host'; Field = 'tv.config'; Value = 'abc123'; Check = { param($run) $run.Files.'teamviewer-host-config.json'.CustomConfigId -eq 'abc123' } }
)

function Open-ReviewWindow {
    param($Dialog, [bool]$Dark, [bool]$Maximized)
    $w = $Dialog.Window
    $global:ReviewProbeDark = $Dark
    Initialize-SuiteTheme -Window $w -IsDarkGetter { [bool]$global:ReviewProbeDark }
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

try {
    Set-Content -LiteralPath $global:FixturePrefsPath -Value '{ "FirstRunCompleted": true }' -Encoding UTF8
    $global:SavedCount = 0
    if ($ShotFolder) { [void](New-Item -ItemType Directory -Path $ShotFolder -Force) }

    foreach ($theme in 'dark', 'light') {
        $dark = ($theme -eq 'dark')
        foreach ($case in $cases) {
            $label = '{0} / {1}' -f $theme, $case.Id
            $script:Prefs = Read-Preferences
            $before = ConvertTo-Json -InputObject $script:Prefs -Depth 8
            $owner = New-OptionsTestOwner
            $dialog = New-StageReviewDialog -Owner $owner -Application $case.App -SectionId $case.Id -Prefs $script:Prefs
            $w = $dialog.Window
            Open-ReviewWindow -Dialog $dialog -Dark $dark -Maximized ([bool]$ShotFolder)
            try {
                & $assert ($w.Title -ceq ('Configure: ' + $case.App)) ('[{0}] the title names the application' -f $label)
                & $assert ($dialog.Page.Editors.Contains($case.Field)) ('[{0}] the dialog has the field {1} of the Options section' -f $label, $case.Field)
                $heading = @(Get-VisualDescendants $w | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and [string]$_.Tag -eq 'section-title' })
                & $assert (@($heading | Where-Object { Test-VisibleInTree $_ }).Count -eq 0) ('[{0}] the section heading is not repeated under the title' -f $label)
                & $assert ($dialog.ContinueButton.IsDefault -and $dialog.CancelButton.IsCancel -and $dialog.ContinueButton.Content -ceq 'Continue' -and $dialog.CancelButton.Content -ceq 'Cancel') ('[{0}] Continue is the default button and Cancel the cancel button' -f $label)

                # The fields sit inside the window, and the buttons stay in view.
                foreach ($control in @(Get-VisualDescendants $w | Where-Object { ($_ -is [System.Windows.Controls.TextBox] -or $_ -is [System.Windows.Controls.ComboBox]) -and $null -eq $_.TemplatedParent -and (Test-VisibleInTree $_) -and $_.ActualWidth -gt 0 })) {
                    $box = $control.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $control.ActualWidth, $control.ActualHeight)))
                    if ($box.Right -gt $w.ActualWidth + 0.5) { throw ('FAIL: [{0}] an input ends at {1} in a window of {2}' -f $label, [int]$box.Right, [int]$w.ActualWidth) }
                }
                $continueBox = $dialog.ContinueButton.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $dialog.ContinueButton.ActualWidth, $dialog.ContinueButton.ActualHeight)))
                & $assert ($continueBox.Bottom -le $w.ActualHeight + 0.5 -and $continueBox.Right -le $w.ActualWidth + 0.5) ('[{0}] the buttons are inside the window' -f $label)

                # The text has enough contrast in the theme, and so has the title bar.
                $weak = @(Get-TextContrast $w '' $w | Where-Object { $need = $(if ($_.Icon -or $_.Size -ge 24 -or ($_.Size -ge 18.66 -and $_.Bold)) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
                & $assert ($weak.Count -eq 0) ('[{0}] every text has enough contrast {1}' -f $label, (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
                $titleBar = Get-TitleBarContrast $w
                & $assert (($null -eq $titleBar.Active -or $titleBar.Active -ge 4.5) -and ($null -eq $titleBar.Inactive -or $titleBar.Inactive -ge 4.5)) ('[{0}] the title bar text has enough contrast' -f $label)

                if ($ShotFolder) {
                    $size = Save-WindowPng -Window $w -Path (Join-Path $ShotFolder ('review-{0}-{1}.png' -f $case.Id, $theme))
                    Write-Host ('  shot  review-{0}-{1}.png  {2}' -f $case.Id, $theme, $size)
                }

                # Nothing changed: the run takes the saved values.
                & $assert ((& $dialog.Apply) -eq $true -and $dialog.Result.Continue -eq $true -and $dialog.Result.RunOptionsJson -ceq '') ('[{0}] Continue with no change returns no run values {1} {2}' -f $label, ($global:Messages -join ' | '), ([string]$dialog.Result.RunOptionsJson).Replace(',', "`n"))

                # A changed field: the run values hold it, and nothing is saved.
                & $dialog.Page.Editors[$case.Field].Set $case.Value
                & $assert ((& $dialog.Apply) -eq $true -and [bool]$dialog.Result.RunOptionsJson) ('[{0}] Continue with a changed field returns run values' -f $label)
                $run = $dialog.Result.RunOptionsJson | ConvertFrom-Json
                & $assert ([bool](& $case.Check $run)) ('[{0}] the run values hold the changed field' -f $label)
                & $assert ((ConvertTo-Json -InputObject $script:Prefs -Depth 8) -ceq $before) ('[{0}] the saved preferences are the same object values as before' -f $label)
                & $assert ($global:SavedCount -eq 0 -and -not (Test-Path -LiteralPath (Get-CwaSwitchesPath)) -and -not (Test-Path -LiteralPath (Get-TvHostConfigPath)) -and -not (Test-Path -LiteralPath (Get-M365CustomConfigPath))) ('[{0}] no preference file and no settings file was written' -f $label)
            }
            finally { $w.Close() }
        }
    }

    # The buttons: Continue closes the window with the run values, Cancel closes it with none.
    $script:Prefs = Read-Preferences
    $dialog = New-StageReviewDialog -Owner (New-OptionsTestOwner) -Application 'DBeaver Community' -SectionId 'dbeaver' -Prefs $script:Prefs
    Open-ReviewWindow -Dialog $dialog -Dark $true -Maximized $false
    & $dialog.Page.Editors['dbv.scope'].Set 'User'
    $dialog.ContinueButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    Wait-Layout $dialog.Window 2
    & $assert ($dialog.Result.Continue -eq $true -and $dialog.Result.RunOptionsJson -match 'InstallScope' -and -not $dialog.Window.IsVisible) 'the Continue button closes the window with the run values'

    $dialog = New-StageReviewDialog -Owner (New-OptionsTestOwner) -Application 'DBeaver Community' -SectionId 'dbeaver' -Prefs $script:Prefs
    Open-ReviewWindow -Dialog $dialog -Dark $true -Maximized $false
    & $dialog.Page.Editors['dbv.scope'].Set 'User'
    $dialog.CancelButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    Wait-Layout $dialog.Window 2
    & $assert ($dialog.Result.Continue -eq $false -and $dialog.Result.RunOptionsJson -ceq '' -and -not $dialog.Window.IsVisible) 'the Cancel button closes the window with no run values'

    # A value that cannot be stored stops Continue and names the field.
    $global:Messages = @()
    $dialog = New-StageReviewDialog -Owner (New-OptionsTestOwner) -Application 'Microsoft 365 Apps' -SectionId 'm365' -Prefs $script:Prefs
    Open-ReviewWindow -Dialog $dialog -Dark $false -Maximized $false
    & $dialog.Page.Editors['m365.languages'].Set 'not a language'
    & $assert ((& $dialog.Apply) -eq $false -and $dialog.Result.Continue -eq $false -and ($global:Messages -join ' ') -match 'culture codes') 'a language that is not a culture code stops Continue'
    $dialog.Window.Close()

    # The Microsoft 365 review has no import rows, because importing the file would store it.
    $dialog = New-StageReviewDialog -Owner (New-OptionsTestOwner) -Application 'Microsoft 365 Apps' -SectionId 'm365' -Prefs $script:Prefs
    & $assert (-not $dialog.Page.Editors.Contains('m365.file') -and -not $dialog.Page.Editors.Contains('m365.useCustom') -and $dialog.Page.Editors.Contains('m365.exclude')) 'the Microsoft 365 review has the generated settings and no imported file rows'

    Write-Host ('PASS: Stage review probe ({0} checks)' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
