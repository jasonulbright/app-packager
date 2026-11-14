#Requires -Version 5.1
# Window probe of the Options window: the text and the control borders of every section have enough contrast
# in the light and in the dark theme, for enabled and for disabled controls.
#   Text:     4.5:1 (3:1 for large text: 24 px, or 18.7 px and bold, and for an icon glyph such as the clear button of a box).
#   Border:   3:1 between the border of a control and the backdrop around it.
# Each defect has a key. A defect that Tests\KnownContrastDefects.ps1 lists is a known limit of the shared theme and
# does not fail the probe. A new defect fails it, and a listed defect that no longer happens fails it too.
# Limits: the border check reads the BorderBrush of the control, not the parts of its template. Hover, focus and pressed states are not measured.
# Run with -Report to print every defect and exit with 0.
param([switch]$Report)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-options-contrast-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir
. (Join-Path $PSScriptRoot 'WindowMeasure.ps1')
. $script:MeasureFunctions
. (Join-Path $PSScriptRoot 'KnownContrastDefects.ps1')

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

# One label for a control: its name, or its type and its text.
function Get-ControlLabel {
    param($Control)
    if ($Control.Name) { return [string]$Control.Name }
    $text = ''
    if ($Control -is [System.Windows.Controls.TextBox]) { $text = [string]$Control.Text }
    elseif ($Control -is [System.Windows.Controls.ComboBox]) { $text = [string]$Control.Text }
    elseif ($Control -is [System.Windows.Controls.ContentControl] -and $Control.Content -is [string]) { $text = [string]$Control.Content }
    if ($text.Length -gt 24) { $text = $text.Substring(0, 24) }
    return ('{0}[{1}]' -f $Control.GetType().Name, $text)
}

# The buttons, boxes and check boxes of the window, enabled and then disabled. A control inside the template of another
# control (the clear button of a box, the arrow of a combo box) belongs to that control.
function Get-OptionsControlContrast {
    # -Root limits the scan to one part of the window (the page of a section); the window scan runs once.
    param($Window, $Root = $Window)
    $rows = @()
    $controls = @(Get-VisualDescendants $Root | Where-Object {
            ($_ -is [System.Windows.Controls.Primitives.ButtonBase] -or $_ -is [System.Windows.Controls.TextBox] -or $_ -is [System.Windows.Controls.ComboBox]) -and
            $_.Name -notlike 'PART_*' -and $null -eq $_.TemplatedParent -and (Test-VisibleInTree $_) -and $_.ActualWidth -gt 0 })
    # Enabled first. Then the whole root is disabled at once, so that one layout pass serves every control.
    $enabled = @{}
    foreach ($c in $controls) { $enabled[$c] = Get-ControlColors $c $Window }
    $local = $Root.ReadLocalValue([System.Windows.UIElement]::IsEnabledProperty)
    try {
        $Root.IsEnabled = $false; Wait-Layout $Window 2
        foreach ($c in $controls) {
            $en = $enabled[$c]
            $dis = Get-ControlColors $c $Window
            $rows += [pscustomobject]@{ Name = (Get-ControlLabel $c); Type = $en.Type; EnabledText = $en.TextRatio; DisabledText = $dis.TextRatio; EnabledBorder = $en.BorderRatio; DisabledBorder = $dis.BorderRatio }
        }
    }
    finally {
        if ($local -eq [System.Windows.DependencyProperty]::UnsetValue) { $Root.ClearValue([System.Windows.UIElement]::IsEnabledProperty) } else { $Root.IsEnabled = [bool]$local }
        Sync-Layout $Window
    }
    return $rows
}

$known = @{ dark = @(Get-KnownContrastDefect -Theme 'dark'); light = @(Get-KnownContrastDefect -Theme 'light') }
$found = @{ dark = @{}; light = @{} }
$sections = 0
try {
    Set-Content -LiteralPath $global:FixturePrefsPath -Value '{ "FirstRunCompleted": true, "Systems": { "ConfigMgr": true, "Intune": true, "Wsus": true } }' -Encoding UTF8
    foreach ($theme in 'dark', 'light') {
        $themeName = $(if ($theme -eq 'dark') { 'Dark.Steel' } else { 'Light.Blue' })
        $script:Prefs = Read-Preferences
        $owner = New-OptionsTestOwner
        $dialog = New-OptionsDialog -Owner $owner
        $w = $dialog.Window
        # The same call that the shell makes for a dialog, so that the SuiteCommon title bar brushes are measured.
        $global:ContrastProbeDark = ($theme -eq 'dark')
        Initialize-SuiteTheme -Window $w -IsDarkGetter { [bool]$global:ContrastProbeDark }
        Set-DialogTheme -Dialog $w -IsDark ($theme -eq 'dark')
        $w.WindowStartupLocation = 'Manual'; $w.Left = -32000; $w.Top = -32000; $w.ShowInTaskbar = $false
        $w.Show()
        Wait-Layout $w 6
        $rootGrid = Get-VisualDescendants $w | Where-Object { $_ -is [System.Windows.FrameworkElement] -and $_.Name -eq 'RootGrid' } | Select-Object -First 1
        for ($i = 0; $rootGrid -and $i -lt 100 -and $rootGrid.Opacity -lt 0.999; $i++) { Start-Sleep -Milliseconds 50; Sync-Layout $w }
        try {
            $firstView = $true
            foreach ($section in $dialog.Definition) {
                if ($section.Kind -eq 'Group') { continue }
                & $dialog.Navigate ([string]$section.Id) ''
                Wait-Layout $w 2
                $sections++
                $scan = $(if ($firstView) { $w } else { $dialog.Context.Sections[[string]$section.Id].Element })
                foreach ($t in @(Get-TextContrast $w "" $scan)) {
                    $large = $t.Icon -or ($t.Size -ge 24) -or ($t.Size -ge 18.66 -and $t.Bold)
                    $need = $(if ($large) { 3.0 } else { 4.5 })
                    if ($t.Ratio -lt $need) { $found[$theme][('text:' + $t.Text)] = ('[{0}] {1} on {2} = {3}' -f $section.Id, $t.Foreground, $t.Backdrop, $t.Ratio) }
                }
                foreach ($c in @(Get-OptionsControlContrast $w $scan)) {
                    foreach ($m in 'EnabledText', 'DisabledText') { if ($null -ne $c.$m -and $c.$m -lt 4.5) { $found[$theme][('{0}:{1}' -f $c.Name, $m)] = ('[{0}] {1}' -f $section.Id, $c.$m) } }
                    foreach ($m in 'EnabledBorder', 'DisabledBorder') { if ($null -ne $c.$m -and $c.$m -lt 3.0) { $found[$theme][('{0}:{1}' -f $c.Name, $m)] = ('[{0}] {1}' -f $section.Id, $c.$m) } }
                }
                $firstView = $false
            }
            $tb = Get-TitleBarContrast $w
            if ($null -ne $tb.Active -and $tb.Active -lt 4.5) { $found[$theme]['titlebar:active'] = [string]$tb.Active }
            if ($null -ne $tb.Inactive -and $tb.Inactive -lt 4.5) { $found[$theme]['titlebar:inactive'] = [string]$tb.Inactive }
        }
        finally { $w.Close() }
    }

    Write-Host ('Measured {0} section views in two themes.' -f $sections)
    $failed = $false
    foreach ($theme in 'dark', 'light') {
        $keys = @($found[$theme].Keys | Sort-Object)
        Write-Host ('{0}: {1} defects' -f $theme, $keys.Count)
        if ($Report) { foreach ($k in $keys) { Write-Host ('  {0}   {1}' -f $k, $found[$theme][$k]) } }
        else {
            foreach ($k in $keys) { if ($known[$theme] -notcontains $k) { Write-Host ('  NEW   {0}   {1}' -f $k, $found[$theme][$k]); $failed = $true } }
            foreach ($k in $known[$theme]) { if (-not $found[$theme].ContainsKey($k)) { Write-Host ('  FIXED {0}: remove it from Tests\KnownContrastDefects.ps1' -f $k); $failed = $true } }
        }
    }
    if ($failed) { Write-Host 'FAIL: Options contrast probe'; exit 1 }
    Write-Host 'PASS: Options contrast probe'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
