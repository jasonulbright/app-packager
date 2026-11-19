#Requires -Version 5.1
# Window probe of the Options window: no section shows a horizontal scroll bar, in the light and in the dark theme, at the
# smallest size of the window, at its own size and at two larger sizes. The navigation pane (the tree and the search list)
# fits its content too, and a splitter lets the user change its width.
# Run with -Report to print every size of every scroll viewer that scrolls sideways and exit with 0.
# SUITE_VISIBLE_WINDOWS=1 puts the window at 40,40 (the visible lab run).
param([switch]$Report)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-options-scrollbar-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir

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

# The scroll viewers of the window that show a horizontal bar. The owner is the control that the viewer belongs to.
function Get-HorizontalBar {
    param($Window)
    foreach ($sv in @(Get-VisualDescendants $Window | Where-Object { $_ -is [System.Windows.Controls.ScrollViewer] -and (Test-VisibleInTree $_) })) {
        if ($sv.ComputedHorizontalScrollBarVisibility -ne [System.Windows.Visibility]::Visible) { continue }
        $owner = $sv.TemplatedParent
        if (-not $owner) { $owner = $sv }
        $label = [string]$owner.Name
        if (-not $label) { $label = $owner.GetType().Name }
        [pscustomobject]@{
            Owner      = $label
            Extent     = [Math]::Round($sv.ExtentWidth, 1)
            Viewport   = [Math]::Round($sv.ViewportWidth, 1)
            Over       = [Math]::Round($sv.ExtentWidth - $sv.ViewportWidth, 1)
        }
    }
}

$sizes = @(
    @{ Name = 'smallest'; Width = 0; Height = 0 },
    @{ Name = 'own size'; Width = 1260; Height = 760 },
    @{ Name = '1440x900'; Width = 1440; Height = 900 },
    @{ Name = '1920x1032'; Width = 1920; Height = 1032 }
)
$defects = New-Object System.Collections.Generic.List[string]
$views = 0
try {
    Set-Content -LiteralPath $global:FixturePrefsPath -Value '{ "FirstRunCompleted": true, "Systems": { "ConfigMgr": true, "Intune": true, "Wsus": true } }' -Encoding UTF8
    foreach ($theme in 'dark', 'light') {
        $script:Prefs = Read-Preferences
        $owner = New-OptionsTestOwner
        $dialog = New-OptionsDialog -Owner $owner
        $w = $dialog.Window
        # The same calls that the shell makes for a dialog.
        $global:ScrollBarProbeDark = ($theme -eq 'dark')
        Initialize-SuiteTheme -Window $w -IsDarkGetter { [bool]$global:ScrollBarProbeDark }
        Set-DialogTheme -Dialog $w -IsDark ($theme -eq 'dark')
        $where = $(if ($env:SUITE_VISIBLE_WINDOWS -eq '1') { 40 } else { -32000 })
        $w.WindowStartupLocation = 'Manual'; $w.Left = $where; $w.Top = $where; $w.ShowInTaskbar = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
        $w.Show()
        Wait-Layout $w 6
        $splitter = $w.FindName('splitNav')
        try {
            foreach ($size in $sizes) {
                $w.Width = $(if ($size.Width -gt 0) { $size.Width } else { $w.MinWidth })
                $w.Height = $(if ($size.Height -gt 0) { $size.Height } else { $w.MinHeight })
                Wait-Layout $w 3
                $nav = $w.FindName('treeNav')
                $tag = '{0} / {1} ({2}x{3})' -f $theme, $size.Name, [int]$w.ActualWidth, [int]$w.ActualHeight
                foreach ($section in $dialog.Definition) {
                    if ($section.Kind -eq 'Group') { continue }
                    & $dialog.Navigate ([string]$section.Id) ''
                    Wait-Layout $w 2
                    $views++
                    foreach ($bar in @(Get-HorizontalBar $w)) {
                        $line = '{0} / {1}: {2} scrolls sideways, content {3} in {4} ({5} too wide)' -f $tag, $section.Id, $bar.Owner, $bar.Extent, $bar.Viewport, $bar.Over
                        $defects.Add($line)
                    }
                }
            }
        }
        finally { $w.Close() }
    }

    Write-Host ('Measured {0} section views in two themes at {1} sizes.' -f $views, $sizes.Count)
    if (-not $splitter) { $defects.Add('The window has no splitter between the list of sections and the section.') }
    elseif ($splitter.Width -lt 4) { $defects.Add(('The splitter is {0} wide; it is hard to grab.' -f $splitter.Width)) }
    if ($Report) { foreach ($d in $defects) { Write-Host ('  ' + $d) }; Write-Host ('{0} horizontal bars.' -f $defects.Count); exit 0 }
    foreach ($d in $defects) { Write-Host ('  FAIL: ' + $d) }
    if ($defects.Count -gt 0) { Write-Host 'FAIL: Options scroll bar probe'; exit 1 }
    Write-Host 'PASS: Options scroll bar probe'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
