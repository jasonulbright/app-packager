#Requires -Version 5.1
# Window probe: every window and dialog that the repository builds from XAML (MainWindow.xaml, WorkbenchWindow.xaml
# and each MetroWindow string of start-apppackager.ps1) is shown at its native size. The probe fails any control
# that is cut off at the right edge of its window or of a scroll area that does not scroll sideways, and any button,
# check box or label whose text is wider than its box. No server is contacted. Run it in an STA process.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-dialog-overflow-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir

# The files that the windows may carry that this probe cannot judge. Each entry names the reason.
$script:Exempt = @{}

function Get-WindowXaml {
    $docs = New-Object System.Collections.Generic.List[object]
    foreach ($name in 'MainWindow.xaml', 'WorkbenchWindow.xaml') {
        $docs.Add([pscustomobject]@{ Source = $name; Text = (Expand-SharedStyle (Get-Content -LiteralPath (Join-Path $root $name) -Raw)) })
    }
    $nodes = $script:LauncherAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and ([string]$n.Value) -match '^\s*<Controls:MetroWindow' }, $true)
    foreach ($n in $nodes) { $docs.Add([pscustomobject]@{ Source = ('start-apppackager.ps1:{0}' -f $n.Extent.StartLineNumber); Text = (Expand-SharedStyle ([string]$n.Value)) }) }
    return $docs
}

function Test-ScrollsSideways {
    param($Element)
    if ($Element -is [System.Windows.Controls.ScrollViewer]) { return ($Element.HorizontalScrollBarVisibility -ne [System.Windows.Controls.ScrollBarVisibility]::Disabled) }
    return ($Element -is [System.Windows.Controls.DataGrid] -or $Element -is [System.Windows.Controls.ListView] -or $Element -is [System.Windows.Controls.TreeView] -or $Element -is [System.Windows.Controls.ListBox] -or $Element -is [System.Windows.Controls.TextBox])
}

function Get-OverflowFinding {
    param([Parameter(Mandatory)]$Window)
    $found = New-Object System.Collections.Generic.List[string]
    $rootElement = [System.Windows.FrameworkElement]$Window.Content
    $walk = {
        param($node, $clipper)
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
        for ($i = 0; $i -lt $count; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($node, $i)
            if ($child -is [System.Windows.Controls.Primitives.ScrollBar]) { continue }
            $next = $clipper
            if ($child -is [System.Windows.FrameworkElement]) {
                if ($child.IsVisible -and $child.ActualWidth -gt 0) {
                    $id = $(if ($child.Name) { $child.Name } else { $child.GetType().Name })
                    $limit = $null
                    if ($clipper -is [System.Windows.Controls.ScrollViewer]) { $limit = $clipper.ViewportWidth } elseif ([object]::ReferenceEquals($clipper, $rootElement)) { $limit = $clipper.ActualWidth }
                    if ($null -ne $limit -and $limit -gt 0) {
                        $box = $child.TransformToAncestor($clipper).TransformBounds((New-Object System.Windows.Rect(0, 0, $child.ActualWidth, $child.ActualHeight)))
                        if ($box.Right -gt $limit + 1.5 -and $child.Name -ne 'PART_Border' -and -not ($child -is [System.Windows.Controls.Primitives.ScrollBar])) { $found.Add(('{0} ends at {1}; the window or scroll area is {2} wide' -f $id, [int]$box.Right, [int]$limit)) }
                    }
                    $tight = ($child.DesiredSize.Width - $child.Margin.Left - $child.Margin.Right) - $child.ActualWidth
                    if ($tight -gt 1.5 -and ($child -is [System.Windows.Controls.Button] -or $child -is [System.Windows.Controls.CheckBox] -or $child -is [System.Windows.Controls.RadioButton] -or $child -is [System.Windows.Controls.Label] -or ($child -is [System.Windows.Controls.TextBlock] -and $child.TextWrapping -eq [System.Windows.TextWrapping]::NoWrap -and $child.TextTrimming -eq [System.Windows.TextTrimming]::None -and $child.Text))) {
                        $found.Add(('{0} needs {1} and has {2}' -f $id, [int]$child.DesiredSize.Width, [int]$child.ActualWidth))
                    }
                }
                # The content of a scroll area is judged against its viewport; the parts of its template are not.
                if ($child -is [System.Windows.Controls.ScrollContentPresenter] -and $child.TemplatedParent -is [System.Windows.Controls.ScrollViewer] -and -not (Test-ScrollsSideways $child.TemplatedParent) -and $child.TemplatedParent.TemplatedParent -isnot [System.Windows.Controls.Primitives.TextBoxBase]) { $next = $child.TemplatedParent }
                elseif ($child -is [System.Windows.Controls.ScrollViewer]) { $next = $null }
                elseif (Test-ScrollsSideways $child) { $next = $null }
            }
            & $walk $child $next
        }
    }
    if ($rootElement) { & $walk $rootElement $rootElement }
    return $found
}

$total = 0
$problems = New-Object System.Collections.Generic.List[string]
try {
    foreach ($doc in @(Get-WindowXaml)) {
        $window = $null
        try {
            [xml]$xml = $doc.Text
            $window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
        }
        catch { Write-Host ('  skip  {0}: {1}' -f $doc.Source, $_.Exception.Message); continue }
        $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        $window.Left = 40; $window.Top = 40
        $window.ShowInTaskbar = $false
        try {
            $window.Show()
            $window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
            $window.UpdateLayout()
            $findings = @(Get-OverflowFinding -Window $window)
            $total++
            if ($findings.Count) { foreach ($f in $findings) { $problems.Add(('{0} ({1}x{2}): {3}' -f $doc.Source, [int]$window.ActualWidth, [int]$window.ActualHeight, $f)) } }
            else { Write-Host ('  ok  {0} ({1}x{2})' -f $doc.Source, [int]$window.ActualWidth, [int]$window.ActualHeight) }
        }
        finally { try { $window.Close() } catch { $null = $_ } }
    }
    if ($total -lt 15) { throw ('FAIL: only {0} windows were judged' -f $total) }
    if ($problems.Count) { $problems | ForEach-Object { Write-Host ('  FOUND ' + $_) }; throw ('FAIL: {0} cut-off control(s)' -f $problems.Count) }
    Write-Host ''
    Write-Host ('PASS: dialog overflow probe, {0} windows' -f $total)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
