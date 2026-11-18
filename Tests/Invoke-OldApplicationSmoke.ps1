#Requires -Version 5.1
# Window probe of "Use as old application": the dialog opens with the search text, lists what the read returns, links only a
# row that can be linked and only after the confirmation (the text shows the title and the version first), removes a link,
# says why a row cannot be linked, reports a failed search, keeps its text readable and its buttons in view in the dark and in
# the light theme, and never touches ConfigMgr (the three site blocks are fakes).
#   -ShotFolder   saves a maximized screenshot of the window in each theme to that folder.
# SUITE_VISIBLE_WINDOWS=1 puts the window at 80,80 (the visible lab run).
param([string]$ShotFolder = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$script:SharedStyleRootOverride = $root

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global
Import-Module (Join-Path $root 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking -Global
. (Join-Path $PSScriptRoot 'WindowMeasure.ps1')
. $script:MeasureFunctions

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('Expand-SharedStyle', 'Get-SharedStyleXaml', 'Show-OldApplicationDialog', 'Set-DialogBusyStatus', 'Update-UiNow')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw ('function not found: ' + $name) }
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}
function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
function global:Add-LogLine { param([string]$Message) $global:Log += $Message }
function global:Get-UserErrorText { param($ErrorRecord, [string]$Fallback) return [string]$ErrorRecord.Exception.Message }
function global:Show-ThemedMessage {
    param($Owner, [string]$Title, [string]$Message, [string]$Buttons, [string]$Icon)
    $global:Asked += [pscustomobject]@{ Title = $Title; Message = $Message; Buttons = $Buttons }
    return $global:Answer
}

$script:Checks = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:Checks += 1
    Write-Host ('PASS  ' + $Message)
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
$click = { param($Control) $Control.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }

function New-Row {
    param([string]$Title, [string]$Version, [string]$State, [string]$Reason, [bool]$CanLink, [bool]$CanClear)
    [pscustomobject]@{ Title = $Title; SoftwareVersion = $Version; State = $State; Reason = $Reason; CanLink = $CanLink; CanClear = $CanClear }
}

if ($ShotFolder) { [void](New-Item -ItemType Directory -Path $ShotFolder -Force) }
$visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')

foreach ($theme in 'dark', 'light') {
    $dark = ($theme -eq 'dark')
    $global:ProbeDark = $dark
    $global:Log = @(); $global:Asked = @(); $global:Answer = 'Yes'; $global:Linked = @(); $global:Unlinked = @(); $global:Searches = @()
    $rows = @(
        (New-Row 'Vendor App 1.5' '1.5.0' 'Linkable' '' $true $false)
        (New-Row 'Vendor App 1.0' '1.0.0' 'Linkable' '' $true $false)
        (New-Row 'Vendor App Finance' '1.2.0' 'OwnKeywords' 'It has keywords of its own, which a write would replace.' $false $false)
        (New-Row 'Vendor App Linked' '0.9.0' 'Linked' 'It already holds the marker of this line.' $false $true)
    )
    $global:Rows = $rows
    $search = { param($Text) $global:Searches += $Text; if ($Text -eq 'boom') { throw 'the site is down' }; $global:Rows }
    $link = { param($Name) $global:Linked += $Name; [pscustomobject]@{ Ok = $true; Message = "'$Name' is the old application of this line." } }
    $unlink = { param($Name) $global:Unlinked += $Name; [pscustomobject]@{ Ok = $true; Message = "The marker is removed from '$Name'." } }

    Show-OldApplicationDialog -Owner $null -NewTitle 'Vendor App 2.0' -NewVersion '2.0.0' -Marker 'APD-Line:catalog:vendor-app|default|x64' -SearchText 'Vendor App' -Search $search -Link $link -Unlink $unlink -Probe {
        param($p)
        $w = $p.Window
        Initialize-SuiteTheme -Window $w -IsDarkGetter { [bool]$global:ProbeDark }
        Set-DialogTheme -Dialog $w -IsDark $global:ProbeDark
        $w.WindowStartupLocation = 'Manual'
        if ($ShotFolder) { $w.Left = 0; $w.Top = 0 } else { $w.Left = $(if ($visible) { 80 } else { -32000 }); $w.Top = $(if ($visible) { 80 } else { -32000 }) }
        $w.ShowInTaskbar = ($visible -or [bool]$ShotFolder)
        $w.Show()
        if ($ShotFolder) { $w.WindowState = 'Maximized' }
        Wait-Layout $w 6
        try {
            $label = $theme
            & $assert ($w.Title -ceq 'Old Application') "[$label] the window title"
            & $assert ($p.Search.Text -eq 'Vendor App') "[$label] the search starts with the application name"
            & $assert ($p.Intro.Text -like "*'Vendor App 2.0' (version 2.0.0)*") "[$label] the intro names the new application and its version"
            & $assert ($p.Use.Content -ceq 'Use as Old Application' -and $p.Remove.Content -ceq 'Remove Link' -and $p.Find.Content -ceq 'Search') "[$label] the buttons carry their labels"

            & $click $p.Find
            Wait-Layout $w
            & $assert ($global:Searches[-1] -eq 'Vendor App' -and $p.Grid.Items.Count -eq 4) "[$label] Search reads the site with the text and lists four applications"
            & $assert ($p.Status.Text -eq '4 application(s) found; 2 can be linked.') "[$label] the status counts what can be linked"

            $p.Grid.SelectedIndex = 2
            & $click $p.Use
            & $assert ($global:Linked.Count -eq 0 -and $global:Asked.Count -eq 0 -and $p.Status.Text -like "*cannot be linked: It has keywords of its own*") "[$label] a row with keywords of its own is refused with the reason and nothing is asked"

            $p.Grid.SelectedIndex = 0
            $global:Answer = 'No'
            & $click $p.Use
            & $assert ($global:Asked.Count -eq 1 -and $global:Linked.Count -eq 0) "[$label] answering No writes nothing"
            & $assert ($global:Asked[0].Message.StartsWith("Use 'Vendor App 1.5' version 1.5.0 as the old application of 'Vendor App 2.0'?") -and $global:Asked[0].Buttons -eq 'YesNo') "[$label] the question shows the title and the version first"

            $global:Answer = 'Yes'
            & $click $p.Use
            & $assert ($global:Linked.Count -eq 1 -and $global:Linked[0] -eq 'Vendor App 1.5') "[$label] answering Yes links the selected application"
            & $assert ($p.Status.Text -like "'Vendor App 1.5' is the old application of this line.*" -and ($global:Log -join '|') -like "*Old application: 'Vendor App 1.5'*") "[$label] the status and the log say what was written"

            $p.Grid.SelectedIndex = 0
            & $click $p.Remove
            & $assert ($global:Unlinked.Count -eq 0 -and $p.Status.Text -like "*does not hold the marker of this line alone*") "[$label] Remove Link on an application without the marker does nothing"
            $p.Grid.SelectedIndex = 3
            & $click $p.Remove
            & $assert ($global:Unlinked.Count -eq 1 -and $global:Unlinked[0] -eq 'Vendor App Linked' -and $global:Asked[-1].Title -ceq 'Remove Link') "[$label] Remove Link asks and then takes the marker off"

            $p.Grid.SelectedIndex = -1
            & $click $p.Use
            & $assert ($p.Status.Text -like 'Select an application first.*') "[$label] no selection asks for one"

            $p.Search.Text = 'boom'
            & $click $p.Find
            & $assert ($p.Grid.Items.Count -eq 0 -and $p.Status.Text -like 'The search failed: the site is down*') "[$label] a failed search is reported"
            $p.Search.Text = '  '
            & $click $p.Find
            & $assert ($p.Status.Text -like 'Type part of the application name*') "[$label] an empty text asks for a name"
            $p.Search.Text = 'Vendor App'
            & $click $p.Find
            Wait-Layout $w

            $buttonBox = $p.Use.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $p.Use.ActualWidth, $p.Use.ActualHeight)))
            & $assert ($buttonBox.Bottom -le $w.ActualHeight + 0.5 -and $buttonBox.Right -le $w.ActualWidth + 0.5) "[$label] the buttons are inside the window"
            $weak = @(Get-TextContrast $w '' $w | Where-Object { $need = $(if ($_.Icon -or $_.Size -ge 24 -or ($_.Size -ge 18.66 -and $_.Bold)) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
            & $assert ($weak.Count -eq 0) ("[$label] every text has enough contrast " + (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
            if ($ShotFolder) {
                $size = Save-WindowPng -Window $w -Path (Join-Path $ShotFolder ('old-application-{0}.png' -f $theme))
                Write-Host ('  shot {0} {1}' -f $theme, $size)
            }
        }
        finally { $w.Close() }
    }
}

Write-Host ("`nPASS {0} checks." -f $script:Checks)
