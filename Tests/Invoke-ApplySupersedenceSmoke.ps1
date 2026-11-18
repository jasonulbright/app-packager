#Requires -Version 5.1
# Window probe of the Apply Supersedence button: the main window loads with the button in the sidebar, in Title Case, inside the
# sidebar scroll area, with enough contrast in the dark and in the light theme. The button is enabled when ConfigMgr is ready and
# the supersedence mode is on, and it is disabled with the reason in its tooltip when ConfigMgr is not in use or the mode is Off.
# The click handler (Start-ApplySupersedence) runs against a stub of the engine entry point: it passes the newest stage manifest,
# the site, the title mode and the fallback identifier of each checked row, shows the result in the row and in the log, and
# starts no run while another run is in progress. No server is contacted and no preference file of the repository is written.
#   -ShotFolder   saves a maximized screenshot of the main window in each theme to that folder.
# SUITE_VISIBLE_WINDOWS=1 puts the window at 80,80 (the visible lab run).
param([string]$ShotFolder = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-apply-supersedence-' + [guid]::NewGuid().ToString('N'))
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

function New-ProbePrefs {
    param([bool]$ConfigMgr = $true, [string]$Mode = 'Replace')
    [pscustomobject]@{
        Systems       = [pscustomobject]@{ ConfigMgr = $ConfigMgr; Intune = $false; Wsus = $false }
        SiteCode      = 'MCM'
        DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $true } }
        Intune        = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = ''; DeploymentTarget = 'ConfigMgr' }
        Wsus          = [pscustomobject]@{ ServerName = '' }
        Supersedence  = [pscustomobject]@{ Mode = $Mode; TitleOverrides = [pscustomobject]@{} }
    }
}

function Show-MainWindow {
    param([bool]$Dark)
    [xml]$xml = Expand-SharedStyle (Get-Content -LiteralPath (Join-Path $root 'MainWindow.xaml') -Raw)
    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    $toggle = $w.FindName('toggleTheme')
    $toggle.IsOn = $Dark
    [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($w, $(if ($Dark) { 'Dark.Steel' } else { 'Light.Blue' }))
    $global:ProbeToggle = $toggle
    $names = 'btnFullRun', 'btnCheckLatest', 'btnCheckMECM', 'btnDiscover', 'btnConvertCm', 'btnCheckCve', 'btnStage', 'btnPackage', 'btnApplySupersedence', 'btnPublishIntune', 'btnPublishWsus', 'btnWsusUpdates', 'btnCveReport', 'btnWorkbench', 'btnAddInstaller'
    $buttons = @($names | ForEach-Object { $w.FindName($_) })
    Initialize-SuiteTheme -Window $w -IsDarkGetter { $global:ProbeToggle.IsOn -eq $true } -WorkflowButtons $buttons -OptionsButtons @($w.FindName('btnOptions')) -LogLabel $w.FindName('lblLogOutput')
    Set-ButtonTheme -IsDark $Dark
    $visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    $w.WindowStartupLocation = 'Manual'
    if ($ShotFolder) { $w.Left = 0; $w.Top = 0 }
    else { $w.Left = $(if ($visible) { 80 } else { -32000 }); $w.Top = $(if ($visible) { 80 } else { -32000 }) }
    $w.ShowInTaskbar = ($visible -or [bool]$ShotFolder)
    $w.Show()
    if ($ShotFolder) { $w.WindowState = 'Maximized' }
    Wait-Layout $w 6
    return $w
}

try {
    if ($ShotFolder) { [void](New-Item -ItemType Directory -Path $ShotFolder -Force) }

    foreach ($theme in 'dark', 'light') {
        $dark = ($theme -eq 'dark')
        $w = Show-MainWindow -Dark $dark
        try {
            $button = $w.FindName('btnApplySupersedence')
            & $assert ($null -ne $button) ('[{0}] the main window has the button' -f $theme)
            & $assert ($button.Content -ceq 'Apply Supersedence') ('[{0}] the label is in Title Case' -f $theme)
            & $assert (-not [string]::IsNullOrWhiteSpace([string]$button.ToolTip)) ('[{0}] the button has a tooltip' -f $theme)
            $sidebarButtons = @($w.FindName('btnPackage'), $button, $w.FindName('btnPublishIntune'))
            $order = @($sidebarButtons | ForEach-Object { $_.TransformToAncestor($w).Transform((New-Object System.Windows.Point(0, 0))).Y })
            & $assert ($order[0] -lt $order[1] -and $order[1] -lt $order[2]) ('[{0}] the button sits between Publish to ConfigMgr and Publish to Intune' -f $theme)
            $box = $button.TransformToAncestor($w).TransformBounds((New-Object System.Windows.Rect(0, 0, $button.ActualWidth, $button.ActualHeight)))
            $scroll = $null; $p = [System.Windows.Media.VisualTreeHelper]::GetParent($button)
            while ($p -and -not ($p -is [System.Windows.Controls.ScrollViewer])) { $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p) }
            $scroll = $p
            & $assert ($null -ne $scroll -and $button.ActualWidth -le $scroll.ViewportWidth + 0.5) ('[{0}] the button fits the width of the sidebar' -f $theme)
            & $assert ($box.Right -le $w.ActualWidth + 0.5) ('[{0}] the button is inside the window' -f $theme)

            # Availability and tooltip for three states of the preferences, as Update-SidebarForSystems applies them.
            foreach ($case in @(
                    @{ Name = 'ready'; Prefs = (New-ProbePrefs); Enabled = $true; Tip = 'Set the supersedence relation' }
                    @{ Name = 'mode Off'; Prefs = (New-ProbePrefs -Mode 'Off'); Enabled = $false; Tip = 'Supersedence is Off' }
                    @{ Name = 'ConfigMgr not in use'; Prefs = (New-ProbePrefs -ConfigMgr $false); Enabled = $false; Tip = 'ConfigMgr is not in use' }
                )) {
                $state = Get-SidebarSystemState -Prefs $case.Prefs
                $button.IsEnabled = [bool]$state.ApplySupersedenceEnabled
                $button.ToolTip = $state.ApplySupersedenceToolTip
                [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($button, $true)
                Wait-Layout $w 2
                & $assert ($button.IsEnabled -eq $case.Enabled) ('[{0}] {1}: enabled is {2}' -f $theme, $case.Name, $case.Enabled)
                & $assert ([string]$button.ToolTip -like ('*' + $case.Tip + '*')) ('[{0}] {1}: the tooltip says "{2}"' -f $theme, $case.Name, $case.Tip)
                if (-not $case.Enabled) { & $assert ([System.Windows.Controls.ToolTipService]::GetShowOnDisabled($button)) ('[{0}] {1}: the tooltip shows on the disabled button' -f $theme, $case.Name) }
            }

            # An enabled button has enough contrast in the theme.
            $button.IsEnabled = $true
            Wait-Layout $w 2
            $weak = @(Get-TextContrast $w '' $button | Where-Object { $need = $(if ($_.Size -ge 18.66 -and $_.Bold) { 3.0 } else { 4.5 }); $_.Ratio -lt $need })
            & $assert ($weak.Count -eq 0) ('[{0}] the label has enough contrast {1}' -f $theme, (($weak | ForEach-Object { '{0} {1}' -f $_.Text, $_.Ratio }) -join '; '))
            & $assert ($button.FontWeight.ToOpenTypeWeight() -ge 700) ('[{0}] the label is bold like the other buttons' -f $theme)

            if ($ShotFolder) {
                $size = Save-WindowPng -Window $w -Path (Join-Path $ShotFolder ('apply-supersedence-{0}.png' -f $theme))
                Write-Host ('  shot  apply-supersedence-{0}.png  {1}' -f $theme, $size)
            }
        }
        finally { try { $w.Close() } catch { $null = $_ } }
    }

    # The click handler against a stub of the engine entry point.
    $global:LogLines = @()
    $global:ProbeCalls = @()
    $global:ProbeBusy = @()
    $global:ProbePipelineBusy = $false
    $global:txtStatus = [pscustomobject]@{ Text = '' }
    $global:dataGrid = [pscustomobject]@{ Items = ([pscustomobject]@{}) }
    $global:dataGrid.Items | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $global:ProbeRefresh++ }
    $global:ProbeRefresh = 0
    function global:Test-PipelineBusy { [bool]$global:ProbePipelineBusy }
    function global:Show-BusyIndicator { param($Title, $Step, [switch]$AllowCancel) $global:ProbeBusy += ('show:' + $Title) }
    function global:Set-BusyStep { param($Text) $global:ProbeBusy += ('step:' + $Text) }
    function global:Hide-BusyIndicator { $global:ProbeBusy += 'hide' }
    function global:Set-ActionButtonsEnabled { param([bool]$Enabled) $global:ProbeBusy += ('buttons:' + $Enabled) }
    function global:Get-SelectedRows { $global:ProbeRows }
    function global:Get-TitleModesMapForContext { @{ 'package-vendor-b' = 'NoVersion' } }
    function global:Get-DefaultTitleModeForContext { 'IncludeVersion' }
    function global:Find-NewestStageManifestForPackager { param($PackagerPath, $DownloadRoot) if ($PackagerPath -like '*vendor-c*') { return $null }; return ('{0}\{1}\stage-manifest.json' -f $DownloadRoot, (Split-Path $PackagerPath -Leaf)) }
    function global:Invoke-SupersedenceApply {
        param($ManifestPath, $SiteCode, $ProviderMachineName, $FallbackApplicationId, $TitleMode, $Settings, $CommandLine)
        $global:ProbeCalls += [pscustomobject]@{ Manifest = $ManifestPath; Site = $SiteCode; Provider = $ProviderMachineName; Fallback = $FallbackApplicationId; TitleMode = $TitleMode; Mode = $Settings.Mode }
        if ($FallbackApplicationId -eq 'catalog:package-vendor-a') { return [pscustomobject]@{ Ok = $true; Message = 'Vendor A 2.0 supersedes Vendor A 1.0 1.0.0'; Status = 'Applied'; RowStatus = 'Supersedence applied'; Title = 'Vendor A 2.0'; Log = @('Supersedence mode           : Replace (site default)') } }
        if ($FallbackApplicationId -eq 'catalog:package-vendor-b') { return [pscustomobject]@{ Ok = $true; Message = 'No older version found.'; Status = 'NoOlderVersion'; RowStatus = 'No older version'; Title = 'Vendor B'; Log = @() } }
        return [pscustomobject]@{ Ok = $false; Message = 'No stage manifest was found for this application. Stage and publish it first.'; Status = 'NotApplied'; RowStatus = 'Supersedence not applied'; Title = ''; Log = @() }
    }
    $global:Prefs = New-ProbePrefs
    $script:Prefs = $global:Prefs
    $script:Prefs | Add-Member -NotePropertyName DownloadRoot -NotePropertyValue 'D:\ap' -Force
    $script:Prefs | Add-Member -NotePropertyName ProviderMachineName -NotePropertyValue 'cm01.contoso.com' -Force
    function New-ProbeRow([string]$Script, [string]$Application) { [pscustomobject]@{ Script = $Script + '.ps1'; FullPath = ('C:\Packagers\' + $Script + '.ps1'); Application = $Application; Status = '' } }

    $global:ProbeRows = @()
    Start-ApplySupersedence
    & $assert (@($global:ProbeCalls).Count -eq 0 -and (($global:LogLines -join '|') -match 'No rows selected')) 'no checked row: nothing runs and the log says so'

    $global:ProbeRows = @((New-ProbeRow 'package-vendor-a' 'Vendor A'), (New-ProbeRow 'package-vendor-b' 'Vendor B'), (New-ProbeRow 'package-vendor-c' 'Vendor C'))
    $global:LogLines = @(); $global:ProbeBusy = @()
    Start-ApplySupersedence
    $calls = @($global:ProbeCalls)
    & $assert ($calls.Count -eq 3) 'each checked row gets one call of the engine entry point'
    & $assert ($calls[0].Fallback -ceq 'catalog:package-vendor-a' -and $calls[0].Site -ceq 'MCM' -and $calls[0].Provider -ceq 'cm01.contoso.com' -and $calls[0].Manifest -like '*package-vendor-a.ps1*stage-manifest.json') 'the call carries the newest stage manifest, the site, the provider and the fallback identifier'
    & $assert ($calls[0].TitleMode -ceq 'IncludeVersion' -and $calls[1].TitleMode -ceq 'NoVersion') 'the title mode is the one of the application, else the default'
    & $assert ($calls[0].Mode -ceq 'Replace') 'the call carries the settings of the window'
    & $assert ($calls[2].Manifest -ceq '') 'a row with no stage passes an empty manifest path to the engine, which refuses it'
    & $assert (@($global:ProbeRows | ForEach-Object { $_.Status }) -join '|' -ceq 'Supersedence applied|No older version|Supersedence not applied') 'the rows show the outcome of the step'
    & $assert ($global:ProbeRefresh -ge 3) 'the grid is refreshed for each row'
    & $assert ((($global:LogLines -join "`n") -match 'Supersedence: Vendor A 2\.0 supersedes Vendor A 1\.0 1\.0\.0') -and (($global:LogLines -join "`n") -match 'Supersedence not applied: No stage manifest was found')) 'the log holds the message of each row'
    & $assert ($global:txtStatus.Text -ceq 'Apply Supersedence finished: 1 applied, 1 with no older version, 1 not applied.') 'the status line counts the outcomes'
    $busy = $global:ProbeBusy -join ','
    & $assert ($busy -match '^buttons:False,show:Applying supersedence' -and $busy -match 'hide,buttons:True$') 'the busy display shows during the run, and the buttons come back'

    # A run in progress blocks the action, and a mode of Off with no override blocks it too.
    $global:ProbeCalls = @(); $global:LogLines = @(); $global:ProbePipelineBusy = $true
    Start-ApplySupersedence
    & $assert (@($global:ProbeCalls).Count -eq 0 -and (($global:LogLines -join '|') -match 'another run is in progress')) 'no second run starts while a run is in progress'
    $global:ProbePipelineBusy = $false
    $script:Prefs.Supersedence.Mode = 'Off'
    $global:LogLines = @()
    Start-ApplySupersedence
    & $assert (@($global:ProbeCalls).Count -eq 0 -and (($global:LogLines -join '|') -match 'Supersedence is Off') -and $global:txtStatus.Text -match 'Supersedence is Off') 'with the mode Off nothing runs and the reason shows in the log and the status line'

    Write-Host ''
    Write-Host ('PASS: Apply Supersedence probe, {0} checks' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
