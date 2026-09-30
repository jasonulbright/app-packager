#Requires -Version 5.1
# Headless probe for the One Click plan window. Builds the real window from
# stubbed preferences, history and grid rows; checks the plan rows, the count
# line, the blocking line, a destination change, and the plan-only report.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerOneClick.psd1') -Force

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('Show-OneClickPlanDialog', 'Get-OneClickPlanApps', 'New-AppFlowPanel', 'ConvertTo-DeploymentTargetName')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw "function $name not found" }
    # Published globally, the way the shell publishes them for its closure handlers.
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$dataRoot = Join-Path $env:TEMP ('ap-oneclick-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $dataRoot -Force)

try {
    $script:Prefs = [pscustomobject]@{
        SiteCode = 'MCM'; FileShareRoot = '\\server\share'; DownloadRoot = (Join-Path $dataRoot 'dl'); ProviderMachineName = ''
        Systems = [pscustomobject]@{ ConfigMgr = $true; Intune = $false; Wsus = $true }
        DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $true } }
        Intune = [pscustomobject]@{ DeploymentTarget = 'MECMAndWSUS'; PublishToIntune = $false; TenantId = ''; ClientId = ''; ClientSecretProtected = '' }
        Wsus = [pscustomobject]@{ ServerName = 'wsus01'; ApprovalGroup = 'Pilot' }
        ContentDistribution = [pscustomobject]@{ AutoDistribute = $false; DPGroupName = ''; DeployToTestCollection = $false; TestCollectionName = '' }
        AppFlow = [pscustomobject]@{
            Tracked = @('package-alpha', 'package-beta', 'package-gamma'); Action = 'StageAndPackage'; ForceOnLaunch = $false
            CadenceOverrides = [pscustomobject]@{}
            Destinations = [pscustomobject]@{ 'package-gamma' = [pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $false } }
            DefaultDestinations = [pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $false }
            OnExisting = 'Skip'
        }
    }
    $script:History = @{}
    [void](Set-OneClickPublishedVersion -History $script:History -PackagerName 'package-beta' -Destination 'WSUS' -Version '2.0' -Id 'guid-beta')
    function global:Read-PackagerHistory { return $script:History }
    function global:Get-PackagerMetadata { param($Path) [pscustomobject]@{ UpdateCadenceDays = 7 } }
    function global:Add-LogLine { param([string]$Message) $script:Log += $Message }
    function global:Add-LogSeparator { }
    function global:Install-TitleBarDragFallback { param($Window) }
    function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
    function global:Get-AppLogFolder { Join-Path $dataRoot 'Logs' }
    $script:Log = @()

    $rows = @(
        [pscustomobject]@{ Script = 'package-alpha.ps1'; FullPath = 'C:\x\package-alpha.ps1'; Application = 'Alpha'; Vendor = 'V'; LatestVersion = '1.0'; CurrentVersion = '' }
        [pscustomobject]@{ Script = 'package-beta.ps1'; FullPath = 'C:\x\package-beta.ps1'; Application = 'Beta'; Vendor = 'V'; LatestVersion = '2.0'; CurrentVersion = '' }
        [pscustomobject]@{ Script = 'package-gamma.ps1'; FullPath = 'C:\x\package-gamma.ps1'; Application = 'Gamma'; Vendor = 'V'; LatestVersion = ''; CurrentVersion = '' }
    )

    $probeOwner = New-Object System.Windows.Window
    $probeOwner.Width = 10; $probeOwner.Height = 10
    $probeOwner.WindowStartupLocation = 'Manual'
    $probeOwner.Left = -32000; $probeOwner.Top = -32000

    $assert = { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw ('FAIL: ' + $Message) } }
    $script:PassedChecks = 0
    $ok = { param([string]$m) $script:PassedChecks += 1; Write-Host ('  ok  ' + $m) }

    Show-OneClickPlanDialog -Owner $probeOwner -Rows $rows -Action 'StageAndPackage' -Probe {
        param($p)
        & $assert ($p.Rows.Count -eq 3) 'one plan row per tracked application'
        & $ok 'three plan rows'

        $alpha = $p.Rows | Where-Object { $_.Packager -eq 'package-alpha' }
        $beta  = $p.Rows | Where-Object { $_.Packager -eq 'package-beta' }
        $gamma = $p.Rows | Where-Object { $_.Packager -eq 'package-gamma' }
        & $assert ($alpha.Planned -eq 'Check, Stage, Publish to ConfigMgr+WSUS' -and $alpha.Include) 'a default row plans both default destinations'
        & $ok 'default destinations planned'
        & $assert ($beta.Planned -eq 'Check, Stage, Publish to ConfigMgr' -and $beta.Reason -eq 'WSUS: 2.0 already published' -and $beta.LastPublished -eq 'WSUS 2.0') 'the duplicate guard drops the destination that has the version'
        & $ok 'duplicate guard shown'
        & $assert (-not $gamma.Include -and $gamma.Reason -eq 'no destination selected') 'a row without destinations is excluded'
        & $ok 'row without destinations excluded'
        & $assert ($p.Summary.Text -eq '2 application(s), 1 skipped: 2 to ConfigMgr, 1 to WSUS, 0 to Intune') ('the count line reads: ' + $p.Summary.Text)
        & $ok 'count line'
        & $assert ($p.Scope.Text -match 'WSUS: approved for group Pilot' -and $p.Scope.Text -match 'ConfigMgr: application only') 'the scope line names where each destination lands'
        & $ok 'scope line'
        & $assert ($p.Blocking.Visibility -eq 'Collapsed') 'nothing blocks a ready destination'
        & $ok 'no blocking line'
        & $assert ($p.Run.IsEnabled) 'Run is enabled with included rows'
        & $ok 'Run enabled'

        # The operator gives Gamma a destination for this run only.
        $gamma.WSUS = $true
        & $p.Refresh
        & $assert ($gamma.Include -and $gamma.Planned -eq 'Check, Stage, Publish to WSUS') 'a destination box checked in the window plans that destination'
        & $ok 'destination change replanned'
        & $assert ($p.Summary.Text -eq '3 application(s): 2 to ConfigMgr, 2 to WSUS, 0 to Intune') ('the count line follows: ' + $p.Summary.Text)
        & $ok 'count line follows'
        & $assert ($null -eq $script:Prefs.AppFlow.Destinations.PSObject.Properties['package-alpha']) 'the saved preferences are untouched'
        & $ok 'preferences untouched'

        # The operator clears Include of one row.
        $gamma.Include = $false; $gamma.IncludeTouched = $true
        & $p.Refresh
        & $assert ($p.Summary.Text -eq '2 application(s), 1 skipped: 2 to ConfigMgr, 1 to WSUS, 0 to Intune') ('the count line follows Include: ' + $p.Summary.Text)
        & $ok 'count line follows Include'
        $gamma.Include = $true
        & $p.Refresh

        # Force publishes the version WSUS already has.
        $p.Force.IsChecked = $true
        & $p.Refresh
        & $assert ($beta.Planned -eq 'Check, Stage, Publish to ConfigMgr+WSUS') 'Force ignores the duplicate guard'
        & $ok 'force replanned'
        $p.Force.IsChecked = $false
        & $p.Refresh

        # Plan only writes the report without a run.
        $p.State.Started = Get-Date
        $paths = & $p.WriteReport (Get-Date) $false
        & $assert ((Test-Path $paths.MarkdownPath) -and (Test-Path $paths.JsonPath)) 'the plan report is written'
        $md = Get-Content $paths.MarkdownPath -Raw
        & $assert ($md -match '\| Alpha \| 1\.0 \| Planned \|' -and $md -match '\| Beta \| 2\.0 \| Planned \|') 'the plan report lists the included rows as planned'
        & $assert ($md -notmatch 'Gamma' -or $gamma.Include) 'the report lists only included rows'
        & $ok 'plan-only report'
        & $assert ($p.Dialog.FindName('btnOpenReport').IsEnabled) 'Open report is enabled after a report'
        & $ok 'open report enabled'

        # A blocked destination is named when a row selects it.
        $script:Prefs.Wsus.ServerName = ''
        & $p.Refresh
        & $assert ($p.Blocking.Visibility -eq 'Visible' -and $p.Blocking.Text -match 'WSUS: WSUS server not set') 'the blocking line names the destination and the reason'
        & $ok 'blocking line'

        # Shown, the window replans only on an operator change: the check
        # boxes that a refresh regenerates must not start another refresh.
        # A refresh loop starves the dispatcher, so the counter itself ends
        # the loop (State.Running stops the box handler) and the frame.
        $script:SummaryCalls = 0
        $script:SummaryModule = Get-Module AppPackagerOneClick
        $script:ProbeState = $p.State
        $script:ProbeFrame = New-Object System.Windows.Threading.DispatcherFrame
        function global:Get-OneClickPlanSummary {
            param($Plan, $Prefs, $Action)
            $script:SummaryCalls++
            if ($script:SummaryCalls -gt 20) { $script:ProbeState.Running = $true; $script:ProbeFrame.Continue = $false }
            & $script:SummaryModule { param($a, $b, $c) Get-OneClickPlanSummary -Plan $a -Prefs $b -Action $c } $Plan $Prefs $Action
        }
        try {
            $p.Dialog.Left = -32000; $p.Dialog.Top = -32000; $p.Dialog.WindowStartupLocation = 'Manual'
            $p.Dialog.Show()
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds(2)
            $frame = $script:ProbeFrame
            $timer.Add_Tick({ $timer.Stop(); $frame.Continue = $false }.GetNewClosure())
            $timer.Start()
            [System.Windows.Threading.Dispatcher]::PushFrame($script:ProbeFrame)
            $timer.Stop()
            $p.State.Running = $false
            $p.Dialog.Close()
        }
        finally { Remove-Item -Path function:global:Get-OneClickPlanSummary -ErrorAction SilentlyContinue }
        & $assert ($script:SummaryCalls -le 2) ('the shown window replans without an operator change: {0} refreshes in 2 s' -f $script:SummaryCalls)
        & $ok 'shown window is idle'
    }

    # --- Run: the progress and done hooks reach the window ----------------
    # The pipeline is stubbed; it keeps the context the window hands it. The
    # hooks are then called the way the pipeline timer calls them.
    $script:Prefs.Wsus.ServerName = 'wsus01'
    $global:txtComment = New-Object System.Windows.Controls.TextBox
    $global:txtStatus = New-Object System.Windows.Controls.TextBlock
    foreach ($stub in 'Get-SevenZipPathForContext', 'Get-IntuneWinToolPathForContext', 'Get-RequirementsMapForContext', 'Get-VariantsMapForContext',
        'Get-CommandsMapForContext', 'Get-InstallModesMapForContext', 'Get-TitleModesMapForContext', 'Get-DefaultTitleModeForContext',
        'Get-IntunePublishConfigForContext', 'Get-WsusPublishConfigForContext', 'Get-WorkbenchRunPlanForContext',
        'Get-WorkbenchSigningPolicyJson', 'Get-WorkbenchSigningPolicyDigest') {
        Set-Item -Path ('function:global:' + $stub) -Value { param($Rows, $Target) $null }
    }
    function global:Confirm-LocalSourceFolders { param($Rows) $Rows }
    function global:Invoke-MultiAppPipeline { param($Operation, $Rows, $Context) $script:RunContext = $Context }
    $script:RunContext = $null
    Show-OneClickPlanDialog -Owner $probeOwner -Rows $rows -Action 'StageAndPackage' -Probe {
        param($p)
        $btnRun = $p.Run
        $btnRun.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $btnRun)))
        & $assert ($null -ne $script:RunContext -and $p.State.Running) 'Run starts the pipeline'
        $col = { param($n) $p.Dialog.FindName($n) }
        & $assert ((& $col 'colResultMecm').Visibility -eq 'Visible' -and (& $col 'colResultWsus').Visibility -eq 'Visible' -and (& $col 'colResultIntune').Visibility -eq 'Collapsed') 'the run shows a result column only for a destination in the run'
        & $assert (@('colMecm', 'colWsus', 'colIntune' | Where-Object { (& $col $_).Visibility -ne 'Collapsed' }).Count -eq 0) 'the result columns replace the destination boxes'
        $wraps = { param($c) @($c.ElementStyle.Setters | Where-Object { $_.Property -eq [System.Windows.Controls.TextBlock]::TextWrappingProperty -and $_.Value -eq [System.Windows.TextWrapping]::Wrap }).Count -gt 0 }
        & $assert (@('colResultMecm', 'colResultWsus', 'colReason' | Where-Object { -not (& $wraps (& $col $_)) }).Count -eq 0) 'a result longer than its column wraps'
        & $ok 'run view columns'
        $alpha = $p.Rows | Where-Object { $_.Packager -eq 'package-alpha' }
        $record = $script:RunContext.OneClickRows['package-alpha']
        $record.Step = 'publishing'; $record.Outcome = 'Published'; $record.ResultWSUS = 'published guid-alpha'; $record.IdWSUS = 'guid-alpha'
        & $script:RunContext.OneClickTick
        & $assert ($alpha.Step -eq 'publishing' -and $alpha.ResultWSUS -eq 'published guid-alpha') ('the tick copies progress into the grid: step {0}' -f $alpha.Step)
        & $ok 'progress reaches the grid'
        & $script:RunContext.OneClickDone ([pscustomobject]@{ Canceled = $false })
        & $assert (-not $p.State.Running -and $p.Summary.Text -like 'Complete: 1 published*') ('the done hook ends the run: ' + $p.Summary.Text)
        & $assert ($p.State.ReportPath -and (Get-Content -LiteralPath $p.State.ReportPath -Raw) -match '\| Alpha \| 1\.0 \| Published \|') 'the run report lists the outcome'
        & $ok 'run report written'
    }

    # --- One Click Settings panel: per-row boxes round trip ---------------
    $script:Prefs.Wsus.ServerName = 'wsus01'
    function global:Get-Packagers { param($Root) @(
        [pscustomobject]@{ Script = 'package-alpha.ps1'; Application = 'Alpha'; Vendor = 'V'; UpdateCadenceDays = $null }
        [pscustomobject]@{ Script = 'package-beta.ps1'; Application = 'Beta'; Vendor = 'V'; UpdateCadenceDays = 14 }
        [pscustomobject]@{ Script = 'package-gamma.ps1'; Application = 'Gamma'; Vendor = 'V'; UpdateCadenceDays = $null }
    ) }
    $PackagersRoot = $dataRoot
    $panel = New-AppFlowPanel
    $grid = $panel.Element.FindName('dgApps')
    $panelRows = @($grid.ItemsSource)
    $alphaRow = $panelRows | Where-Object { $_.Packager -eq 'package-alpha' }
    $gammaRow = $panelRows | Where-Object { $_.Packager -eq 'package-gamma' }
    & $assert ($alphaRow.ConfigMgr -and $alphaRow.WSUS -and -not $alphaRow.Intune) 'a row without its own selection shows the default boxes'
    & $assert (-not $gammaRow.ConfigMgr -and -not $gammaRow.WSUS -and -not $gammaRow.Intune) 'a row with its own selection shows it'
    & $ok 'panel rows show destinations'
    & $assert ($panel.Element.FindName('chkOneClickWsus').IsChecked -and -not $panel.Element.FindName('chkOneClickIntune').IsChecked) 'the default boxes follow the saved default'
    & $ok 'default boxes'

    $alphaRow.WSUS = $false
    $gammaRow.WSUS = $true; $gammaRow.ConfigMgr = $true
    $panel.Element.FindName('chkOneClickIntune').IsChecked = $true
    $panel.Element.FindName('cboOneClickOnExisting').SelectedIndex = 1
    & $panel.Commit
    $saved = $script:Prefs.AppFlow
    & $assert ($saved.DefaultDestinations.ConfigMgr -and $saved.DefaultDestinations.WSUS -and $saved.DefaultDestinations.Intune) 'the default destinations are saved from the boxes'
    & $assert ($saved.Destinations.PSObject.Properties['package-alpha'] -and -not $saved.Destinations.'package-alpha'.WSUS) 'a row that differs from the default is saved'
    & $assert ($saved.Destinations.PSObject.Properties['package-beta'] -and -not $saved.Destinations.'package-beta'.Intune) 'a row that kept the old boxes differs from the new default and is saved'
    & $assert ($saved.OnExisting -eq 'Overwrite') 'the existing-application policy is saved'
    & $assert ($script:Prefs.Intune.DeploymentTarget -eq 'MECMAndWSUS') 'the nearest single target is kept for the older code paths'
    & $ok 'panel commit'
    Write-Host ''
    Write-Host ('PASS: One Click plan probe, {0} checks' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $dataRoot -Recurse -Force -ErrorAction SilentlyContinue
}
