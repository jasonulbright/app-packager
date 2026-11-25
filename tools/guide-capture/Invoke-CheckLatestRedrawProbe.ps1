<#
.SYNOPSIS
    Window proof that Check Latest updates the status cells of the checked row and never rebuilds the rows or the icon column of the grid.

.DESCRIPTION
    Runs in the interactive session of a lab account (start it with tools\Run-VisibleProbe.ps1 of the suite repository, under the capture
    lock). The run checks the first rows of the grid and starts Check Latest with a real click. While the run goes, it samples the
    UI Automation runtime ids of the row elements and of the icon Image elements of the visible rows every 150 ms. A rebuilt row or a
    rebuilt icon cell is a new element with a new runtime id. The run counts the changes: rows rebuilt, icons rebuilt, and samples with
    a rebuild. It reads the status texts at the end to prove that rows were checked. It ends with "REDRAW-ROWS=<n> REDRAW-ICONS=<n>" and
    "REDRAW-FAILS=<n>". It reads nothing from the site and publishes nothing.
#>
param([int]$Rows = 24, [int]$MaxSeconds = 420)
$ErrorActionPreference = 'Stop'
$GuideRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$appDir = $GuideRoot
$OutDir = Join-Path $GuideRoot 'ui\out'; New-Item -ItemType Directory $OutDir -Force | Out-Null
$me = Get-Process -Id $PID
('{0}|{1}' -f $me.Id, $me.StartTime.Ticks) | Add-Content -LiteralPath (Join-Path $PSScriptRoot 'run.pids')
. (Join-Path $PSScriptRoot 'guilib.ps1')
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-CaptureSiteData
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
$env:APP_PACKAGER_WORKBENCH_ROOT = Join-Path $GuideRoot 'wb'
Remove-Item $env:APP_PACKAGER_WORKBENCH_ROOT -Recurse -Force -ErrorAction SilentlyContinue
& (Join-Path $PSScriptRoot 'install-iconpack.ps1') -AppRoot $appDir | Out-Null
$titlePattern = 'AppPackager*v20??.??.??.????'
$fails = 0
function Say([string]$m) { Write-Output $m; Log $m }
function Items($Scope, $Type) { @($Scope.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::($Type))))) }
function Rid($El) { try { ($El.GetRuntimeId() -join '.') } catch { '' } }

$win = Start-App 'start-apppackager.ps1' $titlePattern
Size-Win $win 1440 1040
$win = Find-Top $titlePattern 10
$s = Find-Last 'Setup' 20
Set-Text (Get-El $s 'cm.siteCode') 'MCM'; Set-Text (Get-El $s 'cm.provider') 'cm01.contoso.com'
Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
$win = Find-Top $titlePattern 10
Focus-Win $win
$grid = Get-El $win 'dataGrid'
$null = Wait-Until { (Items $grid 'DataItem').Count -gt 10 } 60
Start-Sleep -Seconds 2

# Check the first rows through UI Automation.
$dataItems = Items $grid 'DataItem'
$checked = 0
foreach ($row in ($dataItems | Select-Object -First $Rows)) {
    $box = $row.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::CheckBox)))
    if (-not $box) { continue }
    $t = $box.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    if ($t.Current.ToggleState -ne 'On') { $t.Toggle() }
    $checked++
}
Say "rows checked for the run: $checked"
if ($checked -lt 5) { Say 'FAIL: fewer than 5 rows could be checked'; $fails++ }

# The state of the grid before the run: runtime ids of the visible rows and of their icons.
function Get-Snapshot {
    # Two UI Automation searches only, so that a sample is fast: the runtime ids of the visible rows and of the visible icons, by screen row.
    $g = Get-El $win 'dataGrid'
    $map = @{}
    foreach ($row in (Items $g 'DataItem')) {
        $r = $row.Current.BoundingRectangle
        if ($r.IsEmpty -or $r.Width -le 0) { continue }
        $map[[string][int]$r.Y] = [pscustomobject]@{ Row = (Rid $row); Icon = '' }
    }
    foreach ($img in (Items $g 'Image')) {
        $r = $img.Current.BoundingRectangle
        if ($r.IsEmpty -or $r.Width -le 0) { continue }
        $k = [string]([int]$r.Y - 4)
        foreach ($cand in @($k, [string]([int]$r.Y - 3), [string]([int]$r.Y - 5), [string]([int]$r.Y - 2), [string]([int]$r.Y - 6))) {
            if ($map.ContainsKey($cand)) { $map[$cand].Icon = (Rid $img); break }
        }
    }
    $map
}
$prev = Get-Snapshot
$rowRebuilds = 0; $iconRebuilds = 0; $badSamples = 0; $samples = 0
$cl = Get-El $win 'btnCheckLatest'
Say ('Check Latest enabled before: ' + $cl.Current.IsEnabled)
Invoke-El $cl
$start = Get-Date
Start-Sleep -Seconds 4
$wc = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
Say ('windows after the click: ' + ((@($root.FindAll('Descendants', $wc)) | Where-Object { $_.Current.Name -notlike 'GlowWindow*' } | ForEach-Object { $_.Current.Name }) -join ' ; '))
Say ('selected rows in the grid text: ' + (@(Items (Get-El $win 'dataGrid') 'CheckBox' | Where-Object { try { $_.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Current.ToggleState -eq 'On' } catch { $false } }).Count))
Shot-Screen 'after-click'
if ($env:REDRAW_DIAG) { New-Item -ItemType Directory $env:REDRAW_DIAG -Force | Out-Null; Copy-Item (Join-Path $OutDir 'after-click*.png') $env:REDRAW_DIAG -Force -ErrorAction SilentlyContinue }

$sawBusy = $false
while (((Get-Date) - $start).TotalSeconds -lt $MaxSeconds) {
    Start-Sleep -Milliseconds 150
    $now = Get-Snapshot
    $samples++
    $rb = 0; $ib = 0
    foreach ($k in $now.Keys) {
        if (-not $prev.ContainsKey($k)) { continue }
        if ($now[$k].Row -and $prev[$k].Row -and $now[$k].Row -ne $prev[$k].Row) { $rb++ }
        if ($now[$k].Icon -and $prev[$k].Icon -and $now[$k].Icon -ne $prev[$k].Icon) { $ib++ }
    }
    $rowRebuilds += $rb; $iconRebuilds += $ib; if ($rb -or $ib) { $badSamples++ }
    $prev = $now
    $btn = Get-El $win 'btnCheckLatest' -Quiet
    if ($btn -and -not $btn.Current.IsEnabled) { $sawBusy = $true }
    if ($sawBusy -and $btn -and $btn.Current.IsEnabled) { break }
    if (-not $sawBusy -and ((Get-Date) - $start).TotalSeconds -gt 40) { break }
}
Start-Sleep -Seconds 2
Shot $win 'check-latest-redraw-after'
$statusTexts = @(Items (Get-El $win 'dataGrid') 'Text' | ForEach-Object { $_.Current.Name } | Where-Object { $_ -match '^(Up to date|Update available|Latest retrieved|Not found in ConfigMgr|Current newer|Error)' })
Say ('samples: {0} over {1:N0} s, status texts of checked rows: {2}' -f $samples, ((Get-Date) - $start).TotalSeconds, $statusTexts.Count)
Say ("rows rebuilt: $rowRebuilds, icons rebuilt: $iconRebuilds, samples with a rebuild: $badSamples")
if (-not $sawBusy) { Say 'FAIL: the run never started (Check Latest button never disabled)'; $fails++ }
if ($statusTexts.Count -lt 3) { Say 'FAIL: fewer than 3 rows show a result'; $fails++ }
Get-ChildItem -LiteralPath (Join-Path $appDir 'Logs') -Filter *.log -File -ErrorAction SilentlyContinue | Copy-Item -Destination $OutDir -Force
Stop-App 'start-apppackager.ps1'
Write-Output ("REDRAW-ROWS=$rowRebuilds REDRAW-ICONS=$iconRebuilds")
Write-Output ("REDRAW-FAILS=$fails")
