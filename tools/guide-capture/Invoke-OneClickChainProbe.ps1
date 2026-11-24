<#
.SYNOPSIS
    Window proof of the Architecture column, the One Click plan from a state with no Check Latest, and the Application Title page.

.DESCRIPTION
    Runs in the interactive session of a lab account (start it with tools\Run-VisibleProbe.ps1 of the suite repository). Real mouse input
    opens Options and the One Click plan; UI Automation finds the elements and reads the result. The run publishes nothing and never
    starts the run: it closes the plan window with Close. The proof applications are two catalog packagers that need no installer file.
    The result lines go to the output; a failed check stops the run with a FAIL line.
#>
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
$prefsPath = Join-Path $appDir 'AppPackager.preferences.json'
$winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
function WinList { (@($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -notlike 'GlowWindow*' } | ForEach-Object { $_.Current.Name }) -join ' ; ' }
function Require([bool]$Condition, [string]$Message) {
    if (-not $Condition) { Shot-Screen ('fail-' + (Get-Date -Format 'HHmmss')); Log ('windows: ' + (WinList)); throw ('FAIL: ' + $Message) }
    Write-Output ('ok: ' + $Message); Log ('ok: ' + $Message)
}
function Texts($Scope) { @($Scope.FindAll('Descendants', [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name } | Where-Object { $_ }) }
function Items($Scope, $Type) { @($Scope.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::($Type))))) }
$titlePattern = 'AppPackager*v20??.??.??.????'
$tracked = @('package-7zip', 'package-putty')

# 1. First start: the setup wizard, then stop the app and track two applications in the preferences file.
$win = Start-App 'start-apppackager.ps1' $titlePattern
$s = Find-Last 'Setup' 30
Set-Text (Get-El $s 'cm.siteCode') 'MCM'; Set-Text (Get-El $s 'cm.provider') 'cm01.contoso.com'
Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
Stop-App 'start-apppackager.ps1'; Start-Sleep -Seconds 2
Require (Test-Path -LiteralPath $prefsPath) 'the setup wizard wrote the preferences file'
$prefs = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
if (-not $prefs.PSObject.Properties['AppFlow']) { $prefs | Add-Member -NotePropertyName AppFlow -NotePropertyValue ([pscustomobject]@{}) }
$prefs.AppFlow | Add-Member -NotePropertyName Tracked -NotePropertyValue $tracked -Force
$prefs.AppFlow | Add-Member -NotePropertyName DefaultDestinations -NotePropertyValue ([pscustomobject]@{ ConfigMgr = $true; WSUS = $false; Intune = $false }) -Force
($prefs | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $prefsPath -Encoding UTF8
# No history: no application has a version from an earlier check.
Remove-Item (Join-Path $dataDir 'app-history.json') -Force -ErrorAction SilentlyContinue

# 2. Second start: no Check Latest has run. Click One Click.
$win = Start-App 'start-apppackager.ps1' $titlePattern
Size-Win $win 1440 1040
$win = Find-Top $titlePattern 10
Focus-Win $win
Require ((Items (Get-El $win 'dataGrid') 'DataItem').Count -gt 0) 'the grid has rows and no Check Latest has run'
Click-El (Get-El $win 'btnFullRun'); Start-Sleep -Seconds 2
$plan = Find-Last 'One Click' 300
Start-Sleep -Seconds 2
Shot $plan 'q124-oneclick-plan'
$grid = Get-El $plan 'gridPlan'
Log ('plan grid: ' + ((Texts $grid) -join ' | '))
$planRows = @(Items $grid 'DataItem')
Require ($planRows.Count -ge 2) ('the plan has a row for each tracked application (' + $planRows.Count + ')')
foreach ($it in $planRows) {
    $rowText = (Texts $it) -join ' | '
    Log ('plan row: ' + $rowText)
    Require ($rowText -match '(?<![\d.])(\d+\.\d+[\d.]*)\s*\|') ('the plan row shows the version ' + $Matches[1])
    Require ($rowText -notmatch 'not checked|Publish after check') 'the plan row does not say not checked or publish after check'
}
$action = Get-ElText (Get-El $plan 'txtPlanAction')
Require ($action -match 'check, stage and publish') ('the plan runs the whole chain: ' + $action)
Invoke-El (Get-El $plan 'btnClose'); Start-Sleep -Seconds 2

# 3. Options: the Application Title page has no supersedence note; the grid has the column Architecture.
Click-El (Get-El $win 'btnOptions'); Start-Sleep -Seconds 2
$o = Find-LastLike 'Options*' 20
Size-Win $o 1600 1040
function Select-Node($o, [string]$Name, [string]$Parent) {
    $nodes = Items $o 'TreeItem'
    $p = $nodes | Where-Object { $_.Current.Name -eq $Parent } | Select-Object -First 1
    if ($p) { try { Expand-El $p } catch { $null = $_ }; Start-Sleep -Milliseconds 600 }
    $n = (Items $o 'TreeItem') | Where-Object { $_.Current.Name -eq $Name } | Select-Object -First 1
    if (-not $n) { throw "The Options tree has no $Name." }
    Click-El $n; Start-Sleep -Milliseconds 1500
}
Select-Node $o 'Application Title' 'Packaging'
Shot $o 'q124-title-page'
$titleText = (Texts $o) -join ' | '
Require ($titleText -match 'Add the version to each application name') 'the Application Title page is open'
Require ($titleText -notmatch '(?i)supersed') 'the Application Title page holds no supersedence text'
Select-Node $o 'All Applications' 'Applications'
Start-Sleep -Seconds 2
Shot $o 'q124-all-applications'
$headers = @(Items $o 'Header' | ForEach-Object { $_.Current.Name })
$allText = (Texts $o) -join ' | '
Log ('headers: ' + ($headers -join ' | '))
Require (($headers -contains 'Architecture') -or ($allText -match 'Architecture')) 'the grid has the column Architecture'
Require (($headers -notcontains 'Package Source') -and ($allText -notmatch 'Package Source')) 'the grid has no column Package Source'
Require (($headers -contains 'Install Scope') -or ($allText -match 'Install Scope')) 'the grid keeps the column Install Scope'
Invoke-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
Stop-App 'start-apppackager.ps1'
Write-Output 'PASS: q124 window proof'
