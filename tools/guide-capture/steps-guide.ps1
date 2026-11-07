# Capture steps for the user guide of AppPackager. Runs in the interactive session of a lab account.
# Reads the ConfigMgr site and downloads one vendor installer to a local lab folder. It never publishes:
# the File Share Root points to a server that does not exist, so Publish to ConfigMgr stops at its share
# check before it copies content or touches the site. Every confirmation of a change is answered No or Cancel.
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$dl = 'C:\guide\dl'
# Close the browser window of an earlier CVE report, so that it cannot match a lookup.
Add-Type -AssemblyName UIAutomationClient
foreach ($old in @([System.Windows.Automation.AutomationElement]::RootElement.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition)) | Where-Object { $_.Current.Name -like 'AppPackager CVE report*' }) { try { $old.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { $null = $_ } }
$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
Size-Win $win 1440 900
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
$work = Join-Path $OutDir 'work'; New-Item -ItemType Directory $work -Force | Out-Null
$winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
function WinList { (@($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -notlike 'GlowWindow*' } | ForEach-Object { $_.Current.Name }) -join ' ; ' }
function Top-Dialog {
    @($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -notlike 'AppPackager*v20??.??.??.????' -and $_.Current.Name -notin 'Program Manager', '' -and $_.Current.Name -notlike 'Administrator*' -and $_.Current.Name -notlike 'GlowWindow*' -and $_.Current.Name -notlike '*Edge*' } | Select-Object -First 1
}
function Idle([int]$sec = 300) {
    Start-Sleep -Seconds 2
    if (-not (Wait-Until { $b = Get-El $win 'btnCheckLatest' -Quiet; $b -and $b.Current.IsEnabled } $sec)) { throw "Timeout: the app stayed busy for $sec seconds." }
    Start-Sleep -Seconds 2
}
function Status { Get-ElText (Get-El $win 'txtStatus') }
function Wait-Dialog([int]$sec = 20) { if (-not (Wait-Until { [bool](Top-Dialog) } $sec)) { throw ('No dialog. Windows: ' + (WinList)) }; Top-Dialog }
# Capture the top dialog and click a named control in it. A missing control fails the capture.
function Answer([string]$shot, [string]$button) {
    $d = Wait-Dialog; Start-Sleep -Milliseconds 800; Shot $d $shot; Dump-Tree $d $shot
    $b = Get-El $d $button -Quiet
    if (-not $b) { throw "No $button in $shot. Nothing was answered. Windows: $(WinList)" }
    Invoke-El $b; Start-Sleep -Seconds 1
}
function Check-Row([int]$index) { Click-At 255 (110 + 28 * $index); Start-Sleep -Milliseconds 400 }
function Close-Win($w) { $w.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close(); Start-Sleep -Seconds 2 }

Section 'setup' {
    $s = Find-Last 'Setup' 20; Shot $s 'setup-first-run'; Dump-Tree $s 'setup'
    Toggle-El (Get-El $s 'chkWizIntune'); Start-Sleep -Milliseconds 600; Shot $s 'setup-intune'
    Toggle-El (Get-El $s 'chkWizWsus'); Start-Sleep -Milliseconds 600; Shot $s 'setup-wsus'
    Toggle-El (Get-El $s 'chkWizIntune'); Toggle-El (Get-El $s 'chkWizWsus'); Start-Sleep -Milliseconds 500
    Set-Text (Get-El $s 'txtWizSC') $Capture.SiteCode; Set-Text (Get-El $s 'txtWizProvider') $Capture.SMSProvider
    Set-Text (Get-El $s 'txtWizFS') '\\nosuchserver.contoso.com\sources$'; Set-Text (Get-El $s 'txtWizDL') $dl
    Shot $s 'setup-configmgr-filled'
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
    Shot $win 'main-first-start'
}
Section 'grid' {
    Set-Text (Get-El $win 'txtGridFilter') '7-zip'; Start-Sleep -Seconds 1; Shot $win 'grid-filter'
    Set-Text (Get-El $win 'txtGridFilter') 'Adobe'; Start-Sleep -Seconds 1; Shot $win 'grid-filter-vendor'
    Set-Text (Get-El $win 'txtGridFilter') ''; Start-Sleep -Seconds 1
    Toggle-El (Get-El $win 'toggleDebugCols'); Start-Sleep -Seconds 1; Shot $win 'grid-debug-columns'; Toggle-El (Get-El $win 'toggleDebugCols'); Start-Sleep -Seconds 1
    Set-Text (Get-El $win 'txtComment') 'Lab run'; Start-Sleep -Milliseconds 500; Shot $win 'sidebar-comment'; Set-Text (Get-El $win 'txtComment') ''
}
Section 'options' {
    Invoke-El (Get-El $win 'btnOptions'); $o = Find-Last 'Options' 15
    $names = @('configmgr', 'packager', 'one-click', 'product-filter', 'script-signing', 'wsus', 'about'); $i = 0
    $items = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem))))
    if ($items.Count -ne $names.Count) { throw "Options has $($items.Count) pages, not $($names.Count)." }
    foreach ($it in $items) { Select-El $it; Start-Sleep -Milliseconds 1000; Shot $o ('options-' + $names[$i]); $i++ }
    Invoke-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
}
Section 'check latest' {
    Check-Row 0; Shot $win 'grid-row-selected'
    Invoke-El (Get-El $win 'btnCheckLatest'); Start-Sleep -Seconds 3; Shot $win 'check-latest-running'
    Idle 300; Log ('status: ' + (Status)); Shot $win 'check-latest-done'
}
Section 'check configmgr' {
    Invoke-El (Get-El $win 'btnCheckMECM'); Start-Sleep -Seconds 4; Shot $win 'check-configmgr-running'
    Idle 300; Log ('status: ' + (Status)); Shot $win 'check-configmgr-done'
}
Section 'stage' {
    Invoke-El (Get-El $win 'btnStage'); Start-Sleep -Seconds 4; Shot $win 'stage-running'
    Idle 600; Log ('status: ' + (Status)); Shot $win 'stage-done'
}
Section 'publish with an unreachable share' {
    Invoke-El (Get-El $win 'btnPackage'); Start-Sleep -Seconds 3
    Idle 300; Log ('status: ' + (Status)); Shot $win 'publish-share-unreachable'
}
Section 'cve' {
    Invoke-El (Get-El $win 'btnCheckCve'); Start-Sleep -Seconds 4; Shot $win 'cve-running'
    Idle 300; Log ('status: ' + (Status)); Shot $win 'cve-done'
    Invoke-El (Get-El $win 'btnCveReport')
    # The browser opens a few seconds later. Wait for its window; fail the capture when it does not show.
    if (-not (Wait-Until { [bool](@($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -like '*CVE report*' }) } 60)) { throw ('The CVE report window did not show. Windows: ' + (WinList)) }
    Start-Sleep -Seconds 4
    # Dismiss the restore bubble that the browser shows after an earlier forced close.
    Focus-Win (@($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -like '*CVE report*' } | Select-Object -First 1); Send-Keys '{ESC}'; Click-At 1816 105; Start-Sleep -Milliseconds 800
    $edge = @($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -like '*CVE report*' } | Select-Object -First 1
    Shot $edge 'cve-report-browser'; Close-Win $edge
    Shot $win 'cve-report-written'
}
Section 'workbench' {
    Invoke-El (Get-El $win 'btnWorkbench'); $w = Find-Last 'Application Workbench' 120; Start-Sleep -Seconds 3
    Shot $w 'workbench-application'
    $names = @('install', 'detection', 'requirements', 'timing', 'source-files', 'review'); $n = 0
    foreach ($page in 'Install & uninstall', 'Detection', 'Requirements & variants', 'Timing & execution', 'Source files', 'Review & build') {
        $it = Find-Role $w 'ListItem' $page
        if (-not $it) { throw "No Workbench page $page" }
        Select-El $it; Start-Sleep -Milliseconds 1200; Shot $w ('workbench-' + $names[$n]); $n++
    }
    Close-Win $w
}
Section 'add installer' {
    Invoke-El (Get-El $win 'btnAddInstaller'); $d = Wait-Dialog; Start-Sleep -Seconds 1; Shot $d 'add-installer-dialog'
    Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 2
}
Section 'convert' {
    Invoke-El (Get-El $win 'btnConvertCm'); $d = Find-Last 'Convert ConfigMgr App to Intune' 15; Start-Sleep -Seconds 1; Shot $d 'convert-dialog'
    $edit = @($d.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)))) | Select-Object -First 1
    Set-Text $edit '7-Zip (x64 edition) - 26.03'; Start-Sleep -Milliseconds 500
    Invoke-El (Find-Role $d 'Button' 'Preview'); Start-Sleep -Seconds 5
    if (-not (Wait-Until { (Status) -like 'Preview finished*' } 60)) { throw ('The Preview did not finish. Status: ' + (Status)) }
    Shot $win 'convert-preview'
}
Section 'discover' {
    Invoke-El (Get-El $win 'btnDiscover'); Start-Sleep -Seconds 4; Shot $win 'discover-running'
    Idle 300; Log ('status: ' + (Status))
    $d = Top-Dialog
    if ($d) {
        Shot $d 'discover-dialog'; Dump-Tree $d 'discover-dialog'
        $no = Get-El $d 'No' -Quiet; if (-not $no) { $no = Get-El $d 'Cancel' -Quiet }; if (-not $no) { $no = Get-El $d 'Close' -Quiet }
        if ($no) { Invoke-El $no } else { Focus-Win $d; Send-Keys '{ESC}' }
        Start-Sleep -Seconds 2
    }
    Shot $win 'discover-done'
}
Section 'one click' {
    Invoke-El (Get-El $win 'btnFullRun'); Start-Sleep -Seconds 5; Log ('windows: ' + (WinList))
    $d = Top-Dialog
    if ($d) {
        Shot $d 'one-click-not-configured'; Dump-Tree $d 'one-click-not-configured'
        $c = Get-El $d 'Cancel' -Quiet; if (-not $c) { $c = Get-El $d 'Close' -Quiet }
        if (-not $c) { throw 'The One Click plan has no Cancel or Close button. Nothing was answered.' }
        Invoke-El $c; Start-Sleep -Seconds 2
    } else { Log 'no One Click window' }
}
Section 'light' {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2; Shot $win 'light-main'
    Invoke-El (Get-El $win 'btnOptions'); $o = Find-Last 'Options' 15; Start-Sleep -Seconds 1; Shot $o 'light-options'; Invoke-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
}
Stop-App 'start-apppackager.ps1'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
