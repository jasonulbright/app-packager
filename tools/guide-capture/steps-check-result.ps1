# Window proof of the Check ConfigMgr result file and its export (Check ConfigMgr, the result file in the data folder, the row menu, the Save dialog).
# Runs in the interactive session of a lab account. It reads the ConfigMgr site and publishes nothing. Real mouse input opens the row menu and selects the entry;
# UI Automation only finds the elements and reads the result. Image: check-configmgr-export-menu.
# check before it copies content or touches the site. Every confirmation of a change is answered No or Cancel.
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
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
# Check a grid row through UI Automation, so that the display scaling cannot move the click, and verify the state.
function Check-Row([int]$index) {
    $rows = @((Get-El $win 'dataGrid').FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem))))
    if ($rows.Count -le $index) { throw 'The grid has no row ' + $index }
    $box = $rows[$index].FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::CheckBox)))
    if (-not $box) { throw 'The row has no check box.' }
    $toggle = $box.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    if ($toggle.Current.ToggleState -ne 'On') { $toggle.Toggle() }
    Start-Sleep -Milliseconds 400
    if ($toggle.Current.ToggleState -ne 'On') { throw 'The row is not checked.' }
}
function Close-Win($w) { $w.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close(); Start-Sleep -Seconds 2 }

Section 'setup' {
    $s = Find-Last 'Setup' 20; Shot $s 'setup-first-run'; Dump-Tree $s 'setup'
    Toggle-El (Get-El $s 'sys.intune'); Start-Sleep -Milliseconds 600; Shot $s 'setup-intune'
    Toggle-El (Get-El $s 'sys.wsus'); Start-Sleep -Milliseconds 600; Shot $s 'setup-wsus'
    Toggle-El (Get-El $s 'sys.intune'); Toggle-El (Get-El $s 'sys.wsus'); Start-Sleep -Milliseconds 500
    Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
    Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
    Shot $s 'setup-configmgr-filled'
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
    Shot $win 'main-first-start'
}
Section 'check configmgr and export' {
    foreach ($i in 0, 1, 2) { Check-Row $i }
    Invoke-El (Get-El $win 'btnCheckLatest'); Start-Sleep -Seconds 3; Idle 300; Log ('status after Check Latest: ' + (Status))
    Invoke-El (Get-El $win 'btnCheckMECM'); Start-Sleep -Seconds 4
    # The offer to hide the applications that the site lacks opens before the buttons come back; it is answered No, so the proof changes no preference.
    $end = (Get-Date).AddSeconds(300)
    while ((Get-Date) -lt $end) {
        $offer = Top-Dialog
        if ($offer) { Log ('dialog: ' + $offer.Current.Name); $no = Get-El $offer 'No' -Quiet; if ($no) { Invoke-El $no; Start-Sleep -Seconds 1 } }
        $b = Get-El $win 'btnCheckLatest' -Quiet
        if ($b -and $b.Current.IsEnabled) { break }
        Start-Sleep -Seconds 2
    }
    Log ('status after Check ConfigMgr: ' + (Status))
    $resultCsv = Join-Path $dataDir 'CheckConfigMgrResult.csv'
    $resultJson = Join-Path $dataDir 'CheckConfigMgrResult.json'
    if (-not (Test-Path -LiteralPath $resultCsv)) { throw "no result file at $resultCsv" }
    if (-not (Test-Path -LiteralPath $resultJson)) { throw "no result file at $resultJson" }
    Log ('result file: ' + @(Get-Content -LiteralPath $resultCsv).Count + ' lines'); Get-Content -LiteralPath $resultCsv -TotalCount 4 | ForEach-Object { Log $_ }
    $header = (Get-Content -LiteralPath $resultCsv -TotalCount 1).TrimStart([char]0xFEFF)
    if ($header -ne '"Application","Vendor","CMName","InstalledSoftwareName","Architecture","ConfigMgrVersion","LatestVersion","Status","Source","CheckedAtUtc"') { throw "wrong header: $header" }
    # A real right-click on a row opens the menu; a real click selects the entry.
    $grid = Get-El $win 'dataGrid'; $gr = $grid.Current.BoundingRectangle
    Click-At ([int]($gr.X + 260)) ([int]($gr.Y + 52))
    [NativeUi]::mouse_event(8, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(16, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 900
    Shot $win 'check-configmgr-export-menu' -Hover
    $entry = Get-El $root 'Export Check ConfigMgr results...' -Quiet
    if (-not $entry) { throw 'The row menu has no entry Export Check ConfigMgr results...' }
    Click-El $entry; Start-Sleep -Seconds 2
    $target = Join-Path $GuideRoot 'exported-result.csv'
    Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
    Send-Keys $target.Replace('{', '{{}').Replace('}', '{}}'); Send-Keys '{ENTER}'; Start-Sleep -Seconds 2
    if (-not (Test-Path -LiteralPath $target)) { Shot-Screen 'fail-export-dialog'; throw "The export wrote no file at $target. Windows: $(WinList)" }
    $a = (Get-Content -LiteralPath $resultCsv -Raw); $b = (Get-Content -LiteralPath $target -Raw)
    if ($a -ne $b) { throw 'The exported file differs from the result file of the run.' }
    Log ('ok: the export wrote ' + @(Get-Content -LiteralPath $target).Count + ' lines, equal to the result file')
}
