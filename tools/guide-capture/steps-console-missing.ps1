# Test copy only: the console detection reports "not found", as on a computer without the console.
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$f = Join-Path $appDir 'start-apppackager.ps1'
$t = [IO.File]::ReadAllText($f)
$start = $t.IndexOf('function Invoke-DetectConfigMgrConsole')
$hit = if ($start -ge 0) { $t.IndexOf('$result.Found = $true', $start) } else { -1 }
$patched = if ($hit -ge 0) { $t.Substring(0, $hit) + '$result.Found = $false' + $t.Substring($hit + '$result.Found = $true'.Length) } else { $t }
if ($patched -eq $t) { throw 'The console detection was not found. The missing-console test cannot run.' }
[IO.File]::WriteAllText($f, $patched, (New-Object Text.UTF8Encoding($true)))
$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
Size-Win $win 1440 900
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
$winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
$s = Find-Last 'Setup' 20
if (-not (Wait-Until { [bool](Get-El $s 'btnWizSave' -Quiet) } 20)) { throw 'The Setup window has no Save button.' }
Start-Sleep -Seconds 2
Set-Text (Get-El $s 'txtWizSC') $Capture.SiteCode; Set-Text (Get-El $s 'txtWizProvider') $Capture.SMSProvider
Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
Shot $win 'console-missing-main'
# Without the console, the three buttons that need it are disabled. Fail the capture when one is not.
foreach ($id in 'btnCheckMECM', 'btnPackage', 'btnConvertCm') {
    if ((Get-El $win $id).Current.IsEnabled) { throw "$id is enabled without the console." }
}
Move-To (Get-El $win 'btnCheckMECM'); Start-Sleep -Seconds 2; Shot $win 'console-missing-tooltip' -Hover
Invoke-El (Get-El $win 'btnOptions'); $o = Find-Last 'Options' 15; Start-Sleep -Seconds 2; Shot $o 'console-missing-options'
Invoke-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 1
Stop-App 'start-apppackager.ps1'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
