# Measurement of the busy spinner with the icon list loaded. Runs in the interactive session of a lab account.
# It starts the app, checks the visible rows, starts Check Latest (the busy overlay with the spinner shows), and samples the
# process counters of the app once per second for 30 seconds. It publishes nothing and reads nothing from the site.
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class GuiRes { [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr h, uint flag); }
'@
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
Size-Win $win 1440 900
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
Section 'setup' {
    $s = Find-Last 'Setup' 20
    Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
    Set-Text (Get-El $s 'cm.fileShare') '\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 5
}
function Sample([string]$label) {
    $p = $script:proc; $p.Refresh()
    [pscustomobject]@{
        Label = $label; CpuSeconds = [math]::Round($p.TotalProcessorTime.TotalSeconds, 2); Handles = $p.HandleCount
        Gdi = [GuiRes]::GetGuiResources($p.Handle, 0); User = [GuiRes]::GetGuiResources($p.Handle, 1)
        WorkingSetMB = [math]::Round($p.WorkingSet64 / 1MB, 1); PrivateMB = [math]::Round($p.PrivateMemorySize64 / 1MB, 1); Threads = $p.Threads.Count
    }
}
Section 'idle baseline (10 s, no busy overlay)' {
    $a = Sample 'idle start'; Start-Sleep -Seconds 10; $b = Sample 'idle end'
    Log ('IDLE cpu seconds over 10 s: ' + [math]::Round($b.CpuSeconds - $a.CpuSeconds, 2) + ' handles ' + $b.Handles + ' gdi ' + $b.Gdi + ' user ' + $b.User + ' privateMB ' + $b.PrivateMB)
}
Section 'busy spinner 30 s' {
    
    $rows = @((Get-El $win 'dataGrid').FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem))))
    Log ('rows realized: ' + $rows.Count)
    foreach ($row in ($rows | Select-Object -First 8)) {
        $box = $row.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::CheckBox)))
        if ($box) { $t = $box.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); if ($t.Current.ToggleState -ne 'On') { $t.Toggle() } }
    }
    Invoke-El (Get-El $win 'btnCheckLatest')
    Start-Sleep -Seconds 2
    $overlay = Get-El $win 'txtProgressTitle' -Quiet
    Log ('busy overlay shown: ' + [bool]$overlay)
    $samples = New-Object System.Collections.Generic.List[object]
    $samples.Add((Sample 't0'))
    for ($t = 1; $t -le 30; $t++) { Start-Sleep -Seconds 1; $samples.Add((Sample ('t' + $t))) }
    $first = $samples[0]; $last = $samples[$samples.Count - 1]
    $still = Get-El $win 'txtProgressTitle' -Quiet
    Log ('busy overlay still shown after 30 s: ' + [bool]$still)
    Log ('BUSY cpu seconds over 30 s: ' + [math]::Round($last.CpuSeconds - $first.CpuSeconds, 2))
    Log ('BUSY handles ' + $first.Handles + ' -> ' + $last.Handles + ', gdi ' + $first.Gdi + ' -> ' + $last.Gdi + ', user ' + $first.User + ' -> ' + $last.User + ', privateMB ' + $first.PrivateMB + ' -> ' + $last.PrivateMB + ', workingSetMB ' + $first.WorkingSetMB + ' -> ' + $last.WorkingSetMB)
    $samples | Export-Csv -LiteralPath (Join-Path $OutDir 'busy-samples.txt') -NoTypeInformation
    Shot $win 'busy-spinner'
    try { Invoke-El (Get-El $win 'btnCancelPipeline') } catch { $null = $_ }
    Start-Sleep -Seconds 3
}
Stop-App 'start-apppackager.ps1'
