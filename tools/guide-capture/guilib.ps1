# Lab-side UI Automation helpers. Dot-sourced by the runner inside the interactive session.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class NativeUi {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool r);
}
'@
[void][NativeUi]::SetProcessDPIAware()
$AE = [System.Windows.Automation.AutomationElement]
$root = $AE::RootElement
function Log([string]$m) { ('{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $m) | Add-Content -LiteralPath (Join-Path $OutDir 'run.log') }
function Find-Top([string]$Name, [int]$Seconds = 30) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    $end = (Get-Date).AddSeconds($Seconds)
    $winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
    do {
        # A name with a wildcard (the title carries the version) matches with -like; any other name matches exactly.
        if ($Name -match '[*?]') { $el = @($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -like $Name } | Select-Object -First 1 } else { $el = $root.FindFirst('Descendants', $cond) }
        if ($el) { return $el }; Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $end)
    $names = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name + '|' + $_.Current.ClassName }) -join ' ; '
    Log ("top windows: " + $names); Shot-Screen ('fail-' + ($Name -replace '\W', '_')); throw "window not found: $Name"
}
function Get-El($Scope, [string]$Id, [switch]$Quiet) {
    $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, $Id)))
    if (-not $el) { $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Id))) }
    if (-not $el -and -not $Quiet) { throw "control not found: $Id" }
    $el
}
function Invoke-El($El) { [void][NativeUi]::SetCursorPos(1900, 1060); $El.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
function Set-Text($El, [string]$Text) { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Text) }
function Toggle-El($El) { $El.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle() }
function Select-El($El) { $El.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
function Expand-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() }
function Collapse-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse() }
function Get-ElText($El) { try { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { $El.Current.Name } }
function Wait-Until([scriptblock]$Test, [int]$Seconds = 120) { $end = (Get-Date).AddSeconds($Seconds); while ((Get-Date) -lt $end) { if (& $Test) { return $true }; Start-Sleep -Milliseconds 500 }; $false }
function Click-At([int]$x, [int]$y) { [void][NativeUi]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 150; [NativeUi]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 250 }
function Click-El($El) { $r = $El.Current.BoundingRectangle; Click-At ([int]($r.X + $r.Width / 2)) ([int]($r.Y + $r.Height / 2)) }
function Move-To($El) { $r = $El.Current.BoundingRectangle; [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)) }
function Send-Keys([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k); Start-Sleep -Milliseconds 400 }
function Focus-Win($Win) { try { [void][NativeUi]::SetForegroundWindow([IntPtr]$Win.Current.NativeWindowHandle) } catch { $null = $_ }; Start-Sleep -Milliseconds 300 }
function Size-Win($Win, [int]$W = 1440, [int]$H = 900) {
    $hh = [IntPtr]$Win.Current.NativeWindowHandle
    $sw = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $w2 = [Math]::Min($W, $sw.Width); $h2 = [Math]::Min($H, $sw.Height - 40)
    [void][NativeUi]::MoveWindow($hh, 0, 0, $w2, $h2, $true); Start-Sleep -Milliseconds 900
}
function Shot($Win, [string]$Name, [switch]$Hover) {
    if (-not $Hover) { [void][NativeUi]::SetCursorPos(1900, 1060); if ($Win.Current.ClassName -eq 'MetroWindow') { try { (Get-El $Win 'txtLog' -Quiet).SetFocus() } catch { $null = $_ } } }
    Start-Sleep -Milliseconds 1300
    $r = $Win.Current.BoundingRectangle
    $x = [int][Math]::Max(0, $r.X); $y = [int][Math]::Max(0, $r.Y)
    $w = [int][Math]::Min($r.Width, [System.Windows.Forms.SystemInformation]::VirtualScreen.Width - $x)
    $h = [int][Math]::Min($r.Height, [System.Windows.Forms.SystemInformation]::VirtualScreen.Height - $y)
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h)); $g.Dispose()
    $bmp.Save((Join-Path $OutDir ($Name + '.png')), [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Log ('shot {0} {1}x{2}' -f $Name, $w, $h)
}
function Shot-Screen([string]$Name) {
    $s = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $b = New-Object System.Drawing.Bitmap $s.Width, $s.Height; $g = [System.Drawing.Graphics]::FromImage($b); $g.CopyFromScreen(0, 0, 0, 0, $b.Size); $g.Dispose()
    $b.Save((Join-Path $OutDir ($Name + '.png'))); $b.Dispose(); Log ('screen ' + $Name)
}
function Dump-Tree($Scope, [string]$Name) {
    $out = foreach ($e in $Scope.FindAll('Descendants', [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $e.Current; if ($c.Name -or $c.AutomationId) { '{0}|{1}|{2}|{3}|{4}' -f $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.AutomationId, $c.Name, $c.IsEnabled, [math]::Round([double]$c.BoundingRectangle.Width) }
    }
    $out | Set-Content (Join-Path $OutDir ($Name + '.tree.txt'))
}
function Start-App([string]$Script, [string]$Title) {
    [void][NativeUi]::SetCursorPos(1900, 1060)
    $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', ('"' + (Join-Path $appDir $Script) + '"')) -PassThru
    $first = Find-Top $Title 90; Start-Sleep -Seconds 3
    # A modal first-run window can hide the main window from UI Automation for a moment. Look again; keep the first hit when it stays hidden.
    try { Find-Top $Title 30 } catch { $first }
}
function Stop-App([string]$Script) {
    $target = Join-Path $appDir $Script
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$target*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}
function Section([string]$Name, [scriptblock]$Body) { Log ('== ' + $Name); try { & $Body } catch { Log ('ERROR in ' + $Name + ': ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber) } }
function Find-All([string]$Name) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    @($root.FindAll('Descendants', $cond))
}
function Find-Last([string]$Name, [int]$Seconds = 15) {
    $end = (Get-Date).AddSeconds($Seconds); do { $a = Find-All $Name; if ($a.Count) { return $a[$a.Count - 1] }; Start-Sleep -Milliseconds 400 } while ((Get-Date) -lt $end); throw "no window $Name"
}
function Find-Role($Scope, [string]$Type, [string]$Name) {
    $ct = [System.Windows.Automation.ControlType]::($Type)
    $c = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ct)), (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)))
    $Scope.FindFirst('Descendants', $c)
}
function RightClick-El($El) { $r = $El.Current.BoundingRectangle; [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 200; [NativeUi]::mouse_event(8, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(16, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 600 }
function File-Dialog([string]$Path, [string]$Shot) {
    # Wait for the Save As window. Stop the capture when it does not show,
    # so that an export that was never saved cannot count as a success.
    $d = $null; $end = (Get-Date).AddSeconds(20)
    while (-not $d -and (Get-Date) -lt $end) {
        $d = foreach ($n in 'Save As','Open','Save') { $f = @(Find-All $n) | Select-Object -First 1; if ($f) { $f; break } }
        if (-not $d) { $d = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition)) | Where-Object { $_.Current.ClassName -eq '#32770' } | Select-Object -First 1 }
        if (-not $d) { Start-Sleep -Milliseconds 400 }
    }
    if (-not $d) {
        $names = @($root.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window))) | ForEach-Object { $_.Current.Name + '|' + $_.Current.ClassName }) -join ' ; '
        Log ('windows: ' + $names)
        throw 'The Save As dialog did not show.'
    }
    Start-Sleep -Milliseconds 800
    if ($Shot) { Shot $d $Shot }
    Focus-Win $d
    if ($Path) { Send-Keys ($Path -replace '([+^%~{}\[\]()])', '{$1}') }
    Send-Keys '{ENTER}'; Start-Sleep -Seconds 2
}
