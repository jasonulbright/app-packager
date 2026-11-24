<#
.SYNOPSIS
    Window proof that every row of the Packagers grid keeps its own icon after a sort, a filter and a refresh.

.DESCRIPTION
    Runs in the interactive session of a lab account (start it with tools\Run-VisibleProbe.ps1 of the suite repository, under the capture
    lock). The run installs the icon pack, starts the app, and for the realized rows of the grid reads the row's Application text and the
    screen pixels of the row's icon. The expected icon is the pack file named for the row's packager script (the identity of the row).
    A row passes when the pixels of its icon are closer to its own pack file than to the pack file of any other packager (a small
    tolerance covers packs that hold near-identical icons). The run clicks every sortable column header, ascending and then descending,
    and verifies after each click, after a page of scrolling, after a filter and after a theme switch (the grid refreshes). It runs in the
    Dark theme and then in the Light theme. It reads no site and publishes nothing. Each check prints ok or FAIL; the run ends with
    "ICON-FAILS=<n>".
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
$titlePattern = 'AppPackager*v20??.??.??.????'
$script:Fails = 0
$script:Checks = 0
function Say([string]$m) { Write-Output $m; Log $m }
function Test-Icon([bool]$Condition, [string]$Message) {
    $script:Checks++
    if ($Condition) { Say ('ok: ' + $Message) } else { $script:Fails++; Say ('FAIL: ' + $Message) }
}
function Items($Scope, $Type) { @($Scope.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::($Type))))) }

# 1. The icon pack, from the last release, into this copy of the app.
& (Join-Path $PSScriptRoot 'install-iconpack.ps1') -AppRoot $appDir | Out-Null
$packDir = Join-Path $appDir 'Packagers\Icons'

# The signature of an image: a 12 x 12 grid of RGB values, the image composited on a background color.
$grid = 12
function Get-Signature([System.Drawing.Bitmap]$Source, [System.Drawing.Color]$Background) {
    $flat = New-Object System.Drawing.Bitmap $grid, $grid
    $g = [System.Drawing.Graphics]::FromImage($flat)
    $g.Clear($Background)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($Source, 0, 0, $grid, $grid)
    $g.Dispose()
    $sig = New-Object 'int[]' ($grid * $grid * 3)
    $i = 0
    for ($y = 0; $y -lt $grid; $y++) { for ($x = 0; $x -lt $grid; $x++) { $p = $flat.GetPixel($x, $y); $sig[$i++] = $p.R; $sig[$i++] = $p.G; $sig[$i++] = $p.B } }
    $flat.Dispose()
    return , $sig
}
function Get-Distance([int[]]$A, [int[]]$B) { $sum = 0; for ($i = 0; $i -lt $A.Length; $i++) { $sum += [Math]::Abs($A[$i] - $B[$i]) }; return $sum / $A.Length }

$sigCache = @{}
# The icon pack files by packager name, decoded once (alpha kept).
$packBitmaps = @{}
foreach ($f in Get-ChildItem -LiteralPath $packDir -File | Where-Object { $_.Extension -in '.png' }) {
    try { $bytes = [IO.File]::ReadAllBytes($f.FullName); $ms = New-Object IO.MemoryStream (, $bytes); $packBitmaps[$f.BaseName.ToLowerInvariant()] = New-Object System.Drawing.Bitmap $ms } catch { $null = $_ }
}
Say ('icon pack png files: ' + $packBitmaps.Count)

# Application name -> packager base name, from the header of each packager. A name that two packagers share is left out.
$byApp = @{}
$dupApp = @{}
foreach ($f in Get-ChildItem -LiteralPath (Join-Path $appDir 'Packagers') -Filter 'package-*.ps1' -File) {
    $app = $null
    foreach ($line in [IO.File]::ReadLines($f.FullName) | Select-Object -First 60) {
        if ($line -match '^\s*(?:#\s*)?(?:App|Application)\s*:\s*(.+?)\s*$') { $app = $Matches[1].Trim(); break }
    }
    if (-not $app) { continue }
    $base = ($f.BaseName -replace '^package-', '').ToLowerInvariant()
    if ($byApp.ContainsKey($app)) { $dupApp[$app] = $true } else { $byApp[$app] = $base }
}
foreach ($k in $dupApp.Keys) { $byApp.Remove($k) }
Say ('applications with a unique name: ' + $byApp.Count)

# 2. Start the app and finish the setup wizard.
$win = Start-App 'start-apppackager.ps1' $titlePattern
Size-Win $win 1440 1040
$win = Find-Top $titlePattern 10
$s = Find-Last 'Setup' 20
Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
$win = Find-Top $titlePattern 10
Focus-Win $win
function Get-Grid { Get-El $win 'dataGrid' }
$null = Wait-Until { (Items (Get-Grid) 'DataItem').Count -gt 5 } 60
Start-Sleep -Seconds 2

# The verification of the realized rows: the icon of each row against the pack file of the row's packager.
function Test-GridIcons([string]$Label) {
    [void][NativeUi]::SetCursorPos(1900, 1060)
    Start-Sleep -Milliseconds 1200
    $rows = Items (Get-Grid) 'DataItem'
    $checked = 0; $bad = New-Object System.Collections.Generic.List[string]
    $sheet = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $r = $row.Current.BoundingRectangle
        if ($r.Width -le 0 -or $r.Height -le 0) { continue }
        $texts = @(Items $row 'Text' | ForEach-Object { $_.Current.Name } | Where-Object { $byApp.ContainsKey($_) } | Select-Object -Unique)
        if ($texts.Count -ne 1) { continue }
        $app = $texts[0]; $expectedBase = $byApp[$app]
        if (-not $packBitmaps.ContainsKey($expectedBase)) { continue }
        $img = Items $row 'Image' | Select-Object -First 1
        if (-not $img) { $bad.Add("$app has no icon element"); continue }
        $ir = $img.Current.BoundingRectangle
        if ($ir.Width -lt 8 -or $ir.Height -lt 8) { continue }
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $gb = (Get-Grid).Current.BoundingRectangle
        if ($ir.Y -lt ($gb.Y + 34) -or ($ir.Y + $ir.Height) -gt ($gb.Y + $gb.Height - 20) -or $ir.X -lt $gb.X) { continue }
        $w = [int]$ir.Width; $h = [int]$ir.Height
        $shot = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($shot); $g.CopyFromScreen([int]$ir.X, [int]$ir.Y, 0, 0, (New-Object System.Drawing.Size $w, $h)); $g.Dispose()
        # The row background: a pixel of the same row, left of the icon.
        $bgShot = New-Object System.Drawing.Bitmap 1, 1
        $g = [System.Drawing.Graphics]::FromImage($bgShot); $g.CopyFromScreen([int]($ir.X - 4), [int]($ir.Y + $ir.Height / 2), 0, 0, (New-Object System.Drawing.Size 1, 1)); $g.Dispose()
        $bg = $bgShot.GetPixel(0, 0); $bgShot.Dispose()
        $seen = Get-Signature $shot $bg
        $sheet.Add(@($app, $shot.Clone(), $expectedBase, $bg))
        $shot.Dispose()
        $best = [double]::MaxValue; $own = [double]::MaxValue; $bestName = ''
        foreach ($name in $packBitmaps.Keys) {
            $ck = '{0}|{1}' -f $name, $bg.ToArgb()
            if (-not $sigCache.ContainsKey($ck)) { $sigCache[$ck] = Get-Signature $packBitmaps[$name] $bg }
            $d = Get-Distance $seen $sigCache[$ck]
            if ($name -eq $expectedBase) { $own = $d }
            if ($d -lt $best) { $best = $d; $bestName = $name }
        }
        $checked++
        if (-not ($own -le $best + 3 -and $own -lt 70)) { $bad.Add(('{0} shows {1} (own distance {2:N1}, best {3:N1})' -f $app, $bestName, $own, $best)) }
    }
    if ($env:ICON_DIAG) {
        $dir = $env:ICON_DIAG; New-Item -ItemType Directory $dir -Force | Out-Null
        $sh = New-Object System.Drawing.Bitmap 700, ([Math]::Max(1, $sheet.Count) * 36)
        $g = [System.Drawing.Graphics]::FromImage($sh); $g.Clear([System.Drawing.Color]::Gray); $y = 0
        foreach ($e in $sheet) {
            $g.DrawImage($e[1], 2, $y + 2, 32, 32)
            $g.FillRectangle((New-Object System.Drawing.SolidBrush $e[3]), 40, $y + 2, 32, 32)
            $g.DrawImage($packBitmaps[$e[2]], 40, $y + 2, 32, 32)
            $g.DrawString(('{0} -> {1}' -f $e[0], $e[2]), (New-Object System.Drawing.Font 'Segoe UI', 9), [System.Drawing.Brushes]::White, 80, $y + 10); $y += 36
        }
        $g.Dispose(); $sh.Save((Join-Path $dir (($Label -replace '[^A-Za-z0-9]+', '-').Trim('-') + '.png'))); $sh.Dispose()
        $hl = @(Items (Get-Grid) 'HeaderItem' | ForEach-Object { '{0}|{1}|{2}' -f $_.Current.Name, $_.Current.BoundingRectangle, $_.Current.ClassName })
        Add-Content -LiteralPath (Join-Path $dir 'headers.txt') -Value (($Label) + ' :: ' + ($hl -join ' ; '))
    }
    Test-Icon ($checked -ge 8) ("$Label : $checked rows with a pack icon were checked (at least 8 needed)")
    Test-Icon ($bad.Count -eq 0) ("$Label : every checked row shows its own icon" + $(if ($bad.Count) { ' -- WRONG: ' + (($bad | Select-Object -First 6) -join '; ') } else { '' }))
    if ($bad.Count) { Shot $win ('grid-icon-fail-' + (($Label -replace '[^A-Za-z0-9]+', '-').Trim('-'))) }
}

function Wheel-Grid([int]$Notches) {
    $r = (Get-Grid).Current.BoundingRectangle
    [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
    for ($i = 0; $i -lt [Math]::Abs($Notches); $i++) { [NativeUi]::mouse_event(0x800, 0, 0, [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]([Math]::Sign($Notches) * -120)), 0), [UIntPtr]::Zero); Start-Sleep -Milliseconds 40 }
    Start-Sleep -Milliseconds 600
}
function Invoke-SortRun([string]$Theme) {
    Test-GridIcons "$Theme, before sorting"
    $headers = @(Items (Get-Grid) 'HeaderItem' | Where-Object { $_.Current.Name -and $_.Current.BoundingRectangle.Width -gt 0 })
    $names = @($headers | ForEach-Object { $_.Current.Name } | Select-Object -Unique)
    if ($env:ICON_QUICK) { $names = @($names | Where-Object { $_ -in 'Application', 'Status', 'Vendor' }) }
    Say ("$Theme columns: " + ($names -join ', '))
    foreach ($name in $names) {
        foreach ($dir in 'ascending', 'descending') {
            $hdr = @(Items (Get-Grid) 'HeaderItem' | Where-Object { $_.Current.Name -eq $name })[0]
            if (-not $hdr -or $hdr.Current.BoundingRectangle.Width -le 0) { continue }
            Click-El $hdr
            Start-Sleep -Milliseconds 800
            Test-GridIcons "$Theme, sort by $name $dir"
            # The page below the first rows: scroll the grid and verify the rows that come into view.
            $r = (Get-Grid).Current.BoundingRectangle
            [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
            Send-Keys '{PGDN}'
            Test-GridIcons "$Theme, sort by $name $dir, scrolled one page"
            Send-Keys '{PGDN}{PGDN}{PGDN}{PGDN}'
            Test-GridIcons "$Theme, sort by $name $dir, scrolled five pages"
            Send-Keys '^{END}'
            Test-GridIcons "$Theme, sort by $name $dir, scrolled to the end"
            Send-Keys '{PGUP}{PGUP}'
            Test-GridIcons "$Theme, sort by $name $dir, scrolled back up"
            Send-Keys '^{HOME}'
        }
    }
    # The owner's path: scroll with the wheel to the middle of the list, sort there, scroll on, sort again.
    foreach ($name in @('Application', 'Status', 'Vendor')) {
        foreach ($dir in 'ascending', 'descending') {
            Wheel-Grid 25
            $hdr = @(Items (Get-Grid) 'HeaderItem' | Where-Object { $_.Current.Name -eq $name })[0]
            if (-not $hdr) { continue }
            Click-El $hdr; Start-Sleep -Milliseconds 800
            Test-GridIcons "$Theme, wheel to the middle, sort by $name $dir"
            Wheel-Grid 20
            Test-GridIcons "$Theme, wheel on, sort by $name $dir"
            Wheel-Grid -35
            Test-GridIcons "$Theme, wheel back up, sort by $name $dir"
        }
    }
    # Filter, then clear: the ItemsSource changes twice.
    Set-Text (Get-El $win 'txtGridFilter') 'micro'; Start-Sleep -Seconds 1
    Test-GridIcons "$Theme, filter 'micro'"
    Set-Text (Get-El $win 'txtGridFilter') ''; Start-Sleep -Seconds 1
    Test-GridIcons "$Theme, filter cleared"
}

function Run-Part([string]$Name, [scriptblock]$Body) { Say ('== ' + $Name); try { & $Body } catch { $script:Fails++; Say ('FAIL: ' + $Name + ' stopped: ' + $_.Exception.Message) } }
Run-Part 'dark theme' { Invoke-SortRun 'Dark'; Shot $win 'grid-icons-sorted-dark' }
Run-Part 'light theme' {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2
    Invoke-SortRun 'Light'; Shot $win 'grid-icons-sorted-light'
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
}
Get-ChildItem -LiteralPath (Join-Path $appDir 'Logs') -Filter *.log -File -ErrorAction SilentlyContinue | Copy-Item -Destination $OutDir -Force
Stop-App 'start-apppackager.ps1'
Say ('checks: ' + $script:Checks)
Write-Output ('ICON-FAILS=' + $script:Fails)
