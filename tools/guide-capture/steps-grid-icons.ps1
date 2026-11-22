# Window proof of the grid icons. Runs in the interactive session of a lab account.
# It reads nothing from the site and publishes nothing. It checks that the main grid shows the icon pack icons
# in both themes, that a packager without an icon shows the placeholder, and that an icon pack update while the
# grid is open replaces the files (no handle on them) and the grid takes the new icons.
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-CaptureSiteData
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
$packDir = Join-Path $appDir 'Packagers\Icons'
Add-Type -AssemblyName System.Drawing
# One pack icon is replaced by a red square before the window starts. The update later puts the pack icon back.
$target = Join-Path $packDir 'powershell7msix.png'
$bmp = New-Object System.Drawing.Bitmap 128, 128
$g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::FromArgb(255, 220, 20, 20)); $g.Dispose()
$bmp.Save($target, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
$packCount = @(Get-ChildItem -LiteralPath $packDir -Filter *.png -File).Count
Log ('pack png files: ' + $packCount)
$started = Get-Date
$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
Size-Win $win 1440 900
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
$winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
function Image-Count { $grid = Get-El $win 'dataGrid'; @($grid.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Image)))).Count }
function Row-Count { $grid = Get-El $win 'dataGrid'; @($grid.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))).Count }
Section 'setup' {
    $s = Find-Last 'Setup' 20
    Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
    Set-Text (Get-El $s 'cm.fileShare') '\cm01.contoso.com\ContentShare$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
    Log ('grid ready after seconds: ' + [int]((Get-Date) - $started).TotalSeconds)
}
Section 'dark theme' {
    $null = Wait-Until { (Row-Count) -gt 5 } 30
    Log ('rows realized: ' + (Row-Count) + ', Image elements: ' + (Image-Count))
    Shot $win 'grid-icons-dark-before-update'
    Set-Text (Get-El $win 'txtGridFilter') 'PowerShell 7'; Start-Sleep -Seconds 1
    Log ('filtered rows: ' + (Row-Count) + ', Image elements: ' + (Image-Count))
    Shot $win 'grid-icons-powershell7-red'
    Set-Text (Get-El $win 'txtGridFilter') ''; Start-Sleep -Seconds 1
}
Section 'pack update while the grid is open' {
    Invoke-El (Get-El $win 'btnOptions'); $o = Find-LastLike 'Options*' 15
    $tools = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem)))) | Where-Object { $_.Current.Name -eq 'Tools' } | Select-Object -First 1
    if (-not $tools) { throw 'The Options tree has no Tools page.' }
    Select-El $tools; Start-Sleep -Milliseconds 900
    $before = (Get-Item -LiteralPath $target).Length
    Invoke-El (Get-El $o 'tools.iconpack.download')
    # The line of the icon pack shows the installed pack when the download ends.
    $null = Wait-Until { $b = Get-El $o 'tools.iconpack.download' -Quiet; $b -and $b.Current.IsEnabled } 120
    Start-Sleep -Seconds 2
    $after = (Get-Item -LiteralPath $target).Length
    Log ('powershell7msix.png bytes: before {0}, after {1}' -f $before, $after)
    if ($after -eq $before) { throw 'The icon pack update did not replace the icon file.' }
    Shot $o 'grid-icons-options-after-update'
    Invoke-El (Get-El $o 'Cancel'); Start-Sleep -Seconds 2
    Set-Text (Get-El $win 'txtGridFilter') 'PowerShell 7'; Start-Sleep -Seconds 1
    Shot $win 'grid-icons-powershell7-after-update'
    Set-Text (Get-El $win 'txtGridFilter') ''; Start-Sleep -Seconds 1
    Log ('after update rows: ' + (Row-Count) + ', Image elements: ' + (Image-Count))
    Shot $win 'grid-icons-dark-after-update'
}
Section 'light theme' {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2
    Log ('light rows: ' + (Row-Count) + ', Image elements: ' + (Image-Count))
    Shot $win 'grid-icons-light'
    Set-Text (Get-El $win 'txtGridFilter') 'PowerShell 7'; Start-Sleep -Seconds 1
    Shot $win 'grid-icons-light-powershell7'
    Set-Text (Get-El $win 'txtGridFilter') ''; Start-Sleep -Seconds 1
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
}
Get-ChildItem -LiteralPath (Join-Path $appDir 'Logs') -Filter *.log -File -ErrorAction SilentlyContinue | Copy-Item -Destination $OutDir -Force
Stop-App 'start-apppackager.ps1'
