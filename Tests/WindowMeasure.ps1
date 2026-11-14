# Measurement functions for the window probes: contrast of text and of control borders. They return plain data.
# Copied from the other suite tools. The probe dot-sources this file, then the script block MeasureFunctions.

$script:MeasureFunctions = {
    function ConvertTo-Linear([double]$v) {
        if ($v -le 0.03928) { return $v / 12.92 }
        return [math]::Pow(($v + 0.055) / 1.055, 2.4)
    }

    function Get-RelativeLuminance($c) {
        return 0.2126 * (ConvertTo-Linear ($c.R / 255.0)) + 0.7152 * (ConvertTo-Linear ($c.G / 255.0)) + 0.0722 * (ConvertTo-Linear ($c.B / 255.0))
    }

    function Get-ContrastRatio($a, $b) {
        $la = Get-RelativeLuminance $a
        $lb = Get-RelativeLuminance $b
        $hi = [math]::Max($la, $lb); $lo = [math]::Min($la, $lb)
        return ($hi + 0.05) / ($lo + 0.05)
    }

    # Blends a colour with alpha over an opaque colour.
    function Merge-Color($top, [double]$opacity, $under) {
        $a = ($top.A / 255.0) * $opacity
        return [System.Windows.Media.Color]::FromRgb(
            [byte][math]::Round($top.R * $a + $under.R * (1 - $a)),
            [byte][math]::Round($top.G * $a + $under.G * (1 - $a)),
            [byte][math]::Round($top.B * $a + $under.B * (1 - $a)))
    }

    function Get-SolidColor($brush) {
        if ($brush -is [System.Windows.Media.SolidColorBrush]) { return @{ Color = $brush.Color; Opacity = $brush.Opacity } }
        return $null
    }

    # The opaque colour behind an element: walks the visual parents and blends
    # every translucent background it meets over the next one.
    function Get-BackdropColor($element, $windowBackground, $window = $null) {
        $layers = New-Object System.Collections.Generic.List[object]
        $d = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
        while ($d) {
            if ($window -and $d -is [System.Windows.FrameworkElement] -and $d.Name -eq 'PART_TitleBar') {
                # The title bar is painted from the window's title brush, not from a Background.
                # Always the active title bar, so the result does not depend on whether the runner
                # activated the window. Get-TitleBarContrast measures the inactive bar.
                $tbBrush = $window.WindowTitleBrush
                $tsc = Get-SolidColor $tbBrush
                if ($tsc) { $layers.Add(@{ C = $tsc.Color; O = $tsc.Opacity }); break }
            }
            $bg = $null
            if ($d -is [System.Windows.Controls.Border]) { $bg = $d.Background }
            elseif ($d -is [System.Windows.Controls.Panel]) { $bg = $d.Background }
            elseif ($d -is [System.Windows.Controls.Control]) { $bg = $d.Background }
            $sc = Get-SolidColor $bg
            if ($sc -and ($sc.Color.A * $sc.Opacity) -gt 0) {
                $layers.Add(@{ C = $sc.Color; O = $sc.Opacity })
                if ($sc.Color.A -eq 255 -and $sc.Opacity -ge 1) { break }
            }
            $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
        }
        $acc = $windowBackground
        for ($i = $layers.Count - 1; $i -ge 0; $i--) { $acc = Merge-Color $layers[$i].C $layers[$i].O $acc }
        return $acc
    }

    function Get-ElementOpacity($element) {
        $o = 1.0
        $d = $element
        while ($d) {
            if ($d -is [System.Windows.UIElement]) { $o *= $d.Opacity }
            $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
        }
        return $o
    }

    function Get-TextContrast {
        # One row per visible piece of text: its colour, backdrop and ratio.
        param($Window, [string]$Tag = '', $Root = $Window)
        $winBg = (Get-SolidColor $Window.Background)
        $under = if ($winBg) { Merge-Color $winBg.Color $winBg.Opacity ([System.Windows.Media.Colors]::Black) } else { [System.Windows.Media.Colors]::White }
        $rows = @()
        foreach ($e in (Get-VisualDescendants $Root)) {
            $fg = $null; $text = $null
            if ($e -is [System.Windows.Controls.TextBlock]) { $text = $e.Text; $fg = $e.Foreground }
            elseif ($e -is [System.Windows.Controls.TextBox]) { $text = $e.Text; $fg = $e.Foreground }
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            if (-not (Test-VisibleInTree $e)) { continue }
            if ($e.ActualWidth -le 0 -or $e.ActualHeight -le 0) { continue }
            $sc = Get-SolidColor $fg
            if (-not $sc) { continue }
            $backdrop = if ($e -is [System.Windows.Controls.TextBox] -and (Get-SolidColor $e.Background) -and (Get-SolidColor $e.Background).Color.A -gt 0) {
                Merge-Color (Get-SolidColor $e.Background).Color (Get-SolidColor $e.Background).Opacity (Get-BackdropColor $e $under $Window)
            } else { Get-BackdropColor $e $under $Window }
            $opacity = (Get-ElementOpacity $e) * $sc.Opacity
            $fgc = Merge-Color $sc.Color $opacity $backdrop
            $name = $e.Name
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($e)
            while (-not $name -and $p) { if ($p -is [System.Windows.FrameworkElement]) { $name = $p.Name }; $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p) }
            $rows += [PSCustomObject]@{
                Tag = $Tag; Name = $name; Text = $text.Substring(0, [math]::Min(40, $text.Length))
                Size = $e.FontSize; Bold = ($e.FontWeight.ToOpenTypeWeight() -ge 700); Icon = (@("Marlett", "Webdings", "Segoe MDL2 Assets", "Segoe Fluent Icons") -contains $e.FontFamily.Source)
                Foreground = $fgc.ToString(); Backdrop = $backdrop.ToString()
                Ratio = [math]::Round((Get-ContrastRatio $fgc $backdrop), 2)
            }
        }
        return $rows
    }

    function Get-TitleBarContrast {
        # Contrast of the title text on the active and on the inactive title bar.
        param($Window)
        $fg = Get-SolidColor $Window.TitleForeground
        $result = [ordered]@{ Active = $null; Inactive = $null }
        if (-not $fg) { return [PSCustomObject]$result }
        foreach ($state in 'Active', 'Inactive') {
            $brush = if ($state -eq 'Active') { $Window.WindowTitleBrush } else { $Window.NonActiveWindowTitleBrush }
            $bg = Get-SolidColor $brush
            if (-not $bg) { continue }
            $back = Merge-Color $bg.Color $bg.Opacity ([System.Windows.Media.Colors]::Black)
            $front = Merge-Color $fg.Color $fg.Opacity $back
            $result[$state] = [math]::Round((Get-ContrastRatio $front $back), 2)
        }
        return [PSCustomObject]$result
    }

    function Get-ClippedText {
        # Text that does not fit its TextBlock: unwrapped, untrimmed and narrower than the text.
        param($Window)
        $rows = @()
        foreach ($e in (Get-VisualDescendants $Window)) {
            if ($e -isnot [System.Windows.Controls.TextBlock]) { continue }
            if ([string]::IsNullOrEmpty($e.Text) -or -not (Test-VisibleInTree $e) -or $e.ActualWidth -le 0) { continue }
            if ($e.TextWrapping -ne 'NoWrap' -or $e.TextTrimming -ne 'None') { continue }
            $tf = New-Object System.Windows.Media.Typeface($e.FontFamily, $e.FontStyle, $e.FontWeight, $e.FontStretch)
            $ft = New-Object System.Windows.Media.FormattedText($e.Text, [System.Globalization.CultureInfo]::CurrentCulture,
                [System.Windows.FlowDirection]::LeftToRight, $tf, $e.FontSize, [System.Windows.Media.Brushes]::Black, [System.Windows.Media.VisualTreeHelper]::GetDpi($e).PixelsPerDip)
            if ($e.ActualWidth + 1 -lt $ft.WidthIncludingTrailingWhitespace) {
                $rows += [PSCustomObject]@{ Name = $e.Name; Text = $e.Text; Needed = [math]::Round($ft.WidthIncludingTrailingWhitespace, 1); Actual = [math]::Round($e.ActualWidth, 1) }
            }
        }
        return $rows
    }

    function Get-OffscreenControls {
        # Visible named controls whose box is not inside the window content.
        param($Window)
        $rows = @()
        $content = $Window
        foreach ($e in (Get-VisualDescendants $Window)) {
            if ($e -isnot [System.Windows.FrameworkElement] -or -not $e.Name) { continue }
            if (-not (Test-VisibleInTree $e) -or $e.ActualWidth -le 0) { continue }
            if ($e -isnot [System.Windows.Controls.Primitives.ButtonBase] -and $e -isnot [System.Windows.Controls.TextBox] -and $e -isnot [System.Windows.Controls.Primitives.Selector] -and $e -isnot [System.Windows.Controls.TreeView]) { continue }
            $p = $e.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($p.X -lt -1 -or $p.Y -lt -1 -or ($p.X + $e.ActualWidth) -gt ($content.ActualWidth + 1) -or ($p.Y + $e.ActualHeight) -gt ($content.ActualHeight + 1)) {
                $rows += [PSCustomObject]@{ Name = $e.Name; X = [math]::Round($p.X); Y = [math]::Round($p.Y); W = [math]::Round($e.ActualWidth); H = [math]::Round($e.ActualHeight); ContentW = [math]::Round($content.ActualWidth); ContentH = [math]::Round($content.ActualHeight) }
            }
        }
        return $rows
    }

    function Get-ControlColors {
        # Text, border and backdrop colours of one control (its own text and its template text).
        param($Control, $Window)
        $winBg = (Get-SolidColor $Window.Background)
        $under = if ($winBg) { Merge-Color $winBg.Color $winBg.Opacity ([System.Windows.Media.Colors]::Black) } else { [System.Windows.Media.Colors]::White }
        $own = Get-SolidColor $Control.Background
        $backdrop = Get-BackdropColor $Control $under $Window
        if ($own -and ($own.Color.A * $own.Opacity) -gt 0) { $backdrop = Merge-Color $own.Color $own.Opacity $backdrop }
        $o = Get-ElementOpacity $Control
        $row = [ordered]@{ Name = $Control.Name; Type = $Control.GetType().Name; Backdrop = $backdrop.ToString(); TextRatio = $null; BorderRatio = $null }
        $worst = $null
        foreach ($t in (Get-VisualDescendants $Control)) {
            if ($t -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrWhiteSpace($t.Text) -or $t.ActualWidth -le 0 -or -not (Test-VisibleInTree $t) -or (@("Marlett", "Webdings", "Segoe MDL2 Assets", "Segoe Fluent Icons") -contains $t.FontFamily.Source)) { continue }
            $sc = Get-SolidColor $t.Foreground
            if (-not $sc) { continue }
            $fgc = Merge-Color $sc.Color ((Get-ElementOpacity $t) * $sc.Opacity) $backdrop
            $r = Get-ContrastRatio $fgc $backdrop
            if ($null -eq $worst -or $r -lt $worst) { $worst = $r }
        }
        if ($Control -is [System.Windows.Controls.TextBox] -and -not [string]::IsNullOrWhiteSpace($Control.Text)) {
            $sc = Get-SolidColor $Control.Foreground
            if ($sc) { $r = Get-ContrastRatio (Merge-Color $sc.Color ($o * $sc.Opacity) $backdrop) $backdrop; if ($null -eq $worst -or $r -lt $worst) { $worst = $r } }
        }
        if ($null -ne $worst) { $row.TextRatio = [math]::Round($worst, 2) }
        $bsc = Get-SolidColor $Control.BorderBrush
        if ($bsc -and $Control.BorderThickness.Left -gt 0) {
            $outside = Get-BackdropColor $Control $under $Window
            $row.BorderRatio = [math]::Round((Get-ContrastRatio (Merge-Color $bsc.Color ($o * $bsc.Opacity) $outside) $outside), 2)
        }
        return [PSCustomObject]$row
    }

    function Get-InteractiveContrast {
        # Colours of every visible named button, text box, combo box and check box,
        # enabled and then disabled.
        param($Window)
        $rows = @()
        $controls = @(Get-VisualDescendants $Window | Where-Object {
            ($_ -is [System.Windows.Controls.Primitives.ButtonBase] -or $_ -is [System.Windows.Controls.TextBox] -or $_ -is [System.Windows.Controls.ComboBox]) -and
            $_.Name -and $_.Name -notlike 'PART_*' -and (Test-VisibleInTree $_) -and $_.ActualWidth -gt 0 })
        foreach ($c in $controls) {
            # Remember where IsEnabled came from (a local value, or the parent and the style), and put it back also when a measurement throws.
            $local = $c.ReadLocalValue([System.Windows.UIElement]::IsEnabledProperty)
            try {
                $c.IsEnabled = $true; Wait-Layout $Window 2
                $en = Get-ControlColors $c $Window
                $c.IsEnabled = $false; Wait-Layout $Window 2
                $dis = Get-ControlColors $c $Window
            }
            finally {
                if ($local -eq [System.Windows.DependencyProperty]::UnsetValue) { $c.ClearValue([System.Windows.UIElement]::IsEnabledProperty) } else { $c.IsEnabled = [bool]$local }
                Wait-Layout $Window 1
            }
            $rows += [PSCustomObject]@{ Name = $c.Name; Type = $en.Type; EnabledText = $en.TextRatio; DisabledText = $dis.TextRatio; EnabledBorder = $en.BorderRatio; DisabledBorder = $dis.BorderRatio }
        }
        return $rows
    }
}