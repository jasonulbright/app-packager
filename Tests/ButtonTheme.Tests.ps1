BeforeAll {
    $root = Join-Path $PSScriptRoot '..'
    $script:Sources = @(
        Get-Item (Join-Path $root 'start-apppackager.ps1')
        Get-ChildItem (Join-Path $root 'Packagers') -Filter '*.psm1' -File
    )
}

Describe 'Button theme' {
    # A Button with no Style renders with the system chrome instead of the theme.
    It 'gives every XAML Button a Style' {
        $unstyled = foreach ($f in $script:Sources) {
            $text = [IO.File]::ReadAllText($f.FullName)
            foreach ($m in [regex]::Matches($text, '<Button\b[^>]*>', 'Singleline')) {
                if ($m.Value -notmatch '\bStyle=') {
                    $line = ($text.Substring(0, $m.Index) -split "`n").Count
                    '{0}:{1}' -f $f.Name, $line
                }
            }
        }
        @($unstyled) | Should -BeNullOrEmpty
    }

    # The sidebar styles carry fixed dark colors. Set-ButtonTheme recolors them
    # for the light theme, but only for buttons in the registered lists. A button
    # that uses the style and is missing from the list stays black on white.
    It 'registers every sidebar button for the theme change' {
        $xaml = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\MainWindow.xaml'))
        $shell = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\start-apppackager.ps1'))
        $styled = foreach ($m in [regex]::Matches($xaml, '<Button\b[^>]*>', 'Singleline')) {
            if ($m.Value -match 'Style="\{StaticResource (Workflow|Options)ButtonStyle\}"' -and $m.Value -match 'x:Name="(\w+)"') { $Matches[1] }
        }
        $registered = foreach ($m in [regex]::Matches($shell, '(?m)^\s*\$script:(?:WorkflowButtons|OptionsButtons)\s*=\s*@\(([^)]*)\)')) {
            [regex]::Matches($m.Groups[1].Value, '\$(\w+)') | ForEach-Object { $_.Groups[1].Value }
        }
        @($styled).Count | Should -BeGreaterThan 10
        $missing = @($styled | Where-Object { $_ -notin @($registered) })
        $missing | Should -BeNullOrEmpty
    }

    It 'gives every code-built Button a Style' {
        $unstyled = foreach ($f in $script:Sources) {
            $lines = [IO.File]::ReadAllLines($f.FullName)
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^\s*(\$\w+)\s*=\s*New-Object\s+(System\.Windows\.Controls\.)?Button\b') {
                    $var = [regex]::Escape($Matches[1])
                    $block = $lines[$i..([Math]::Min($i + 12, $lines.Count - 1))] -join "`n"
                    if ($block -notmatch ($var + '\.(SetResourceReference\(\[System\.Windows\.FrameworkElement\]::StyleProperty|Style\s*=)')) {
                        '{0}:{1}' -f $f.Name, ($i + 1)
                    }
                }
            }
        }
        @($unstyled) | Should -BeNullOrEmpty
    }
}
