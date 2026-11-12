#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    $root = Join-Path $PSScriptRoot '..'
    $script:Sources = @(
        Get-Item (Join-Path $root 'start-apppackager.ps1')
        Get-ChildItem (Join-Path $root 'Packagers') -Filter '*.psm1' -File
    )
    function Get-RegisteredButtonName([string]$Shell) {
        # The parser tells a comment from a hash sign inside a string.
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($Shell, [ref]$tokens, [ref]$errors)
        $chars = $Shell.ToCharArray()
        foreach ($token in $tokens) {
            # Comments and quoted strings register nothing.
            if ($token.Kind -notin @([System.Management.Automation.Language.TokenKind]::Comment, [System.Management.Automation.Language.TokenKind]::StringLiteral, [System.Management.Automation.Language.TokenKind]::StringExpandable)) { continue }
            for ($i = $token.Extent.StartOffset; $i -lt $token.Extent.EndOffset; $i++) {
                if ($chars[$i] -ne "`n" -and $chars[$i] -ne "`r") { $chars[$i] = ' ' }
            }
        }
        $Shell = -join $chars
        foreach ($m in [regex]::Matches($Shell, '(?m)^\s*\$script:(?:WorkflowButtons|OptionsButtons)\s*=\s*@\(([^)]*)\)')) {
            [regex]::Matches($m.Groups[1].Value, '\$(\w+)') | ForEach-Object { $_.Groups[1].Value }
        }
    }
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
        $registered = Get-RegisteredButtonName $shell

        @($styled).Count | Should -BeGreaterThan 10
        $missing = @($styled | Where-Object { $_ -notin @($registered) })
        $missing | Should -BeNullOrEmpty
    }

    # A name in a comment is not registered at run time.
    It 'ignores button names in comments of the registered lists' {
        $shell = "`$script:WorkflowButtons = @(`$btnA,`n    # `$btnB,`n    <# `$btnC #> `$btnD)"
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnA'
        $names | Should -Contain 'btnD'
        $names | Should -Not -Contain 'btnB'
        $names | Should -Not -Contain 'btnC'
    }

    It 'ignores a comment that follows a comma without a space' {
        $shell = '$script:OptionsButtons = @($btnA,#$btnB
    $btnC)'
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnA'
        $names | Should -Contain 'btnC'
        $names | Should -Not -Contain 'btnB'
    }

    It 'ignores a comment after an opening parenthesis' {
        $shell = '$script:WorkflowButtons = @(#$btnB
    $btnA)'
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnA'
        $names | Should -Not -Contain 'btnB'
    }

    It 'ignores a comment after a semicolon' {
        $shell = '$script:OptionsButtons = @($btnA);#$btnB'
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnA'
        $names | Should -Not -Contain 'btnB'
    }

    # A quoted name is a string, not the button variable.
    It 'ignores a quoted button-like name beside a valid variable' {
        $shell = '$script:WorkflowButtons = @(''$btnQuoted'', $btnA)'
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnA'
        $names | Should -Not -Contain 'btnQuoted'
    }

    # Each registered variable must hold the button of the same name. A name
    # list alone passes when $btnWorkbench points at another control.
    It 'binds every registered button variable to its own XAML button' {
        $shell = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\start-apppackager.ps1'))
        $registered = @(Get-RegisteredButtonName $shell)
        $registered | Should -Contain 'btnWorkbench'
        $registered | Should -Contain 'btnAddInstaller'
        $bad = foreach ($name in $registered) {
            if ($shell -notmatch ('(?m)^\$' + $name + '\s*=\s*\$window\.FindName\(''' + $name + '''\)')) { $name }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'finds a variable bound to another control' {
        $shell = '$btnWorkbench = $window.FindName(''btnOther'')'
        ($shell -match '(?m)^\$btnWorkbench\s*=\s*\$window\.FindName\(''btnWorkbench''\)') | Should -BeFalse
    }

    # A hash sign inside a quoted string starts no comment.
    It 'keeps names that follow a quoted string with a hash sign' {
        $shell = '$script:WorkflowButtons = @($buttonByKey[''Status,#FFF''], $btnWorkbench)'
        $names = @(Get-RegisteredButtonName $shell)
        $names | Should -Contain 'btnWorkbench'
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
