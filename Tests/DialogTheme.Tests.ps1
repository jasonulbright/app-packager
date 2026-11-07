BeforeAll {
    $root = Join-Path $PSScriptRoot '..'
    $script:Shell = [IO.File]::ReadAllText((Join-Path $root 'start-apppackager.ps1'))
    $script:WorkbenchXaml = [IO.File]::ReadAllText((Join-Path $root 'WorkbenchWindow.xaml'))

    # Every inline XAML block of the shell, with its line number.
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    $script:Blocks = @($ast.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
         $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and
        $n.Value -match 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    }, $true) | ForEach-Object { [pscustomobject]@{ Line = $_.Extent.StartLineNumber; Text = $_.Value } })
    $script:Windows = @($script:Blocks | Where-Object { $_.Text -match '^\s*<Controls:MetroWindow' })
    $script:ChromeFunction = $ast.Find({
        param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-DialogChromeFromOwner'
    }, $true)

    function Get-StyleBlock([string]$Xaml, [string]$Key) {
        $m = [regex]::Match($Xaml, '<Style\s+x:Key="' + $Key + '"[^>]*>.*?</Style>', 'Singleline')
        if ($m.Success) { $m.Value } else { '' }
    }
}

Describe 'List item style' {
    # An ItemContainerStyle without BasedOn drops the MahApps item template, so
    # the selected item uses system colors: light grey with light text in Dark.
    It 'bases every ListBoxItem style on the MahApps style' {
        $bad = @()
        foreach ($m in [regex]::Matches($script:WorkbenchXaml, '<Style\s+TargetType="ListBoxItem"[^>]*>')) {
            if ($m.Value -notmatch 'BasedOn="\{StaticResource MahApps\.Styles\.ListBoxItem\}"') { $bad += 'WorkbenchWindow.xaml' }
        }
        foreach ($b in $script:Blocks) {
            foreach ($m in [regex]::Matches($b.Text, '<Style\s+TargetType="ListBoxItem"[^>]*>')) {
                if ($m.Value -notmatch 'BasedOn="\{StaticResource MahApps\.Styles\.ListBoxItem\}"') { $bad += ('start-apppackager.ps1:' + $b.Line) }
            }
        }
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'Disabled field fill' {
    # MahApps paints a disabled field with Control.Disabled, translucent white in
    # both themes. On Dark it is a flat light-grey block.
    It 'overrides the disabled fill and dims the text in <Key>' -ForEach @(
        @{ Key = 'WbField' }, @{ Key = 'WbCombo' }, @{ Key = 'WbCode' }
    ) {
        $style = Get-StyleBlock $script:WorkbenchXaml $Key
        $style | Should -Not -BeNullOrEmpty
        $style | Should -Match 'x:Key="MahApps\.Brushes\.Control\.Disabled"'
        $style | Should -Match 'Property="IsEnabled" Value="False"'
        # The override takes its color from the active theme, not a fixed value.
        $style | Should -Match '\{DynamicResource MahApps\.Colors\.'
    }
}

Describe 'Action row layout' {
    # A horizontal StackPanel in an Auto column never wraps, so the last button
    # is cut off at the default width.
    It 'wraps the Workbench action buttons' {
        $x = [xml]$script:WorkbenchXaml
        $ns = New-Object System.Xml.XmlNamespaceManager $x.NameTable
        $ns.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        foreach ($name in 'btnUseDefaults', 'btnSave', 'btnPublishMecm', 'btnPublishIntune', 'btnPublishWsus') {
            $btn = $x.SelectSingleNode('//p:Button[@x:Name="' + $name + '"]', $ns)
            $btn | Should -Not -BeNullOrEmpty
            $btn.ParentNode.LocalName | Should -Be 'WrapPanel'
        }
        $panel = $x.SelectSingleNode('//p:Button[@x:Name="btnSave"]', $ns).ParentNode
        $grid = $panel.ParentNode
        $col = [int]$panel.GetAttribute('Grid.Column')
        $defs = @($grid.SelectNodes('p:Grid.ColumnDefinitions/p:ColumnDefinition', $ns))
        $defs[$col].GetAttribute('Width') | Should -Not -Be 'Auto'
    }
}

Describe 'Dialog button outline' {
    # The MahApps Square style draws a 2 px outline in the foreground color.
    # Workbench and dialogs use ApButton and ApButtonAccent: 1 px, light grey,
    # accent outline on keyboard focus.
    It 'does not use the Square button style directly in the Workbench or in dialogs' {
        $direct = '(Style="\{(Dynamic|Static)Resource MahApps\.Styles\.Button\.Square(\.Accent)?\}"|StyleProperty,\s*''MahApps\.Styles\.Button\.Square'')'
        $bad = @()
        if ($script:WorkbenchXaml -match $direct) { $bad += 'WorkbenchWindow.xaml' }
        foreach ($b in $script:Blocks) { if ($b.Text -match $direct) { $bad += ('start-apppackager.ps1:' + $b.Line) } }
        if ($script:Shell -match 'StyleProperty,\s*''MahApps\.Styles\.Button\.Square''') { $bad += 'start-apppackager.ps1 code-built button' }
        $bad | Should -BeNullOrEmpty
    }

    It 'defines the button styles with a 1 px border and a focus trigger' {
        foreach ($key in 'ApButton', 'ApButtonAccent') {
            $style = Get-StyleBlock $script:WorkbenchXaml $key
            $style | Should -Not -BeNullOrEmpty
            $style | Should -Match 'Property="BorderThickness" Value="1"'
            $style | Should -Match 'Property="IsKeyboardFocused" Value="True"'
        }
    }

    It 'defines the button styles in every dialog window that uses them' {
        $bad = @()
        foreach ($w in $script:Windows) {
            foreach ($key in 'ApButton', 'ApButtonAccent') {
                $uses = $w.Text -match ('\{(Dynamic|Static)Resource ' + $key + '\}')
                $defines = $w.Text -match ('x:Key="' + $key + '"')
                if ($uses -and -not $defines) { $bad += ('{0}:{1}' -f $key, $w.Line) }
            }
        }
        $bad | Should -BeNullOrEmpty
    }

    It 'gives the Workbench buttons the shared style' {
        $script:WorkbenchXaml | Should -Not -Match 'x:Key="Wb(Accent|Inline)?Button"[^>]*BasedOn="\{StaticResource MahApps'
        foreach ($key in 'WbButton', 'WbInlineButton') {
            (Get-StyleBlock $script:WorkbenchXaml $key) | Should -Match 'BasedOn="\{StaticResource ApButton\}"'
        }
        (Get-StyleBlock $script:WorkbenchXaml 'WbAccentButton') | Should -Match 'BasedOn="\{StaticResource ApButtonAccent\}"'
    }
}

Describe 'Dialog title bar' {
    # Show-ThemedMessage copies the owner's inactive brush into the inactive
    # slot. A dialog that copies the active brush keeps a full-strength title
    # bar when it loses focus, while the owner dims.
    It 'copies the owner inactive title and glow brushes' {
        $text = $script:ChromeFunction.Extent.Text
        $text | Should -Match '\$Dialog\.NonActiveWindowTitleBrush\s*=\s*\$Owner\.NonActiveWindowTitleBrush'
        $text | Should -Match '\$Dialog\.NonActiveGlowBrush\s*=\s*\$Owner\.NonActiveGlowBrush'
    }
}
