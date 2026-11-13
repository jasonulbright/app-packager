#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
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

    function Get-GridCell($Element) {
        $attr = { param($name, $default) $v = $Element.GetAttribute($name); if ($v) { [int]$v } else { $default } }
        $r = & $attr 'Grid.Row' 0; $c = & $attr 'Grid.Column' 0
        @{ R0 = $r; R1 = $r + (& $attr 'Grid.RowSpan' 1) - 1; C0 = $c; C1 = $c + (& $attr 'Grid.ColumnSpan' 1) - 1 }
    }
    # The Opacity a style sets while the control is disabled; 1 when it sets none.
    function Get-DisabledOpacity([string]$Style) {
        $m = [regex]::Match($Style, '<Trigger\s+Property="IsEnabled"\s+Value="False">.*?<Setter\s+Property="Opacity"\s+Value="([0-9.]+)"', 'Singleline')
        if ($m.Success) { [double]::Parse($m.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture) } else { 1.0 }
    }
    function Get-StyleBlock([string]$Xaml, [string]$Key) {
        $m = [regex]::Match($Xaml, '<Style\s+x:Key="' + $Key + '"[^>]*>.*?</Style>', 'Singleline')
        if ($m.Success) { $m.Value } else { '' }
    }
}

Describe 'List item style' {
    # An ItemContainerStyle without BasedOn drops the MahApps item template, so
    # the selected item uses system colors: light grey with light text in Dark.
    It 'bases every ListBoxItem style on the shared selected-state style' {
        $bad = @()
        foreach ($m in [regex]::Matches($script:WorkbenchXaml, '<Style\s+TargetType="ListBoxItem"[^>]*>')) {
            if ($m.Value -notmatch 'BasedOn="\{StaticResource ThemedListBoxItemStyle\}"') { $bad += 'WorkbenchWindow.xaml' }
        }
        foreach ($b in $script:Blocks) {
            foreach ($m in [regex]::Matches($b.Text, '<Style\s+TargetType="ListBoxItem"[^>]*>')) {
                if ($m.Value -notmatch 'BasedOn="\{StaticResource ThemedListBoxItemStyle\}"') { $bad += ('start-apppackager.ps1:' + $b.Line) }
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
        # The disabled state dims the text with a gray foreground. It never changes the opacity.
        $style | Should -Match 'Property="Foreground" Value="\{DynamicResource MahApps\.Brushes\.Gray\.MouseOver\}"'
        (Get-DisabledOpacity $style) | Should -Be 1
        # The override takes its color from the active theme, not a fixed value.
        $style | Should -Match '\{DynamicResource MahApps\.Colors\.'
    }

    It 'reads a disabled opacity of 1 as not dimmed' {
        $style = '<Style><Style.Triggers><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="1"/></Trigger></Style.Triggers></Style>'
        (Get-DisabledOpacity $style) | Should -Be 1
    }}

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

    It 'defines the shared button style with a 1 px border and a keyboard focus trigger, once' {
        $main = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\MainWindow.xaml'))
        $style = Get-StyleBlock $main 'ThemedButtonStyle'
        $style | Should -Match 'Property="BorderThickness" Value="1"'
        $main | Should -Match '<Trigger Property="IsKeyboardFocused" Value="True">'
    }

    It 'takes the shared styles through the marker in every dialog window' {
        $bad = @()
        foreach ($w in $script:Windows) { if ($w.Text -notmatch '<!--SHARED_STYLES-->') { $bad += $w.Line } }
        $bad | Should -BeNullOrEmpty
    }

    It 'defines no Workbench button style of its own' {
        $script:WorkbenchXaml | Should -Match '<!--SHARED_STYLES-->'
        $script:WorkbenchXaml | Should -Not -Match 'x:Key="(Wb(Accent|Inline)?Button|ApButton|ApButtonAccent)"'
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

Describe 'Options panels' {
    # The panels are loaded into the Options window and resolve ApButton from
    # its Window.Resources, so the styles must sit there, not in a child.
    It 'gives the Options window, which hosts the panels, the shared styles' {
        $options = @($script:Windows | Where-Object { $_.Text -match 'x:Name="treeNav"' })
        $options.Count | Should -Be 1
        $options[0].Text | Should -Match '<!--SHARED_STYLES-->'
    }
}
Describe 'Modeless dialogs' {
    # A modeless dialog stays open while the theme toggles. The toggle handler
    # re-applies the owner chrome, including the inactive brushes, to it.
    It 'syncs the owned windows when the theme toggles' {
        $m = [regex]::Match($script:Shell, '(?s)\$toggleTheme\.Add_Toggled\(\{.*?\n\}\)')
        $m.Success | Should -BeTrue
        $m.Value | Should -Match 'OwnedWindows'
        $m.Value | Should -Match 'Set-DialogChromeFromOwner'
    }
}

Describe 'Workbench footer' {
    # The save-state message and the action buttons must sit in grid cells
    # that do not intersect, so a wrapped button line cannot cover the message.
    It 'keeps the save-state message out of the action buttons cells' {
        $x = [xml]$script:WorkbenchXaml
        $ns = New-Object System.Xml.XmlNamespaceManager $x.NameTable
        $ns.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $msg = $x.SelectSingleNode('//p:TextBlock[@x:Name="txtSaveState"]', $ns)
        $panel = $x.SelectSingleNode('//p:Button[@x:Name="btnSave"]', $ns).ParentNode
        $a = Get-GridCell $msg; $b = Get-GridCell $panel
        $rowsOverlap = ($a.R0 -le $b.R1) -and ($b.R0 -le $a.R1)
        $colsOverlap = ($a.C0 -le $b.C1) -and ($b.C0 -le $a.C1)
        ($rowsOverlap -and $colsOverlap) | Should -BeFalse
    }
}