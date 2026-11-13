#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the theme rules. They parse the files, so they run on Windows PowerShell 5.1 and PowerShell 7.
# One test creates a real button in its own runspace to check the foreground helper.
#   (a) Every blue background in XAML and PowerShell has a white foreground, also in hover, pressed and selected states.
#   (b) No font size, weight, style or text decoration depends on the theme. Site code and SMS Provider titles are bold 14 pt, values are not.
#   (c) No banned word is in a tooltip, a UI string, the README or the user guide. Tooltips are one short sentence.
#   (d) Every button takes the shared themed style, has no size of its own, and no style holds a color literal or an opacity state.
#   (e) Titles are in Title Case.
#   (f) No raw exception text reaches the UI.

BeforeDiscovery {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocsD = @(Get-XamlDocument)
    $script:WindowDocsD = @($script:XamlDocsD | Where-Object { $_.Kind -eq 'window' })
}

BeforeAll {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocs = @(Get-XamlDocument)
    $script:MainDoc = [xml](Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw)
    $script:ShellPath = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $script:SuitePath = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
    $script:ShellText = Get-Content -LiteralPath $script:ShellPath -Raw
}

Describe 'The scans find the documents' {
    It 'finds the main window, the Workbench, the dialogs and the panels' {
        $script:XamlDocs.Count | Should -BeGreaterOrEqual 25
        @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' }).Count | Should -Be 1
        @($script:XamlDocs | Where-Object { $_.Source -eq 'WorkbenchWindow.xaml' }).Count | Should -Be 1
        @($script:XamlDocs | Where-Object { $_.Kind -eq 'window' }).Count | Should -BeGreaterOrEqual 15
        @($script:XamlDocs | Where-Object { $_.Kind -eq 'panel' }).Count | Should -BeGreaterOrEqual 5
    }

    It 'expands the shared style marker in every window of the start script' {
        $script:XamlDocs | ForEach-Object { $_.Text | Should -Not -Match '<!--SHARED_STYLES-->' }
        $intMarker = ([regex]::Matches($script:ShellText, [regex]::Escape($script:SharedStyleMarker))).Count 
        $intMarker | Should -BeGreaterOrEqual 15
        $script:ShellText | Should -Match 'function Expand-SharedStyle'
        $script:ShellText | Should -Match 'function Get-SharedStyleXaml'
    }

    It 'has the shared style marker in the raw text of every window of the start script (15) and of the Workbench' {
        $windows = @($script:XamlDocs | Where-Object { $_.Kind -eq 'window' })
        $windows.Count | Should -Be 15
        foreach ($w in $windows) { $w.RawText | Should -Match '<!--SHARED_STYLES-->' -Because $w.Source }
        $wb = @($script:XamlDocs | Where-Object { $_.Source -eq 'WorkbenchWindow.xaml' })[0]
        $wb.RawText | Should -Match '<!--SHARED_STYLES-->'
    }

    It 'does not let a window without the marker take the styles of the main window (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
        [xml]$window = "<Window $ns><Window.Resources><Style x:Key=`"Own`" TargetType=`"Button`"/></Window.Resources><Button Style=`"{DynamicResource ThemedButtonStyle}`"/></Window>"
        [xml]$panel = "<DockPanel $ns><Button Style=`"{DynamicResource ThemedButtonStyle}`"/></DockPanel>"
        (Find-StyleNode -Doc $window -Key 'ThemedButtonStyle') | Should -BeNullOrEmpty
        (Find-StyleNode -Doc $panel -Key 'ThemedButtonStyle') | Should -Not -BeNullOrEmpty
    }

    It 'lists in the start script the same shared styles as the helper' {
        $strKeys = [regex]::Match($script:ShellText, '(?s)function Get-SharedStyleXaml.*?\$keys = @\(([^)]*)\)').Groups[1].Value
        $arrShell = @([regex]::Matches($strKeys, "'(\w+)'") | ForEach-Object { $_.Groups[1].Value })
        $arrShell.Count | Should -BeGreaterOrEqual 12
        $strHelper = Get-Content -LiteralPath "$PSScriptRoot\ThemeGuardHelpers.ps1" -Raw
        foreach ($strKey in $arrShell) { $strHelper | Should -Match ([regex]::Escape("'$strKey'")) -Because $strKey }
    }
}

Describe 'Blue backgrounds have white text (a)' {
    It 'sets a white foreground on every blue background in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-BlueWithoutWhite -Text $_.Text -Source $_.Source) | Should -BeNullOrEmpty
    }

    It 'recognizes blue and white brushes' {
        Test-BlueBrush '{DynamicResource MahApps.Brushes.Accent}' | Should -BeTrue
        Test-BlueBrush '{DynamicResource MahApps.Colors.Accent}' | Should -BeTrue
        Test-BlueBrush '{DynamicResource MahApps.Brushes.Highlight}' | Should -BeTrue
        Test-BlueBrush '{DynamicResource {x:Static SystemColors.HighlightBrushKey}}' | Should -BeTrue
        Test-BlueBrush 'SystemColors.HighlightColor' | Should -BeTrue
        Test-BlueBrush '#0078D4' | Should -BeTrue
        Test-BlueBrush '#07D' | Should -BeTrue
        Test-BlueBrush '#1E1E1E' | Should -BeFalse
        Test-BlueBrush '{DynamicResource MahApps.Brushes.ThemeBackground}' | Should -BeFalse
        Test-WhiteBrush '{DynamicResource MahApps.Brushes.IdealForeground}' | Should -BeTrue
        Test-WhiteBrush 'Black' | Should -BeFalse
    }

    It 'flags a blue background that has no white foreground, with every property form (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        @(Get-BlueWithoutWhite -Text "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`" Value=`"#0078D4`"/></Style>" -Source 's').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`" Value=`"#0078D4`"/><Setter Property=`"Foreground`" Value=`"White`"/></Style>" -Source 's').Count | Should -Be 0
        @(Get-BlueWithoutWhite -Text "<Style $ns TargetType=`"Button`"><Setter Property=`"Border.Background`" Value=`"{DynamicResource MahApps.Brushes.Highlight}`"/></Style>" -Source 's').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`"><Setter.Value><SolidColorBrush Color=`"{DynamicResource MahApps.Colors.Accent}`"/></Setter.Value></Setter></Style>" -Source 's').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text "<Border $ns Background=`"{DynamicResource MahApps.Brushes.Accent}`"/>" -Source 's').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text "<Border $ns><Border.Background><SolidColorBrush Color=`"#0078D4`"/></Border.Background></Border>" -Source 's').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text "<Border $ns TextElement.Foreground=`"White`"><Border.Background><SolidColorBrush Color=`"#0078D4`"/></Border.Background></Border>" -Source 's').Count | Should -Be 0
        $trigger = "<Style $ns TargetType=`"ListBoxItem`"><Style.Triggers><Trigger Property=`"IsSelected`" Value=`"True`"><Setter Property=`"Background`" Value=`"{DynamicResource MahApps.Brushes.Highlight}`"/></Trigger></Style.Triggers></Style>"
        @(Get-BlueWithoutWhite -Text $trigger -Source 's').Count | Should -Be 1
    }

    It 'sets no blue color as a Background in <Name> and each CSS rule with a blue background has white text' -ForEach @(@{ Name = 'start-apppackager.ps1' }) {
        $path = Join-Path $script:RepoRoot $Name
        $bad = foreach ($s in @(Get-ScriptString -Path $path)) {
            if ($s -notmatch '<' -and (Test-BlueBrush $s)) { 'literal ' + $s }
            foreach ($m in [regex]::Matches($s, '\{([^}]*)\}')) {
                $rule = $m.Groups[1].Value
                $bg = [regex]::Match($rule, '(?i)(?<![-\w])background(?:-color)?\s*:\s*(#[0-9A-F]{6,8})\b')
                if ($bg.Success -and (Test-BlueBrush $bg.Groups[1].Value) -and $rule -notmatch '(?i)(?<![-\w])color\s*:\s*(#fff(fff)?|white)\b') { 'css ' + $rule }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives every list, combo box, tree and grid a selected state with white text on blue in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-SelectedStateViolation -Text $_.Text -Source $_.Source) | Should -BeNullOrEmpty
    }

    It 'finds the item controls of the documents (the scan is not empty)' {
        $count = @{ ListBox = 0; ComboBox = 0; DataGrid = 0; TreeView = 0 }
        foreach ($d in $script:XamlDocs) { [xml]$x = $d.Text; foreach ($k in @($count.Keys)) { $count[$k] += @($x.SelectNodes(("//*[local-name()='{0}']" -f $k))).Count } }
        $count.ListBox | Should -BeGreaterOrEqual 3
        $count.ComboBox | Should -BeGreaterOrEqual 15
        $count.DataGrid | Should -BeGreaterOrEqual 10
        $count.TreeView | Should -BeGreaterOrEqual 1
    }

    It 'flags a combo box without the IsHighlighted state, a list without an item container style and a grid without row styles (the check itself)' {
        $open = '<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">'
        $sel = '<Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{DynamicResource MahApps.Brushes.Highlight}"/><Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.IdealForeground}"/></Trigger>'
        $hi = $sel.Replace('IsSelected', 'IsHighlighted')
        $bad = '<Style x:Key="I" TargetType="ComboBoxItem"><Style.Triggers>' + $sel + '</Style.Triggers></Style>'
        $good = '<Style x:Key="I" TargetType="ComboBoxItem"><Style.Triggers>' + $sel + $hi + '</Style.Triggers></Style>'
        $box = '<ComboBox ItemContainerStyle="{DynamicResource I}"/>'
        @(Get-SelectedStateViolation -Text ($open + '<Grid.Resources>' + $bad + '</Grid.Resources>' + $box + '</Grid>') -Source 's').Count | Should -Be 1
        @(Get-SelectedStateViolation -Text ($open + '<Grid.Resources>' + $good + '</Grid.Resources>' + $box + '</Grid>') -Source 's').Count | Should -Be 0
        @(Get-SelectedStateViolation -Text ($open + '<ListBox/></Grid>') -Source 's').Count | Should -Be 1
        @(Get-SelectedStateViolation -Text ($open + '<DataGrid/></Grid>') -Source 's').Count | Should -Be 2
    }

    It 'sets the foreground of every state trigger that sets a Background in <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $bad = foreach ($t in @($x.SelectNodes("//*[local-name()='Trigger'][@Property='IsSelected' or @Property='IsHighlighted' or @Property='IsMouseOver' or @Property='IsPressed']"))) {
            $bg = Get-SetterChild -Node $t -LeafPattern 'Background$' | Select-Object -First 1
            $fg = Get-SetterChild -Node $t -LeafPattern 'Foreground$' | Select-Object -First 1
            if ($bg -and -not $fg) { '{0} trigger sets a Background and no Foreground' -f $t.GetAttribute('Property') }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'shows the check box of a selected tree item in white text' {
        $st = $script:MainDoc.SelectSingleNode("//*[local-name()='Style'][@TargetType='CheckBox'][not(@*[local-name()='Key'])]")
        $st | Should -Not -BeNullOrEmpty
        $st.OuterXml | Should -Match 'AncestorType=TreeViewItem'
        $st.OuterXml | Should -Match 'IdealForeground'
    }
}

Describe 'Theme helpers of SuiteCommon behave and are used (a, b)' {
    BeforeAll {
        $script:SuiteAst = [System.Management.Automation.Language.Parser]::ParseFile($script:SuitePath, [ref]$null, [ref]$null)
        function Get-SuiteFunction {
            param([string]$Name)
            $script:SuiteAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        }
        function Get-CommandCount {
            param($Ast, [string]$Name)
            @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $Name }, $true)).Count
        }
        # The test loads the WPF assemblies itself, so that it does not depend on an earlier test or on the shell.
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        . ([scriptblock]::Create((Get-SuiteFunction 'Set-ThemeForeground').Extent.Text))
    }

    It 'Set-ThemeForeground sets the on-blue brush in light mode and clears it in dark mode' {
        $c = [pscustomobject]@{ Foreground = $null }
        Set-ThemeForeground -Control $c -IsDark $false -OnBlue 'White'
        $c.Foreground | Should -Be 'White'
        Set-ThemeForeground -Control $c -IsDark $true -OnBlue 'White'
        $c.Foreground | Should -BeNullOrEmpty
        { Set-ThemeForeground -Control $null -IsDark $false -OnBlue 'White' } | Should -Not -Throw
    }

    It 'Set-ThemeForeground sets and clears the foreground of a real Button (STA runspace)' {
        $fnText = (Get-SuiteFunction 'Set-ThemeForeground').Extent.Text
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()
        $ps = $null
        try {
            $ps = [powershell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript({
                    param($fnText)
                    Add-Type -AssemblyName PresentationFramework, WindowsBase
                    . ([scriptblock]::Create($fnText))
                    $button = New-Object System.Windows.Controls.Button
                    $white = [System.Windows.Media.Brushes]::White
                    Set-ThemeForeground -Control $button -IsDark $false -OnBlue $white
                    $light = [object]::ReferenceEquals($button.Foreground, $white)
                    Set-ThemeForeground -Control $button -IsDark $true -OnBlue $white
                    $dark = [object]::ReferenceEquals($button.ReadLocalValue([System.Windows.Controls.Control]::ForegroundProperty), [System.Windows.DependencyProperty]::UnsetValue)
                    [pscustomobject]@{ Light = $light; Dark = $dark }
                }).AddArgument($fnText)
            $result = @($ps.Invoke())
            $ps.HadErrors | Should -BeFalse
            $result.Count | Should -Be 1
            $result[0].Light | Should -BeTrue
            $result[0].Dark | Should -BeTrue
        }
        finally {
            if ($ps) { $ps.Dispose() }
            $rs.Dispose()
        }
    }

    It '<Name> calls Set-ThemeForeground and Set-ConnectionTextTheme' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ThemeForeground') | Should -BeGreaterThan 0
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ConnectionTextTheme') | Should -BeGreaterThan 0
    }

    It 'the script starts the theme, themes the sidebar buttons at start and at each toggle' {
        $shell = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
        (Get-CommandCount $shell 'Initialize-SuiteTheme') | Should -BeGreaterThan 0
        (Get-CommandCount $shell 'Set-ButtonTheme') | Should -BeGreaterOrEqual 3
        $script:ShellText | Should -Match 'Set-ButtonTheme -IsDark \(\$toggleTheme\.IsOn -eq \$true\)'
        $script:ShellText | Should -Match '-WorkflowButtons \$script:WorkflowButtons'
        $script:ShellText | Should -Not -Match 'WindowTitleBrush\s*=\s*\[?(New-Object|#)'
    }

    It 'registers every sidebar button for the theme' {
        $reg = [regex]::Match($script:ShellText, '\$script:WorkflowButtons = @\(([^)]*)\)').Groups[1].Value + ' ' + [regex]::Match($script:ShellText, '\$script:OptionsButtons\s*=\s*@\(([^)]*)\)').Groups[1].Value
        $names = @([regex]::Matches($reg, '\$(\w+)') | ForEach-Object { $_.Groups[1].Value })
        $main = @($script:MainDoc.SelectNodes("//*[local-name()='Button'][contains(@Style,'WorkflowButtonStyle') or contains(@Style,'OptionsButtonStyle')]"))
        $main.Count | Should -BeGreaterOrEqual 14
        foreach ($b in $main) { $names | Should -Contain $b.GetAttribute('Name', $script:XamlKeyNamespace) }
    }
}

Describe 'Fonts do not depend on the theme (b)' {
    It 'has no font property in a theme branch of <Name>' -ForEach @(@{ Name = 'start-apppackager.ps1' }, @{ Name = 'Lib\SuiteCommon\SuiteCommon.psm1' }) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot $Name), [ref]$null, [ref]$null)
        $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        $bad = foreach ($a in $assigns) {
            $member = [string]$a.Left.Member.Value
            if ($script:ThemeProperty -notcontains $member) { continue }
            $p = $a.Parent
            while ($p) {
                $cond = ''
                if ($p -is [System.Management.Automation.Language.IfStatementAst]) { $cond = (@($p.Clauses | ForEach-Object { $_.Item1.Extent.Text }) -join ' ') }
                elseif ($p -is [System.Management.Automation.Language.SwitchStatementAst]) { $cond = $p.Condition.Extent.Text }
                if ($cond -match '(?i)IsDark|DarkMode|Light|Theme|IsOn') { '{0}:{1} sets {2} in a theme branch' -f $Name, $a.Extent.StartLineNumber, $member; break }
                $p = $p.Parent
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'detects a font property in a theme branch (the check itself)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('if ($isDark) { $b.Background = 1; $b.FontWeight = 2 } else { $b.Background = 3 }', [ref]$null, [ref]$null)
        $assigns = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $script:ThemeProperty -contains [string]$n.Left.Member.Value }, $true))
        $assigns.Count | Should -Be 1
        $assigns[0].Parent.Parent | Should -BeOfType [System.Management.Automation.Language.IfStatementAst]
    }

    It 'sets no font property in the theme functions of SuiteCommon except the connection titles' {
        foreach ($name in 'Set-ThemeForeground', 'Update-SidebarButtonTheme', 'Set-ButtonTheme', 'Set-DialogTheme', 'Update-TitleBarBrushes', 'Initialize-SuiteTheme') {
            $fn = $script:SuiteAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Not -Match '\.(FontSize|FontWeight|FontStyle|TextDecorations)\s*=' -Because $name
        }
    }

    It 'makes the title of each connection box bold and 14 pt or larger, the same in both themes' {
        $text = Get-Content -LiteralPath $script:SuitePath -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Not -Match '\$box\.(FontSize|FontWeight)\s*='
    }

    It 'uses no theme trigger on a font property in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-ThemeFontTrigger -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'flags a trigger that sets a font for a theme and not one for the selected state (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $theme = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsDark}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        $state = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsSelected}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        @(Get-ThemeFontTrigger -Text $theme).Count | Should -Be 1
        @(Get-ThemeFontTrigger -Text $state).Count | Should -Be 0
    }

    It 'uses no SemiBold title or label in <_.Source>' -ForEach $script:XamlDocsD {
        $_.Text | Should -Not -Match 'FontWeight(=|" Value=)"SemiBold"'
    }

    It 'sets the title of the site code and SMS Provider boxes to bold 14 pt or more and the value to a normal weight below 14 pt in <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        foreach ($name in 'txtSC', 'txtProvider', 'txtWizSC', 'txtWizProvider', 'txtSite') {
            $box = $x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $name))
            if (-not $box) { continue }
            $row = $box.GetAttribute('Grid.Row')
            $title = $x.SelectSingleNode(("//*[local-name()='TextBlock'][@Grid.Row='{0}'][@Grid.Column='0']" -f $row))
            if (-not $title) { $title = $box.PreviousSibling; while ($title -and $title.NodeType -ne 'Element') { $title = $title.PreviousSibling } }
            $title | Should -Not -BeNullOrEmpty -Because $name
            [double]$title.GetAttribute('FontSize') | Should -BeGreaterOrEqual 14 -Because ('title of ' + $name)
            $title.GetAttribute('FontWeight') | Should -Be 'Bold' -Because ('title of ' + $name)
            $box.GetAttribute('FontWeight') | Should -Not -Be 'Bold' -Because ('value of ' + $name)
            if ($box.GetAttribute('FontSize')) { [double]$box.GetAttribute('FontSize') | Should -BeLessThan 14 -Because ('value of ' + $name) }
        }
    }

    It 'finds the site code and SMS Provider boxes (the scan is not empty)' {
        $names = foreach ($d in $script:XamlDocs) { [xml]$x = $d.Text; foreach ($n in 'txtSC', 'txtProvider', 'txtWizSC', 'txtWizProvider') { if ($x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $n))) { $n } } }
        @($names | Sort-Object -Unique) | Should -Be @('txtProvider', 'txtSC', 'txtWizProvider', 'txtWizSC')
    }

    It 'uses the term SMS Provider and no other name for it in the XAML text' {
        $text = foreach ($d in $script:XamlDocs) { Get-XamlUiText -Text $d.Text }
        @($text | Where-Object { $_ -cmatch 'Provider Machine|SMS provider' }) | Should -BeNullOrEmpty
    }

    It 'bolds every field label that ends with a colon in <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $bad = foreach ($t in @($x.SelectNodes("//*[local-name()='TextBlock'][substring(@Text, string-length(@Text)) = ':']"))) { if ($t.GetAttribute('FontWeight') -ne 'Bold' -and $t.GetAttribute('Style') -notmatch 'Label') { $t.GetAttribute('Text') } }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'No banned word in the UI text and the documents (c)' {
    It 'finds the banned word <Word>' -ForEach @(@{ Word = 'simply' }, @{ Word = 'Helps you' }, @{ Word = 'streamlines' }, @{ Word = 'just' }) {
        $Word | Should -Match $script:BannedPattern
    }

    It 'has no banned word in the XAML text of <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-XamlUiText -Text $_.Text | Where-Object { $_ -match $script:BannedPattern }) | Should -BeNullOrEmpty
    }

    It 'has no banned word in the strings of <Name>' -ForEach @(@{ Name = 'start-apppackager.ps1' }) {
        @(Get-ScriptString -Path (Join-Path $script:RepoRoot $Name) | Where-Object { $_ -match $script:BannedPattern }) | Should -BeNullOrEmpty
    }

    It 'has no banned word in the README or the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        $files.Count | Should -BeGreaterThan 5
        $bad = foreach ($f in $files) {
            $n = 0
            foreach ($line in (Get-Content -LiteralPath $f.FullName)) { $n++; if ($line -match $script:BannedPattern) { '{0}:{1}: {2}' -f $f.Name, $n, $Matches[0] } }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives every tooltip in <_.Source> one short sentence' -ForEach $script:XamlDocsD {
        @(Get-XamlToolTip -Text $_.Text | Where-Object { -not (Test-OneSentence $_) }) | Should -BeNullOrEmpty
    }

    It 'checks one sentence (the check itself)' {
        Test-OneSentence 'Deploy an application to a device collection.' | Should -BeTrue
        Test-OneSentence 'Pick the object. Then select OK.' | Should -BeFalse
        Test-OneSentence 'pick the object.' | Should -BeFalse
        Test-OneSentence 'Pick the object' | Should -BeFalse
    }

    It 'writes ConfigMgr and never MECM in the UI text and the documents' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        # A code span or a code fence may hold the name of a parameter value (-Target MECM). Prose never does.
        $bad = foreach ($f in $files) { $prose = [regex]::Replace([regex]::Replace((Get-Content -LiteralPath $f.FullName -Raw), '(?s)```.*?```', ''), '`[^`\r\n]*`', ''); if ($prose -cmatch '\bMECM\b') { $f.Name } }
        @($bad) | Should -BeNullOrEmpty
        $text = foreach ($d in $script:XamlDocs) { Get-XamlUiText -Text $d.Text }
        @($text | Where-Object { $_ -cmatch '\bMECM\b' }) | Should -BeNullOrEmpty
    }
}

Describe 'Every button takes the shared themed style (d)' {
    It 'gives each Button in <_.Source> a style that ends at ThemedButtonStyle, no size of its own and no color literal' -ForEach $script:XamlDocsD {
        @(Get-ButtonViolation -Text $_.Text -Source $_.Source) | Should -BeNullOrEmpty
    }

    It 'finds the buttons of the documents (the scan is not empty)' {
        $count = 0
        foreach ($d in $script:XamlDocs) { [xml]$x = $d.Text; $count += @($x.SelectNodes("//*[local-name()='Button']")).Count }
        $count | Should -BeGreaterThan 90
    }

    It 'flags a button with no style, a size, a literal color and a style of another base (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
        $themed = '<Style x:Key="ThemedButtonStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}"/><Style x:Key="S" TargetType="Button" BasedOn="{StaticResource ThemedButtonStyle}"/>'
        @(Get-ButtonViolation -Text "<Grid $ns><Button x:Name=`"a`"/></Grid>" -Source 's').Count | Should -Be 1
        @(Get-ButtonViolation -Text "<Grid $ns><Grid.Resources>$themed</Grid.Resources><Button Style=`"{StaticResource S}`" Height=`"28`"/></Grid>" -Source 's').Count | Should -Be 1
        @(Get-ButtonViolation -Text "<Grid $ns><Grid.Resources>$themed</Grid.Resources><Button Style=`"{StaticResource S}`" MinWidth=`"90`" Padding=`"4`"/></Grid>" -Source 's').Count | Should -Be 2
        @(Get-ButtonViolation -Text "<Grid $ns><Grid.Resources>$themed</Grid.Resources><Button Style=`"{StaticResource S}`" Background=`"Orange`"/></Grid>" -Source 's').Count | Should -Be 1
        @(Get-ButtonViolation -Text "<Grid $ns><Grid.Resources>$themed<Style x:Key=`"M`" TargetType=`"Button`" BasedOn=`"{StaticResource MahApps.Styles.Button.Square.Accent}`"/></Grid.Resources><Button Style=`"{StaticResource M}`"/></Grid>" -Source 's').Count | Should -Be 2
        @(Get-ButtonViolation -Text "<Grid $ns><Grid.Resources>$themed</Grid.Resources><Button Style=`"{StaticResource S}`" Margin=`"0,0,8,0`"/></Grid>" -Source 's').Count | Should -Be 0
    }

    It 'sets no size on a button that the start script creates, and gives it the inline style' {
        @(Get-CodeButtonSize -Path $script:ShellPath) | Should -BeNullOrEmpty
        $buttons = @(Get-CodeControl -Path $script:ShellPath -TypeName 'Button')
        $buttons.Count | Should -BeGreaterOrEqual 2
        foreach ($b in $buttons) { Get-CodeStyleKey -Function $b.Function -Variable $b.Variable | Should -BeIn @('DialogButton', 'DialogAccentButton', 'InlineButtonStyle') -Because $b.Variable }
        $script:ShellText | Should -Not -Match 'MahApps\.Styles\.Button'
    }

    It 'detects a size on a code-built button (the check itself)' {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('themeguard-' + [guid]::NewGuid().ToString('N') + '.ps1')
        try {
            [System.IO.File]::WriteAllText($tmp, "`$a = New-Object System.Windows.Controls.Button`n`$a.MinWidth = 90`n`$a.Margin = 1`n`$t = New-Object System.Windows.Controls.TextBox`n`$t.Width = 5`n")
            @(Get-CodeButtonSize -Path $tmp).Count | Should -Be 1
        }
        finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }

    It 'sets no color literal in any Style or ControlTemplate of <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        @(Get-HardColor -Nodes @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]')) -Label $_.Source) | Should -BeNullOrEmpty
    }

    It 'sets no color literal outside the styles of <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-LiteralColorOutsideStyle -Text $_.Text -Source $_.Source) | Should -BeNullOrEmpty
    }

    It 'scans the styles and templates of the main window (the scan is not empty)' {
        @($script:MainDoc.SelectNodes('//*[local-name()="Style"]')).Count | Should -BeGreaterThan 15
        @($script:MainDoc.SelectNodes('//*[local-name()="ControlTemplate"]')).Count | Should -BeGreaterThan 0
    }

    It 'detects a color literal in a style, a template, a named color and an attribute (the check itself)' {
        [xml]$x = '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Style x:Key="S" TargetType="Button"><Setter Property="Foreground" Value="White"/><Setter Property="Background"><Setter.Value><SolidColorBrush Color="#1E1E1E"/></Setter.Value></Setter><Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Accent}"/></Style><ControlTemplate x:Key="T" TargetType="Button"><Border BorderBrush="#555555" Background="{TemplateBinding Background}"><Border.Background><SolidColorBrush Color="#123456"/></Border.Background></Border></ControlTemplate></ResourceDictionary>'
        $found = @(Get-HardColor -Nodes @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]')) -Label 'sample')
        $found.Count | Should -Be 4
        ($found -join ' ') | Should -Match '#123456'
        ($found -join ' ') | Should -Match 'White'
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        @(Get-LiteralColorOutsideStyle -Text "<Border $ns Background=`"#FFE07A5F`"/>" -Source 's').Count | Should -Be 1
        @(Get-LiteralColorOutsideStyle -Text "<Border $ns Background=`"Transparent`"/>" -Source 's').Count | Should -Be 1
        @(Get-LiteralColorOutsideStyle -Text "<Border $ns Background=`"{DynamicResource MahApps.Brushes.Gray8}`"/>" -Source 's').Count | Should -Be 0
    }

    It 'changes no opacity in a hover, pressed, disabled, focus or selected state, and none on text, in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-OpacityInStateTrigger -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'detects an opacity in a state trigger and on text (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $bad = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><Trigger Property=`"IsMouseOver`" Value=`"True`"><Setter TargetName=`"Bd`" Property=`"Opacity`" Value=`"0.85`"/></Trigger></ControlTemplate.Triggers></ControlTemplate>"
        $ok = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><Trigger Property=`"IsMouseOver`" Value=`"True`"><Setter TargetName=`"Bd`" Property=`"BorderBrush`" Value=`"{DynamicResource MahApps.Brushes.ThemeForeground}`"/></Trigger></ControlTemplate.Triggers></ControlTemplate>"
        @(Get-OpacityInStateTrigger -Text $bad).Count | Should -Be 1
        @(Get-OpacityInStateTrigger -Text $ok).Count | Should -Be 0
        @(Get-OpacityInStateTrigger -Text "<TextBlock $ns Opacity=`"0.75`"/>").Count | Should -Be 1
        @(Get-OpacityInStateTrigger -Text "<Border $ns Background=`"{DynamicResource MahApps.Brushes.Highlight}`" Opacity=`"0.5`"/>").Count | Should -Be 1
        @(Get-OpacityInStateTrigger -Text "<SolidColorBrush $ns Color=`"{DynamicResource MahApps.Colors.Gray10}`" Opacity=`"0.85`"/>").Count | Should -Be 1
        $multi = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><MultiTrigger><MultiTrigger.Conditions><Condition Property=`"IsMouseOver`" Value=`"True`"/></MultiTrigger.Conditions><Setter Property=`"Opacity`" Value=`"0.8`"/></MultiTrigger></ControlTemplate.Triggers></ControlTemplate>"
        @(Get-OpacityInStateTrigger -Text $multi).Count | Should -Be 1
        @(Get-LiteralColorOutsideStyle -Text "<Border $ns><Border.Background><SolidColorBrush Color=`"Red`"/></Border.Background></Border>" -Source 's').Count | Should -Be 1
    }

    It 'documents the one secondary style and the one primary style, with equal sizes and theme colors only' {
        $main = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
        $main | Should -Match 'The one secondary style'
        function Get-StyleSetter($key) {
            $st = $script:MainDoc.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
            $st | Should -Not -BeNullOrEmpty -Because $key
            $h = @{}
            foreach ($se in @($st.SelectNodes("*[local-name()='Setter']"))) { $h[$se.GetAttribute('Property')] = Get-SetterValue $se }
            $h
        }
        $sec = Get-StyleSetter 'DialogButton'
        $pri = Get-StyleSetter 'DialogAccentButton'
        $sec['Height'] | Should -Be $pri['Height']
        $sec['MinWidth'] | Should -Be $pri['MinWidth']
        $sec['Padding'] | Should -Be $pri['Padding']
        $sec['Background'] | Should -Match 'MahApps\.Brushes\.Gray'
        $sec['Foreground'] | Should -Match 'MahApps\.Brushes\.ThemeForeground'
        $pri['Background'] | Should -Match 'MahApps\.Brushes\.Highlight'
        Test-WhiteBrush $pri['Foreground'] | Should -BeTrue
        $keys = @($script:MainDoc.SelectNodes("//*[local-name()='Style'][@TargetType='Button']") | Where-Object { @(Get-SetterChild -Node $_ -LeafPattern '^Background$').Count -gt 0 } | ForEach-Object { $_.GetAttribute('Key', $script:XamlKeyNamespace) })
        ($keys | Sort-Object) | Should -Be @('DialogAccentButton', 'DialogButton', 'LinkButtonStyle')
    }

    It 'defines the button template and the button style once, in MainWindow.xaml, and in no dialog' {
        foreach ($k in 'ThemedButtonTemplate', 'ThemedButtonStyle', 'DialogButton', 'DialogAccentButton', 'InlineButtonStyle', 'GridRowStyle', 'GridCellStyle') {
            @($script:MainDoc.SelectNodes("//*[@*[local-name()='Key']='$k']")).Count | Should -Be 1 -Because $k
        }
        $script:ShellText | Should -Not -Match 'x:Key="(DialogButton|DialogAccentButton|ThemedButtonStyle|ThemedButtonTemplate|InlineButtonStyle|GridRowStyle|GridCellStyle|ApButton|ApButtonAccent)"'
        $wb = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'WorkbenchWindow.xaml') -Raw
        $wb | Should -Not -Match 'x:Key="(DialogButton|DialogAccentButton|ThemedButtonStyle|ApButton|WbButton|WbAccentButton|WbInlineButton)"'
        @($script:MainDoc.SelectNodes("//*[local-name()='ControlTemplate'][@TargetType='Button']")).Count | Should -Be 1
    }

    It 'shows hover as a border, pressed as swapped theme colors and disabled as gray text' {
        $tpl = $script:MainDoc.SelectSingleNode("//*[local-name()='ControlTemplate'][@*[local-name()='Key']='ThemedButtonTemplate']")
        $hover = $tpl.SelectSingleNode(".//*[local-name()='Trigger'][@Property='IsMouseOver']/*[local-name()='Setter']")
        $hover.GetAttribute('Property') | Should -Be 'BorderBrush'
        $hover.GetAttribute('Value') | Should -Match 'ThemeForeground'
        $pressed = @($tpl.SelectNodes(".//*[local-name()='Trigger'][@Property='IsPressed']/*[local-name()='Setter']") | ForEach-Object { $_.GetAttribute('Value') }) -join ' '
        $pressed | Should -Match 'ThemeForeground'
        $pressed | Should -Match 'ThemeBackground'
        $st = $script:MainDoc.SelectSingleNode("//*[local-name()='Style'][@*[local-name()='Key']='ThemedButtonStyle']")
        $dis = @($st.SelectNodes(".//*[local-name()='Trigger'][@Property='IsEnabled']/*[local-name()='Setter'][@Property='Foreground']"))
        $dis.Count | Should -Be 1
        $dis[0].GetAttribute('Value') | Should -Match 'Gray\.MouseOver'
    }
}

Describe 'Known theme limits are documented (d)' {
    It 'documents each known limit with a reason and an owner, in the list and in MainWindow.xaml' {
        . "$PSScriptRoot\KnownThemeLimits.ps1"
        @($script:KnownThemeLimit).Count | Should -BeGreaterOrEqual 1
        foreach ($l in $script:KnownThemeLimit) {
            foreach ($field in 'Id', 'Where', 'Reason', 'Source') { [string]$l[$field] | Should -Not -BeNullOrEmpty -Because ('{0} needs {1}' -f $l.Id, $field) }
            ([string]$l.Reason).Length | Should -BeGreaterOrEqual 60
        }
        $main = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
        $main | Should -Match 'Known limit of the vendored SuiteCommon'
        $main | Should -Match 'KnownThemeLimits\.ps1'
        (Get-Content -LiteralPath $script:SuitePath -Raw) | Should -Match '\.Foreground\s*=\s*\$OnBlue'
    }

    It 'shows hover with a 2 px border, so that it differs from the rest state of every button style' {
        $tpl = $script:MainDoc.SelectSingleNode("//*[local-name()='ControlTemplate'][@*[local-name()='Key']='ThemedButtonTemplate']")
        $hover = @($tpl.SelectNodes(".//*[local-name()='Trigger'][@Property='IsMouseOver']/*[local-name()='Setter']") | ForEach-Object { '{0}={1}' -f $_.GetAttribute('Property'), $_.GetAttribute('Value') })
        ($hover -join ' ') | Should -Match 'BorderThickness=2'
    }
}
Describe 'Titles are in Title Case (e)' {
    It 'recognizes Title Case (the check itself)' {
        Test-TitleCase 'Stage Packages' | Should -BeTrue
        Test-TitleCase 'Publish to ConfigMgr' | Should -BeTrue
        Test-TitleCase 'Install from File...' | Should -BeTrue
        Test-TitleCase 'Stage packages' | Should -BeFalse
        Test-TitleCase 'Plan only' | Should -BeFalse
        Test-TitleCase 'Fixed by update' | Should -BeFalse
    }

    It 'has Title Case for every button, column header, window title and bold label in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-XamlTitle -Text $_.Text | Where-Object { -not (Test-TitleCase $_) }) | Should -BeNullOrEmpty
    }

    It 'finds titles to check (the scan is not empty)' {
        $count = 0
        foreach ($d in $script:XamlDocs) { $count += @(Get-XamlTitle -Text $d.Text).Count }
        $count | Should -BeGreaterThan 150
    }

    It 'has Title Case in the title of every message box, dialog and file dialog of the script' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
        $calls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in 'Show-ThemedMessage', 'Show-PreviewDialog', 'Show-WorkbenchNameDialog' }, $true))
        $calls.Count | Should -BeGreaterThan 20
        $bad = foreach ($c in $calls) {
            $el = @($c.CommandElements)
            for ($i = 0; $i -lt $el.Count - 1; $i++) {
                if ($el[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and $el[$i].ParameterName -eq 'Title' -and $el[$i + 1] -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    if (-not (Test-TitleCase ([string]$el[$i + 1].Value))) { [string]$el[$i + 1].Value }
                }
            }
        }
        @($bad) | Should -BeNullOrEmpty
        $assign = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and [string]$n.Left.Member.Value -eq 'Title' -and $n.Right.Extent.Text -match "^'" }, $true))
        $assign.Count | Should -BeGreaterThan 8
        @($assign | Where-Object { -not (Test-TitleCase ($_.Right.Extent.Text.Trim("'"))) } | ForEach-Object { $_.Right.Extent.Text }) | Should -BeNullOrEmpty
    }
}

Describe 'No raw exception text in the UI (f)' {
    It 'flags raw exception text that leaves the catch block (the check itself)' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-LogLine ("x {0}" -f $_.Exception.Message) }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $txt.Text = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $State.ErrorMsg = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { return @{ Success = $false; Error = $_.ToString() } }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $msg = $_.Exception.Message; Add-LogLine $msg }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Show-ThemedMessage -Message $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { throw $_ }').Count | Should -Be 1
    }

    It 'allows the log file, the helper, a condition and a deliberate swallow (the check itself)' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Write-Log ("x {0}" -f $_.Exception.Message) -Level ERROR; Add-LogLine "Failed. See the log file." }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-LogLine (Get-UserErrorText -ErrorRecord $_ -Fallback "Failed.") }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { if ($_.Exception.Message -match "x") { 1 } }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $null = $_ }').Count | Should -Be 0
        # Add-Content writes to a file, the log file of the window.
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-Content -LiteralPath $log -Value $_.Exception.Message }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $x = @(1) | ForEach-Object { $_ + 1 } }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { throw }').Count | Should -Be 0
    }

    It 'finds the catch blocks of the start script (the scan is not empty)' {
        (Get-CatchCount -Path $script:ShellPath) | Should -BeGreaterOrEqual 200
    }

    It 'sends no raw exception text to the UI from the start script' {
        @(Get-RawExceptionToUi -Path $script:ShellPath) | Should -BeNullOrEmpty
    }

    It 'defines Get-UserErrorText in the start script' {
        $script:ShellText | Should -Match 'function Get-UserErrorText'
    }
}

Describe 'Every test file requires Pester 5.1.0' {
    It 'has the #Requires line in <Name>' -ForEach @(Get-ChildItem -LiteralPath (Split-Path -Parent $PSScriptRoot) -Filter '*.Tests.ps1' -File -Recurse | Where-Object { $_.FullName.Substring((Split-Path -Parent $PSScriptRoot).Length).TrimStart('\', '/') -notmatch '^(\.claude|\.git)([\\/]|$)' } | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $req = @($ast.ScriptRequirements.RequiredModules | Where-Object { $_.Name -eq 'Pester' })
        $req.Count | Should -Be 1
        $req[0].Version | Should -Be ([version]'5.1.0')
    }

    It 'finds the test files (the scan is not empty)' {
        @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File -Recurse).Count | Should -BeGreaterOrEqual 60
    }
}
