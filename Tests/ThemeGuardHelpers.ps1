# Helpers of Theme.Guard.Tests.ps1. Dot-source from BeforeDiscovery and from BeforeAll.
# They parse files. They create no WPF control and open no window.

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:ThemeProperty = @('FontSize', 'FontWeight', 'FontStyle', 'TextDecorations', 'FontFamily', 'FontStretch')
$script:BannedPattern = '(?i)\b(easily|simply|seamless(ly)?|powerful|robust|streamline[sd]?|streamlining|quickly|just|effortless(ly)?|intuitive(ly)?|handy|helps you|allows you to)\b'
$script:XamlKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
$script:XamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
$script:ExcludedFolder = '^(\.claude|\.git)([\\/]|$)'
$script:SharedStyleMarker = '<!--SHARED_STYLES-->'
$script:SizeAttribute = @('Height', 'Width', 'MinWidth', 'MinHeight', 'MaxWidth', 'MaxHeight', 'Padding')
# Hex colors that may stand outside a Style and a ControlTemplate. None is needed: every color is a theme resource.
$script:AllowedHexOutsideStyle = @()

# The path of a file relative to the repository root. A worktree lives under .claude, so a filter on the full path
# would match every file. A filter on the relative path does not.
function Get-RelativeRepoPath {
    param([string]$FullName)
    return $FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
}

function Get-PropertyLeaf {
    param([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return '' }
    return ($Name -split '[.:]')[-1]
}

# True when a brush value names an accent, a highlight or a blue color.
function Test-BlueBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -match 'MahApps\.(Brushes|Colors)\.(Accent\d*|AccentBase|Highlight\d*|Button\.Flat\.Background\.Pressed)\b') { return $true }
    if ($Value -match 'SystemColors\.Highlight(Brush|Color)?(Key)?\b') { return $true }
    if ($Value -match '(?i)^\s*(blue|dodgerblue|royalblue|steelblue|navy|mediumblue|darkblue|cornflowerblue|deepskyblue|lightskyblue|skyblue)\s*$') { return $true }
    $hex = $Value.Trim()
    if ($hex -match '^#([0-9A-Fa-f])([0-9A-Fa-f])([0-9A-Fa-f])$') { $hex = '#{0}{0}{1}{1}{2}{2}' -f $Matches[1], $Matches[2], $Matches[3] }
    if ($hex -match '^#(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0
        $g = [Convert]::ToInt32($Matches[2], 16) / 255.0
        $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $max = [Math]::Max($r, [Math]::Max($g, $b)); $min = [Math]::Min($r, [Math]::Min($g, $b))
        $d = $max - $min
        if ($d -lt 0.2 -or $max -eq 0) { return $false }
        if ($max -ne $b) { return $false }
        $hue = 60.0 * (4.0 + (($r - $g) / $d))
        return ($hue -ge 190 -and $hue -le 260)
    }
    return $false
}

function Test-WhiteBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)^\s*(white|#FFFFFF|#FFFFFFFF|#FFF)\s*$' -or $Value -match 'MahApps\.Brushes\.(Ideal|ThemeIdeal)Foreground')
}

# The value of a Setter: the Value attribute, or the Color of a nested brush at any depth.
function Get-SetterValue {
    param($Setter)
    if ($Setter.HasAttribute('Value')) { return $Setter.GetAttribute('Value') }
    $node = $Setter.SelectSingleNode('.//*[@Color]')
    if ($node) { return $node.GetAttribute('Color') }
    return ''
}

function Get-SetterChild {
    param($Node, [string]$LeafPattern)
    if (-not $Node) { return @() }
    return @($Node.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and (Get-PropertyLeaf $_.GetAttribute('Property')) -match $LeafPattern })
}

function Get-ResourceKey {
    param([string]$Value)
    if ($Value -match '^\{(?:Static|Dynamic)Resource\s+([^}\s]+)\}$') { return $Matches[1] }
    return ''
}

# The styles that dialogs and the Workbench take from MainWindow.xaml (the start script holds the same list).
function Get-SharedStyleXaml {
    [xml]$main = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
    $keys = @('ThemedButtonTemplate', 'ThemedButtonStyle', 'DialogAccentButton', 'DialogButton', 'InlineButtonStyle', 'LinkButtonStyle', 'GridLinkButtonStyle', 'ThemedListBoxItemStyle', 'ThemedComboBoxItemStyle', 'ThemedTreeViewItemStyle', 'GridRowStyle', 'GridCellStyle', 'PaddedGridCellStyle')
    $types = @('ListBoxItem', 'ComboBoxItem', 'TreeViewItem', 'CheckBox')
    $nodes = @($main.SelectNodes("//*[local-name()='ResourceDictionary']/*[local-name()='ControlTemplate' or local-name()='Style']") | Where-Object {
            $key = $_.GetAttribute('Key', $script:XamlKeyNamespace)
            ($key -and ($keys -contains $key)) -or (-not $key -and ($types -contains $_.GetAttribute('TargetType')))
        })
    return (($nodes | ForEach-Object { $_.OuterXml }) -join [Environment]::NewLine)
}

# The script files of the app. The tests, the tools, the samples and the vendored library are outside.
function Get-AppScriptFile {
    return @(Get-ChildItem -LiteralPath $script:RepoRoot -File -Recurse | Where-Object {
            $rel = Get-RelativeRepoPath $_.FullName
            $_.Extension -in '.ps1', '.psm1' -and $rel -notmatch '^(\.claude|\.git|Tests|tools|Lib|Logs|Samples|Packagers|VersionMonitor)([\\/]|$)'
        })
}

# Every XAML document of the app: the XAML files and the XAML strings of the start script (window and panel).
# The marker of the shared styles is replaced, as the start script does at run time.
function Get-XamlDocument {
    $docs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.xaml' -File -Recurse | Where-Object { (Get-RelativeRepoPath $_.FullName) -notmatch $script:ExcludedFolder })) {
        $text = (Get-Content -LiteralPath $f.FullName -Raw).Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
        $docs += [PSCustomObject]@{ Source = $f.Name; Text = $text; Kind = 'file' }
    }
    $path = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
    $nodes = @($ast.FindAll({ param($n) ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match 'xmlns="http://schemas\.microsoft\.com/winfx/2006/xaml/presentation"' }, $true))
    foreach ($n in $nodes) {
        $text = ([string]$n.Value).Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
        $kind = if ($text -match '^\s*<Controls:MetroWindow') { 'window' } else { 'panel' }
        $docs += [PSCustomObject]@{ Source = ('start-apppackager.ps1:{0}' -f $n.Extent.StartLineNumber); Text = $text; Kind = $kind }
    }
    return $docs
}

# The style node for a key, from the document, or from the shared styles of MainWindow.xaml for a panel.
function Find-StyleNode {
    param($Doc, [string]$Key)
    $q = "//*[local-name()='Style' or local-name()='ControlTemplate'][@*[local-name()='Key']='{0}']" -f $Key
    $n = $Doc.SelectSingleNode($q)
    if (-not $n -and $script:MainDoc) { $n = $script:MainDoc.SelectSingleNode($q) }
    return $n
}

function Get-StyleChain {
    param($Doc, [string]$Key)
    $chain = @(); $guard = 0
    while ($Key -and $guard -lt 10) {
        $guard++
        $chain += $Key
        $st = Find-StyleNode -Doc $Doc -Key $Key
        if (-not $st) { break }
        $Key = Get-ResourceKey $st.GetAttribute('BasedOn')
    }
    return $chain
}

# The style chain of a Button: the Style attribute, or the style inside the Button.
function Get-ButtonStyleChain {
    param($Doc, $Button)
    $key = Get-ResourceKey $Button.GetAttribute('Style')
    if ($key) { return @(Get-StyleChain -Doc $Doc -Key $key) }
    return @()
}

# Returns one text per blue-background rule that has no white foreground (Setter, nested brush, attribute, property element).
function Get-BlueWithoutWhite {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ($el.LocalName -eq 'Setter') {
            if ((Get-PropertyLeaf $el.GetAttribute('Property')) -match 'Background$') {
                $v = Get-SetterValue $el
                if (Test-BlueBrush $v) {
                    $scope = @($el.ParentNode)
                    $st = $el.ParentNode
                    while ($st -and $st.LocalName -notin 'Style', 'ControlTemplate') { $st = $st.ParentNode }
                    if ($st) { $scope += $st }
                    $white = $false
                    foreach ($s in $scope) { foreach ($fg in @(Get-SetterChild -Node $s -LeafPattern 'Foreground$')) { if (Test-WhiteBrush (Get-SetterValue $fg)) { $white = $true } } }
                    if (-not $white) { [void]$found.Add(('{0}: Setter {1} {2} in <{3}> has no white Foreground' -f $Source, $el.GetAttribute('Property'), $v, $el.ParentNode.LocalName)) }
                }
            }
            continue
        }
        if ($el.LocalName -match '\.\w*Background$') {
            $nested = $el.SelectSingleNode('.//*[@Color]')
            if ($nested -and (Test-BlueBrush $nested.GetAttribute('Color')) -and $el.ParentNode -and -not (Test-WhiteBrush $el.ParentNode.GetAttribute('Foreground')) -and -not (Test-WhiteBrush $el.ParentNode.GetAttribute('TextElement.Foreground'))) {
                [void]$found.Add(('{0}: <{1}> has a blue brush and its owner has no white Foreground' -f $Source, $el.LocalName))
            }
            continue
        }
        foreach ($a in @($el.Attributes | Where-Object { (Get-PropertyLeaf $_.Name) -match 'Background$' })) {
            if (-not (Test-BlueBrush $a.Value)) { continue }
            $own = $false
            foreach ($f in @($el.Attributes | Where-Object { (Get-PropertyLeaf $_.Name) -match 'Foreground$' })) { if (Test-WhiteBrush $f.Value) { $own = $true } }
            if ($own) { continue }
            [void]$found.Add(('{0}: <{1}> {2} {3} has no white Foreground' -f $Source, $el.LocalName, $a.Name, $a.Value))
        }
    }
    return $found.ToArray()
}

# Lists every color literal in the Style and ControlTemplate nodes (Setters, attributes, nested brushes). Only a theme
# resource (DynamicResource, StaticResource, TemplateBinding), a binding or x:Null is allowed. Any bare word and any hex is a literal.
function Get-HardColor {
    param($Nodes, [string]$Label)
    $brushProp = '(^|[.:])\w*(Brush|Background|Foreground|Fill|Stroke)$'
    $allowed = '^\{(DynamicResource|StaticResource|TemplateBinding|Binding|x:Null)\b'
    foreach ($n in $Nodes) {
        foreach ($setter in @($n.SelectNodes('descendant-or-self::*[local-name()="Setter"]'))) {
            if ($setter.GetAttribute('Property') -notmatch $brushProp) { continue }
            $v = $setter.GetAttribute('Value')
            if ($v -and $v -notmatch $allowed) { '{0}: Setter {1}="{2}"' -f $Label, $setter.GetAttribute('Property'), $v }
        }
        foreach ($el in @($n.SelectNodes('descendant-or-self::*[not(local-name()="Setter")]'))) {
            foreach ($attr in @($el.Attributes)) {
                if ($attr.Name -match '^xmlns' -or $attr.NamespaceURI -eq 'http://www.w3.org/2000/xmlns/') { continue }
                if ($attr.LocalName -match $brushProp -and $attr.Value -notmatch $allowed) { '{0}: {1} {2}="{3}"' -f $Label, $el.LocalName, $attr.LocalName, $attr.Value }
                elseif ($attr.LocalName -eq 'Color' -and $attr.Value -notmatch $allowed) { '{0}: {1} holds the color {2}' -f $Label, $el.LocalName, $attr.Value }
            }
        }
    }
}

# Color literals outside the styles: a hex value or a bare color word in a brush attribute of a control, unless it is on the allow list.
function Get-LiteralColorOutsideStyle {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*[not(ancestor-or-self::*[local-name()="Style" or local-name()="ControlTemplate"])]'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.Name -match '^xmlns' -or $a.NamespaceURI -eq 'http://www.w3.org/2000/xmlns/') { continue }
            if ((Get-PropertyLeaf $a.Name) -notmatch '(Brush|Background|Foreground|Fill|Stroke|Color)$') { continue }
            if ($a.Value -match '^\s*#[0-9A-Fa-f]{3,8}\s*$' -and $script:AllowedHexOutsideStyle -notcontains $a.Value.Trim()) { '{0}: <{1}> {2}="{3}"' -f $Source, $el.LocalName, $a.Name, $a.Value }
            elseif ($a.Value -match '^\s*[A-Za-z]+\s*$' -and $a.LocalName -ne 'Color' -and $a.LocalName -ne 'LastChildFill') { '{0}: <{1}> {2}="{3}"' -f $Source, $el.LocalName, $a.Name, $a.Value }
        }
    }
}

# One text per state trigger that sets the opacity, and one per Opacity attribute on a control with a blue or text role.
function Get-OpacityInStateTrigger {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($trigger in @($x.SelectNodes("//*[local-name()='Trigger' or local-name()='DataTrigger']"))) {
        if ($trigger.LocalName -eq 'Trigger' -and $trigger.GetAttribute('Property') -notin 'IsMouseOver', 'IsPressed', 'IsEnabled', 'IsKeyboardFocused', 'IsSelected', 'IsHighlighted') { continue }
        foreach ($s in @($trigger.SelectNodes("*[local-name()='Setter'][@Property='Opacity']"))) { 'Trigger {0} sets Opacity {1}' -f $trigger.GetAttribute('Property'), $s.GetAttribute('Value') }
    }
    foreach ($el in @($x.SelectNodes("//*[local-name()='TextBlock' or local-name()='Button'][@Opacity]"))) { '<{0}> sets Opacity {1}' -f $el.LocalName, $el.GetAttribute('Opacity') }
}

# One text per button problem: no style, a style that does not end at ThemedButtonStyle, a local size attribute,
# a literal color attribute, a template of its own.
function Get-ButtonViolation {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
        $name = $b.GetAttribute('Name', $script:XamlKeyNamespace)
        $chain = @(Get-ButtonStyleChain -Doc $x -Button $b)
        if ($chain -notcontains 'ThemedButtonStyle') { '{0}: Button {1} does not take a style that is based on ThemedButtonStyle' -f $Source, $name }
        foreach ($a in @($b.Attributes)) {
            if ($script:SizeAttribute -contains $a.LocalName) { '{0}: Button {1} sets its own {2}' -f $Source, $name, $a.LocalName }
            if ((Get-PropertyLeaf $a.Name) -match '(Background|Foreground|BorderBrush)$' -and $a.Value -match '^\s*(#[0-9A-Fa-f]{3,8}|[A-Za-z]+)\s*$') { '{0}: Button {1} sets {2} to the literal {3}' -f $Source, $name, $a.Name, $a.Value }
        }
    }
    foreach ($st in @($x.SelectNodes("//*[local-name()='Style'][@TargetType='Button'][@BasedOn]"))) {
        $key = $st.GetAttribute('Key', $script:XamlKeyNamespace)
        if ($key -ne 'ThemedButtonStyle' -and $st.GetAttribute('BasedOn') -match 'MahApps\.') { '{0}: Button style {1} is based on MahApps and not on ThemedButtonStyle' -f $Source, $key }
    }
}

# Own size attributes on the buttons that a script creates: $btn.Height, Width, MinWidth, Padding.
function Get-CodeButtonSize {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $names = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Right.Extent.Text -match 'New-Object\s+(-TypeName\s+)?System\.Windows\.Controls\.Button\b' }, $true) | ForEach-Object { $_.Left.VariablePath.UserPath })
    foreach ($a in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))) {
        if ($names -contains $a.Left.Expression.VariablePath.UserPath -and $script:SizeAttribute -contains [string]$a.Left.Member.Value) { 'line {0}: {1}' -f $a.Extent.StartLineNumber, $a.Extent.Text }
    }
}

# The controls of one type that a script creates with New-Object, with the variable name and the function.
function Get-CodeControl {
    param([string]$Path, [string]$TypeName)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $pattern = 'New-Object\s+(-TypeName\s+)?System\.Windows\.Controls\.' + [regex]::Escape($TypeName) + '\b'
    foreach ($a in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Right.Extent.Text -match $pattern }, $true))) {
        $fn = $a.Parent
        while ($fn -and $fn -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
        if (-not $fn) { $fn = $ast }
        [PSCustomObject]@{ Variable = $a.Left.VariablePath.UserPath; Line = $a.Extent.StartLineNumber; Function = $fn }
    }
}

function Get-CodeStyleKey {
    param($Function, [string]$Variable)
    $calls = @($Function.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$n.Member.Value -eq 'SetResourceReference' -and $n.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Expression.VariablePath.UserPath -eq $Variable -and $n.Arguments.Count -eq 2 -and $n.Arguments[0].Extent.Text -match 'StyleProperty' }, $true))
    if ($calls.Count -eq 0) { return '' }
    return [string]$calls[0].Arguments[1].Value
}

function Test-StateTrigger {
    param($Style, [string]$Property)
    foreach ($t in @($Style.SelectNodes(".//*[local-name()='Trigger'][@Property='$Property']"))) {
        $bg = Get-SetterChild -Node $t -LeafPattern '^Background$' | Select-Object -First 1
        $fg = Get-SetterChild -Node $t -LeafPattern '^Foreground$' | Select-Object -First 1
        if ($bg -and (Test-BlueBrush (Get-SetterValue $bg)) -and $fg -and (Test-WhiteBrush (Get-SetterValue $fg))) { return $true }
    }
    return $false
}

# One text per list, combo box, tree or grid whose item container, row or cell style lacks a blue Background with white Foreground
# for IsSelected (and IsHighlighted for a combo box item). The style comes from the Style attribute chain or the item container attribute.
function Get-SelectedStateViolation {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $rules = @(
        @{ Control = 'ListBox'; Attribute = 'ItemContainerStyle'; Property = @('IsSelected') }
        @{ Control = 'ComboBox'; Attribute = 'ItemContainerStyle'; Property = @('IsSelected', 'IsHighlighted') }
        @{ Control = 'TreeView'; Attribute = 'ItemContainerStyle'; Property = @('IsSelected') }
    )
    foreach ($r in $rules) {
        foreach ($el in @($x.SelectNodes(("//*[local-name()='{0}']" -f $r.Control)))) {
            $name = $el.GetAttribute('Name', $script:XamlKeyNamespace)
            $key = Get-ResourceKey $el.GetAttribute($r.Attribute)
            if (-not $key) {
                # The control style may set the item container style (Wb styles do).
                $ctl = Find-StyleNode -Doc $x -Key (Get-ResourceKey $el.GetAttribute('Style'))
                $set = if ($ctl) { Get-SetterChild -Node $ctl -LeafPattern ('^' + $r.Attribute + '$') | Select-Object -First 1 }
                if ($set) { $key = Get-ResourceKey $set.GetAttribute('Value') }
                elseif ($el.SelectSingleNode(("*[local-name()='{0}.{1}']/*[local-name()='Style']" -f $r.Control, $r.Attribute))) {
                    $inline = $el.SelectSingleNode(("*[local-name()='{0}.{1}']/*[local-name()='Style']" -f $r.Control, $r.Attribute))
                    $key = Get-ResourceKey $inline.GetAttribute('BasedOn')
                }
            }
            $st = if ($key) { Find-StyleNode -Doc $x -Key $key }
            if (-not $st) { '{0}: {1} {2} has no {3}' -f $Source, $r.Control, $name, $r.Attribute; continue }
            foreach ($p in $r.Property) { if (-not (Test-StateTrigger -Style $st -Property $p)) { '{0}: style {1} of {2} {3} has no blue Background with white Foreground for {4}' -f $Source, $key, $r.Control, $name, $p } }
        }
    }
    foreach ($el in @($x.SelectNodes("//*[local-name()='DataGrid']"))) {
        $name = $el.GetAttribute('Name', $script:XamlKeyNamespace)
        foreach ($pair in @(@('RowStyle', 'row'), @('CellStyle', 'cell'))) {
            $key = Get-ResourceKey $el.GetAttribute($pair[0])
            $st = if ($key) { Find-StyleNode -Doc $x -Key $key }
            $ok = $false
            if ($st) {
                foreach ($k in (Get-StyleChain -Doc $x -Key $key)) { $s2 = Find-StyleNode -Doc $x -Key $k; if ($s2 -and (Test-StateTrigger -Style $s2 -Property 'IsSelected')) { $ok = $true } }
            }
            if (-not $ok) { '{0}: DataGrid {1} has no {2} style with a blue Background and a white Foreground for IsSelected' -f $Source, $name, $pair[1] }
        }
    }
}

function Get-ThemeFontTrigger {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($trigger in @($x.SelectNodes("//*[local-name()='Trigger' or local-name()='DataTrigger' or local-name()='MultiTrigger' or local-name()='MultiDataTrigger']"))) {
        $sets = @($trigger.SelectNodes("*[local-name()='Setter']") | Where-Object { $script:ThemeProperty -contains (Get-PropertyLeaf $_.GetAttribute('Property')) })
        if ($sets.Count -eq 0) { continue }
        $condition = (@($trigger.Attributes | ForEach-Object { $_.Value }) + @($trigger.SelectNodes(".//*[local-name()='Condition']") | ForEach-Object { @($_.Attributes | ForEach-Object { $_.Value }) })) -join ' '
        if ($condition -match '(?i)IsDark|DarkMode|IsLight|LightMode|ThemeMode|\bTheme\b|\bDark\b') { $trigger.OuterXml.Substring(0, [Math]::Min(100, $trigger.OuterXml.Length)) }
    }
}

function Get-ScriptString {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    foreach ($n in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true))) {
        $p = $n.Parent; $inAttribute = $false
        while ($p) { if ($p -is [System.Management.Automation.Language.AttributeAst]) { $inAttribute = $true; break }; $p = $p.Parent }
        if (-not $inAttribute) { [string]$n.Value }
    }
}

function Get-XamlUiText {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.LocalName -match '^(ToolTip|Text|Content|Header|Title|Watermark|OnContent|OffContent|ToolTipService.*)$' -or $a.Name -match 'Watermark$') {
                if ($a.Value -notmatch '^\{') { $a.Value }
            }
        }
    }
}

function Get-XamlToolTip {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*[@ToolTip]'))) { $v = $el.GetAttribute('ToolTip'); if ($v -notmatch '^\{') { $v } }
}

# One short sentence: a capital letter first, one full stop at the end, 100 characters or less.
function Test-OneSentence {
    param([string]$Text)
    if ($Text.Length -gt 100) { return $false }
    if ($Text -cnotmatch '^[A-Z]') { return $false }
    if ($Text -notmatch '\.$') { return $false }
    return ($Text.TrimEnd('.') -notmatch '\.\s')
}

function Test-TitleCase {
    param([string]$Text)
    $small = 'a', 'an', 'and', 'as', 'at', 'by', 'for', 'from', 'in', 'of', 'on', 'or', 'the', 'to', 'with'
    $clean = ($Text -replace '\.\.\.$', '' -replace ':$', '').Trim()
    $words = @($clean -split '\s+' | Where-Object { $_ })
    if ($words.Count -eq 0) { return $true }
    for ($i = 0; $i -lt $words.Count; $i++) {
        $w = $words[$i].Trim(',', ':', '(', ')')
        if (-not $w) { continue }
        $edge = ($i -eq 0 -or $i -eq $words.Count - 1)
        if (-not $edge -and ($small -contains $w)) { continue }
        if ($w -cnotmatch '^[A-Z0-9#<\[\(]') { return $false }
    }
    return $true
}

# The titles of a XAML document that must be in Title Case: button content, column headers, window titles, label TextBlocks that end with a colon.
function Get-XamlTitle {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        $v = ''
        switch -Regex ($el.LocalName) {
            '^Button$' { $v = $el.GetAttribute('Content') }
            '^DataGrid\w*Column$' { $v = $el.GetAttribute('Header') }
            '^(MetroWindow|Window)$' { $v = $el.GetAttribute('Title') }
            '^TextBlock$' { if ($el.GetAttribute('Text') -match ':$' -and $el.GetAttribute('FontWeight') -eq 'Bold') { $v = $el.GetAttribute('Text') } }
        }
        if ($v -and $v -notmatch '^[{$]') { $v }
    }
}

# One text per catch block statement that sends exception text anywhere but to the log file.
function Get-RawExceptionToUi {
    param([string]$Path, [string]$ScriptText)
    $ast = if ($ScriptText) { [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$null, [ref]$null) } else { [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null) }
    $errorName = @('_', 'PSItem', 'Error', 'ex', 'err')
    $safeCommand = @('Write-Log', 'Write-LogErrorRecord', 'Get-UserErrorText', 'Get-HttpStatusCode', 'Write-Warning', 'Write-Verbose')
    $found = @()
    foreach ($c in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true))) {
        foreach ($v in @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and $errorName -contains $n.VariablePath.UserPath }, $true))) {
            $allowed = $false; $skip = $false; $q = $v; $statement = $null
            while ($q -and $q -ne $c) {
                if ($q -is [System.Management.Automation.Language.CatchClauseAst]) { $skip = $true; break }
                if ($q -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and @('_', 'PSItem') -contains $v.VariablePath.UserPath) { $skip = $true; break }
                if ($q -is [System.Management.Automation.Language.CommandAst] -and $safeCommand -contains [string]$q.GetCommandName()) { $allowed = $true; break }
                if ($q -is [System.Management.Automation.Language.AssignmentStatementAst] -and $q.Left.Extent.Text -eq '$null') { $allowed = $true; break }
                if ($q.Parent -is [System.Management.Automation.Language.IfStatementAst]) {
                    foreach ($cl in $q.Parent.Clauses) { if ($cl.Item1 -eq $q) { $allowed = $true } }
                    if ($allowed) { break }
                }
                if ($q -is [System.Management.Automation.Language.StatementAst] -and -not $statement) { $statement = $q }
                $q = $q.Parent
            }
            if ($skip -or $allowed) { continue }
            $where = if ($statement) { $statement } else { $v }
            $text = $where.Extent.Text
            if ($text.Length -gt 120) { $text = $text.Substring(0, 120) }
            $found += ('line {0}: {1}' -f $where.Extent.StartLineNumber, $text)
        }
    }
    return @($found | Select-Object -Unique)
}

function Get-CatchCount {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    return @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true)).Count
}
