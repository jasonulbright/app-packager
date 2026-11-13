# Helpers of FontWeight.Guard.Tests.ps1. Dot-source from BeforeDiscovery and from BeforeAll.
# They parse files. They create no WPF control and open no window.
#
# Rule: only a field title and an app title are bold. A field title is the text label of an input control.
# Every other label (page header, section header, status text, about text, log header) keeps the weight
# that it had before the theme change (Normal or SemiBold).

$script:FwRepoRoot = Split-Path -Parent $PSScriptRoot
$script:FwXamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
$script:FwKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
$script:FwExcluded = '^(\.claude|\.git|Tests|tools|Logs)([\\/]|$)'
$script:FwBoldWord = '(?i)\b(Bold|ExtraBold|UltraBold|Black|Heavy|ExtraBlack|UltraBlack)\b'
$script:FwInputType = @('TextBox', 'ComboBox', 'CheckBox', 'RadioButton', 'PasswordBox', 'DatePicker', 'DateTimePicker', 'TimePicker', 'ListBox', 'ListView', 'Slider', 'NumericUpDown')
$script:FwHolderType = @('Border', 'StackPanel', 'Grid', 'DockPanel', 'WrapPanel')

# ---- Settings of this repository ---------------------------------------------------------------------------------
# FieldTitleStyle : style keys that make a field title bold. A style key listed here may set FontWeight Bold.
# AppTitleName    : x:Name of an app title. It may be bold. The sidebar app title is the one case.
# KeptBoldName    : x:Name of a label that was bold before the theme change and stays so. Keep this list short.
# KeptBoldStyle   : key of a style that sets bold and was bold before the theme change (the selected tab header). Keep this list short.
# FieldTitleName  : x:Name of a field title whose input is not a listed input control (a label of a button row). Keep this list short.
# StaticName      : x:Name of a static header or status element. It must never be bold, even next to an input.
# AllowedText     : the text of a label that was bold before the theme change and stays so (not a field title). Keep this list short.
# ExemptType      : control types that are not labels. A button text is not a static label.
# CodeAllowed     : 'Function:Variable' of a run-time FontWeight assignment that may set bold ('*' for any variable).
$script:FwConfig = @{
    FieldTitleStyle = @('WbLabel')
    AppTitleName    = @('txtAppTitle')
    KeptBoldName    = @('txtDiscoverPrompt', 'txtCveTitle', 'txtPlanSummary')
    KeptBoldStyle   = @('WbHeading')
    FieldTitleName  = @()
    StaticName      = @('txtProgressTitle', 'txtStatus', 'txtThemeLabel', 'lblLogOutput')
    AllowedText     = @('AppPackager', 'Welcome to AppPackager', 'Console:', '7-Zip CLI:', 'GitHub API:', 'Content Prep:', 'Icon Pack:', 'Intunewin:', 'Not Offered:')
    ExemptType      = @('Button', 'ToggleButton', 'RepeatButton')
    CodeAllowed     = @('Set-ConnectionTextTheme:title', 'New-PackagerPreferencesPanel:tb', 'New-PackagerPreferencesPanel:lbl', 'New-ProductFilterPanel:vendorCB')
}
# ------------------------------------------------------------------------------------------------------------------

function Get-FwRelativePath {
    param([string]$FullName)
    return $FullName.Substring($script:FwRepoRoot.Length).TrimStart('\', '/')
}

# The script files of the app: the start script, the module and the vendored library. The tests and the tools are outside.
function Get-FwScriptFile {
    return @(Get-ChildItem -LiteralPath $script:FwRepoRoot -File -Recurse | Where-Object {
            $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' -and (Get-FwRelativePath $_.FullName) -notmatch $script:FwExcluded
        })
}

# The XAML here-strings of a script text, with their line range. A here-string counts as XAML when its first tag declares the WPF namespace.
function Get-FwHereString {
    param([string]$Text)
    $pattern = '(?s)@[''"]\r?\n(\s*<[^\r\n]*\r?\n(?:[^\r\n]*\r?\n)*?)[''"]@'
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        $body = $m.Groups[1].Value
        if ($body -notmatch ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:FwXamlNamespace) + '"')) { continue }
        $startLine = ($Text.Substring(0, $m.Index) -split "`n").Count
        $endLine = $startLine + ($m.Value -split "`n").Count - 1
        [PSCustomObject]@{ Text = $body; StartLine = $startLine; EndLine = $endLine }
    }
}

# Every XAML document of the app: the XAML files and the XAML here-strings of the scripts.
function Get-FwXamlDocument {
    $docs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:FwRepoRoot -Filter '*.xaml' -File -Recurse)) {
        $rel = Get-FwRelativePath $f.FullName
        if ($rel -match $script:FwExcluded) { continue }
        $docs += @{ Source = $rel; Text = (Get-Content -LiteralPath $f.FullName -Raw) }
    }
    foreach ($f in @(Get-FwScriptFile)) {
        $raw = Get-Content -LiteralPath $f.FullName -Raw
        $i = 0
        foreach ($h in @(Get-FwHereString -Text $raw)) {
            $i++
            $docs += @{ Source = ('{0} here-string {1}' -f (Get-FwRelativePath $f.FullName), $i); Text = $h.Text }
        }
    }
    return $docs
}

function Test-FwBoldValue {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match $script:FwBoldWord) -or ($Value -match '^\s*[7-9]\d\d\s*$')
}

function Get-FwElementName {
    param($Element)
    $n = $Element.GetAttribute('Name', $script:FwKeyNamespace)
    if (-not $n) { $n = $Element.GetAttribute('Name') }
    return $n
}

# True when a style, or a style that it is based on, sets a bold weight. Style keys of the settings are bold by definition.
function Test-FwStyleBold {
    param($Doc, [string]$Key)
    $guard = 0
    while ($Key -and $guard -lt 10) {
        $guard++
        if ($script:FwConfig.FieldTitleStyle -contains $Key) { return $true }
        $st = $Doc.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $Key))
        if (-not $st) { return $false }
        foreach ($s in @($st.SelectNodes("*[local-name()='Setter']"))) {
            if ($s.GetAttribute('Property') -match '(^|\.)FontWeight$' -and (Test-FwBoldValue $s.GetAttribute('Value'))) { return $true }
        }
        $based = $st.GetAttribute('BasedOn')
        $Key = if ($based -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { $Matches[1] } else { '' }
    }
    return $false
}

# True when the element has a bold weight of its own or through its style. A weight set on the element wins.
function Test-FwElementBold {
    param($Doc, $Element)
    foreach ($a in @($Element.Attributes)) {
        # A weight on the element wins over the weight of its style.
        if ($a.Name -match '(^|\.)FontWeight$') { return (Test-FwBoldValue $a.Value) }
    }
    $style = $Element.GetAttribute('Style')
    if ($style -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { return (Test-FwStyleBold -Doc $Doc -Key $Matches[1]) }
    return $false
}

function Test-FwExemptType {
    param([string]$Type)
    $t = ($Type -replace '^\{x:Type\s+', '' -replace '\}$', '').Trim()
    $t = $t -replace '^.*:', ''
    return ($script:FwConfig.ExemptType -contains $t)
}

# The text of a label: Text for a TextBlock, Content for a Label.
function Get-FwLabelText {
    param($Element)
    $t = $Element.GetAttribute('Text')
    if (-not $t) { $t = $Element.GetAttribute('Content') }
    return $t
}

# True when the element is the text label of an input control: it sits in the same grid row next to an input
# (or in the row above it), or it is the sibling before or after an input in a panel. A key label next to a
# TextBlock value is not a field title.
function Test-FwFieldTitle {
    param($Element)
    if ($Element.LocalName -notin 'TextBlock', 'Label') { return $false }
    $parent = $Element.ParentNode
    if ($null -eq $parent -or $parent.NodeType -ne 'Element') { return $false }
    $siblings = @($parent.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -notmatch '\.' })
    $hasInput = {
        param($n)
        if ($script:FwInputType -contains $n.LocalName) { return $true }
        if ($script:FwHolderType -contains $n.LocalName) {
            foreach ($c in @($n.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })) { if ($script:FwInputType -contains $c.LocalName) { return $true } }
        }
        return $false
    }
    if ($parent.LocalName -eq 'Grid') {
        $row = 0; $col = 0
        [void][int]::TryParse($Element.GetAttribute('Grid.Row'), [ref]$row)
        [void][int]::TryParse($Element.GetAttribute('Grid.Column'), [ref]$col)
        foreach ($s in $siblings) {
            if ($s -eq $Element -or -not (& $hasInput $s)) { continue }
            $sr = 0; $sc = 0
            [void][int]::TryParse($s.GetAttribute('Grid.Row'), [ref]$sr)
            [void][int]::TryParse($s.GetAttribute('Grid.Column'), [ref]$sc)
            if ($sr -eq $row -or ($sr -eq ($row + 1) -and $sc -eq $col)) { return $true }
        }
        return $false
    }
    for ($i = 0; $i -lt $siblings.Count; $i++) {
        if ($siblings[$i] -ne $Element) { continue }
        if ($i -gt 0 -and (& $hasInput $siblings[$i - 1])) { return $true }
        if ($i -lt ($siblings.Count - 1) -and (& $hasInput $siblings[$i + 1])) { return $true }
    }
    return $false
}

# One text per bold style and per bold element of a XAML document that is not a field title or an app title.
function Get-FwXamlViolation {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($s in @($x.SelectNodes("//*[local-name()='Setter'][@Property]"))) {
        if ($s.GetAttribute('Property') -notmatch '(^|\.)FontWeight$') { continue }
        $value = $s.GetAttribute('Value')
        if (-not $value) { $value = $s.InnerText }
        if (-not (Test-FwBoldValue $value)) { continue }
        $owner = $s.SelectSingleNode("ancestor::*[local-name()='Style' or local-name()='ControlTemplate'][1]")
        $key = ''; $target = ''
        if ($owner) { $key = $owner.GetAttribute('Key', $script:FwKeyNamespace); $target = $owner.GetAttribute('TargetType') }
        if ($target -and (Test-FwExemptType $target)) { continue }
        if ($key -and ($script:FwConfig.FieldTitleStyle -contains $key -or $script:FwConfig.KeptBoldStyle -contains $key)) { continue }
        [void]$found.Add(('{0}: style "{1}" (TargetType {2}) sets FontWeight Bold' -f $Source, $key, $target))
    }
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ($el.LocalName -eq 'Setter') { continue }
        if ($el.LocalName -match '\.FontWeight$') {
            if ((Test-FwBoldValue $el.InnerText) -and -not (Test-FwExemptType $el.ParentNode.LocalName)) {
                [void]$found.Add(('{0}: <{1}> sets FontWeight Bold' -f $Source, $el.LocalName))
            }
            continue
        }
        if (-not (Test-FwElementBold -Doc $x -Element $el)) { continue }
        if (Test-FwExemptType $el.LocalName) { continue }
        $ownWeight = $false
        foreach ($a in @($el.Attributes)) { if ($a.Name -match '(^|\.)FontWeight$' -and (Test-FwBoldValue $a.Value)) { $ownWeight = $true } }
        if (-not $ownWeight -and $el.GetAttribute('Style') -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)' -and $script:FwConfig.KeptBoldStyle -contains $Matches[1]) { continue }
        $name = Get-FwElementName $el
        $label = if ($name) { $name } else { Get-FwLabelText $el }
        if ($name -and ($script:FwConfig.AppTitleName -contains $name -or $script:FwConfig.KeptBoldName -contains $name)) { continue }
        if ($script:FwConfig.AllowedText -contains (Get-FwLabelText $el)) { continue }
        if ($name -and $script:FwConfig.StaticName -contains $name) {
            [void]$found.Add(('{0}: <{1}> {2} is a static header or status element and is bold' -f $Source, $el.LocalName, $name))
            continue
        }
        if (-not ($script:FwConfig.FieldTitleName -contains $name) -and -not (Test-FwFieldTitle $el)) {
            [void]$found.Add(('{0}: <{1}> "{2}" is bold and is not a field title' -f $Source, $el.LocalName, $label))
        }
    }
    return $found.ToArray()
}

# The number of text labels (TextBlock and Label) and of field titles in a XAML document.
function Get-FwLabelCount {
    param([string]$Text)
    [xml]$x = $Text
    $labels = @($x.SelectNodes("//*[local-name()='TextBlock' or local-name()='Label']"))
    $titles = @($labels | Where-Object { Test-FwFieldTitle $_ })
    return [PSCustomObject]@{ Labels = $labels.Count; Titles = $titles.Count }
}

# True when the named element of a XAML document exists, and whether it is bold.
function Get-FwNamedElementState {
    param([string]$Text, [string]$Name)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ((Get-FwElementName $el) -ne $Name) { continue }
        return [PSCustomObject]@{ Found = $true; Bold = (Test-FwElementBold -Doc $x -Element $el) }
    }
    return [PSCustomObject]@{ Found = $false; Bold = $false }
}

# One text per FontWeight Bold in script code (not in XAML here-strings) that is not a field title.
# A run-time assignment is a field title when the same scope sets the Text of the same variable to a text that ends in a colon,
# or when the settings list 'Function:Variable'. A line that sets a bold weight in a form that this scan does not parse is a violation too.
function Get-FwCodeViolation {
    param([string]$Path, [string]$ScriptText, [string]$Label)
    $cfg = $script:FwConfig
    $ast = if ($ScriptText) { [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$null, [ref]$null) } else { [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null) }
    $raw = if ($ScriptText) { $ScriptText } else { Get-Content -LiteralPath $Path -Raw }
    $found = New-Object System.Collections.Generic.List[string]
    $accounted = New-Object System.Collections.Generic.HashSet[int]
    $functionOf = {
        param($node)
        $p = $node
        while ($p -and $p -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $p = $p.Parent }
        return $p
    }
    $assignments = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and [string]$n.Left.Member.Value -eq 'FontWeight' }, $true))
    foreach ($a in $assignments) {
        [void]$accounted.Add($a.Extent.StartLineNumber)
        if (-not (Test-FwBoldValue $a.Right.Extent.Text) -and $a.Right.Extent.Text -notmatch 'FromOpenTypeWeight\(\s*[7-9]\d\d') { continue }
        $fn = & $functionOf $a
        $fnName = if ($fn) { $fn.Name } else { '' }
        $scope = if ($fn) { $fn } else { $ast }
        $var = $a.Left.Expression.Extent.Text.TrimStart('$')
        $isTitle = $false
        $texts = @($scope.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and [string]$n.Left.Member.Value -eq 'Text' -and $n.Left.Expression.Extent.Text.TrimStart('$') -eq $var }, $true))
        foreach ($t in $texts) { if ($t.Right.Extent.Text -match ':\s*[''"]\s*$') { $isTitle = $true } }
        if ($cfg.CodeAllowed -contains ('{0}:{1}' -f $fnName, $var) -or $cfg.CodeAllowed -contains ('{0}:*' -f $fnName)) { $isTitle = $true }
        if (-not $isTitle) {
            [void]$found.Add(('{0}: line {1}: ${2}.FontWeight is set to bold in {3} and ${2} is not a field title' -f $Label, $a.Extent.StartLineNumber, $var, $(if ($fnName) { $fnName } else { 'the script' })))
        }
    }
    $tables = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true))
    foreach ($h in $tables) {
        $weight = $null; $hasColonText = $false
        foreach ($pair in $h.KeyValuePairs) {
            $k = $pair.Item1.Extent.Text.Trim('''', '"')
            if ($k -eq 'FontWeight') { $weight = $pair.Item2 }
            if ($k -eq 'Text' -and $pair.Item2.Extent.Text -match ':\s*[''"]\s*$') { $hasColonText = $true }
        }
        if ($null -eq $weight) { continue }
        [void]$accounted.Add($weight.Extent.StartLineNumber)
        if (-not (Test-FwBoldValue $weight.Extent.Text)) { continue }
        if (-not $hasColonText) {
            [void]$found.Add(('{0}: line {1}: a property table sets FontWeight to bold and it is not a field title' -f $Label, $weight.Extent.StartLineNumber))
        }
    }
    # A call that asks for a bold label: the -Bold switch, or a bold word as the argument of a parameter (-Weight Bold).
    # The first positional text of the call is the label. It must end in a colon or be listed in AllowedText.
    $calls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))
    foreach ($c in $calls) {
        $asksBold = $false; $labelText = $null; $afterParameter = $false
        $elements = @($c.CommandElements | Select-Object -Skip 1)
        foreach ($e in $elements) {
            if ($e -is [System.Management.Automation.Language.CommandParameterAst]) {
                if ($e.ParameterName -eq 'Bold') { $asksBold = $true }
                if ($e.Argument -and (Test-FwBoldValue $e.Argument.Extent.Text)) { $asksBold = $true }
                $afterParameter = $true
                continue
            }
            if ($e -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $e -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
                if ($afterParameter -and $e.Value -match $script:FwBoldWord -and $e.Value -match '^\s*(Bold|ExtraBold|UltraBold|Black|Heavy|ExtraBlack|UltraBlack)\s*$') { $asksBold = $true }
                elseif (-not $afterParameter -and $null -eq $labelText) { $labelText = [string]$e.Value }
            }
            $afterParameter = $false
        }
        if (-not $asksBold) { continue }
        if ($labelText -and ($labelText -match ':\s*$' -or $cfg.AllowedText -contains $labelText)) { continue }
        [void]$found.Add(('{0}: line {1}: the call {2} asks for a bold label "{3}" that is not a field title' -f $Label, $c.Extent.StartLineNumber, $c.GetCommandName(), $labelText))
    }
    $xamlLine = New-Object System.Collections.Generic.HashSet[int]
    foreach ($h in @(Get-FwHereString -Text $raw)) { foreach ($n in $h.StartLine..$h.EndLine) { [void]$xamlLine.Add($n) } }
    $lines = $raw -split "`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $n = $i + 1
        if ($xamlLine.Contains($n) -or $accounted.Contains($n)) { continue }
        if ($lines[$i] -match '^\s*#') { continue }
        if ($lines[$i] -cmatch 'FontWeight' -and $lines[$i] -match $script:FwBoldWord) {
            [void]$found.Add(('{0}: line {1}: FontWeight is set to bold in a form that the scan does not parse' -f $Label, $n))
        }
    }
    return $found.ToArray()
}
