<#
.SYNOPSIS
    ConfigMgr lexicon checker engine: finds user-facing text in a repository and matches it against
    configmgr-lexicon.json.

.DESCRIPTION
    Windows PowerShell 5.1 and PowerShell 7. Plain ASCII. No ConfigMgr site, module or network needed.

    Get-LexiconTextItem  reads the user-facing text of one repository: README.md, docs\**\*.md, the
                         [Unreleased] section of CHANGELOG.md, XAML text attributes, message box, dialog,
                         status, log, warning and error strings and comment-based help in .ps1/.psm1, and
                         for a C# repository the XAML files and sentence-like string literals in .cs files.
                         It skips code identifiers, cmdlet and class names, URLs, code blocks and inline code.
    Find-LexiconHit      matches the items against the flagged phrases of the lexicon and the allow-list.
    Invoke-LexiconCheck  both steps in one call. Returns one object per hit.

    The lexicon schema is in lexicon\README.md.
#>

$script:SkipDirPattern = '(^|[\\/])(\.git|\.github|\.claude|node_modules|TestResults|bin|obj|Lib|Tests|tests|_cand|_release|installer|lexicon|assets|images)([\\/]|$)'

function Import-LexiconData {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($e in @($json.entries)) {
        $compiled = New-Object System.Collections.Generic.List[object]
        foreach ($p in @($e.flagged)) {
            if ([string]::IsNullOrEmpty($p)) { continue }
            $compiled.Add((New-Object System.Text.RegularExpressions.Regex($p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)))
        }
        $allow = New-Object System.Collections.Generic.List[object]
        foreach ($p in @($e.allowWhen)) {
            if ([string]::IsNullOrEmpty($p)) { continue }
            $allow.Add((New-Object System.Text.RegularExpressions.Regex($p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)))
        }
        $e | Add-Member -NotePropertyName FlaggedRegex -NotePropertyValue $compiled.ToArray() -Force
        $e | Add-Member -NotePropertyName AllowRegex -NotePropertyValue $allow.ToArray() -Force
    }
    $json
}

function Import-LexiconAllowList {
    [CmdletBinding()]
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return @() }
    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    @($json.allow)
}

# --- text clean-up ------------------------------------------------------------------------------

function ConvertTo-ProseText {
    # Removes what is not prose from one markdown or comment line: inline code, URLs, images, HTML tags.
    param([string]$Line)
    $t = $Line
    $t = [regex]::Replace($t, '!\[[^\]]*\]\([^)]*\)', ' ')
    $t = [regex]::Replace($t, '\[([^\]]*)\]\([^)]*\)', '$1')
    $t = [regex]::Replace($t, '\[([^\]]*)\]\[[^\]]*\]', '$1')
    $t = [regex]::Replace($t, '`[^`]*`', ' ')
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    $t = [regex]::Replace($t, '(https?|file)://\S+', ' ')
    $t = [regex]::Replace($t, '\b[\w.-]+@[\w.-]+\.\w+\b', ' ')
    $t
}

function ConvertTo-FormatFreeText {
    # Replaces PowerShell and C# interpolation holes and format items with a neutral token.
    param([string]$Text)
    $t = $Text
    $t = [regex]::Replace($t, '\$\([^)]*\)', 'X')
    $t = [regex]::Replace($t, '\$\{[^}]*\}', 'X')
    $t = [regex]::Replace($t, '\$[A-Za-z_][\w:]*(\.\w+)*', 'X')
    $t = [regex]::Replace($t, '\{\d+(,-?\d+)?(:[^}]*)?\}', 'X')
    $t = [regex]::Replace($t, '\{[A-Za-z_][\w.]*(:[^}]*)?\}', 'X')
    $t = $t.Replace('`n', ' ').Replace('`r', ' ').Replace('\n', ' ').Replace('\r', ' ').Replace('`t', ' ')
    $t
}

function Test-ProseCandidate {
    param([string]$Text, [bool]$Strict)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    if ($Text -notmatch '[A-Za-z]{3}') { return $false }
    $trim = $Text.Trim()
    if ($trim -notmatch '\s') {
        # one token: a path, a key, a name or an identifier is not prose
        if ($trim -match '[\\/:_$%{}\[\]=@#]') { return $false }
        if ($trim -cmatch '^[a-z][a-z0-9]*[A-Z]') { return $false }
        if ($trim -match '^(SMS_|v_|vSMS|fn_|Get-|Set-|New-|Remove-|Invoke-|Add-)') { return $false }
    }
    if ($Strict) {
        # sentence-like: at least three words, starts with a capital letter, no query or pattern syntax
        if (([regex]::Matches($trim, '\s+')).Count -lt 2) { return $false }
        if ($trim -cnotmatch '^[A-Z]') { return $false }
        if ($trim -match '^(SELECT|INSERT|UPDATE|DELETE|WITH|FROM|WHERE)\s') { return $false }
        if ($trim -match '\\[dswDSW\\.]|\(\?[:=!<]|\[\^?[\w\\-]+\]\+?') { return $false }
    }
    $true
}

# --- markdown -----------------------------------------------------------------------------------

function Get-MarkdownTextItem {
    param([string]$Path, [string]$Rel, [switch]$UnreleasedOnly)
    $lines = [IO.File]::ReadAllLines($Path)
    $inFence = $false
    $inHtmlComment = $false
    $inFront = $false
    $section = $true
    if ($UnreleasedOnly) { $section = $false }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $raw = $lines[$i]
        if ($i -eq 0 -and $raw -match '^---\s*$') { $inFront = $true; continue }
        if ($inFront) { if ($raw -match '^---\s*$') { $inFront = $false }; continue }
        if ($UnreleasedOnly) {
            if ($raw -match '^##\s+\[?Unreleased\]?') { $section = $true; continue }
            if ($section -and $raw -match '^##\s+') { $section = $false }
            if (-not $section) { continue }
        }
        if ($raw -match '^\s*(```|~~~)') { $inFence = -not $inFence; continue }
        if ($inFence) { continue }
        if ($inHtmlComment) { if ($raw -match '-->') { $inHtmlComment = $false }; continue }
        if ($raw -match '^\s*<!--' -and $raw -notmatch '-->') { $inHtmlComment = $true; continue }
        if ($raw -match '^\s*<!--.*-->\s*$') { continue }
        if ($raw -match '^\s{4,}\S' -and $raw -notmatch '^\s*([-*+]|\d+\.)\s') { continue }   # indented code block
        if ($raw -match '^\s*\[[^\]]+\]:\s*\S+\s*$') { continue }                                # link reference definition
        $text = ConvertTo-ProseText -Line $raw
        $text = [regex]::Replace($text, '^\s*(#{1,6}|[-*+]|\d+\.|>)\s*', '')
        $text = $text -replace '[|*_~]', ' '
        if (-not (Test-ProseCandidate -Text $text -Strict:$false)) { continue }
        [pscustomobject]@{ File = $Rel; Line = $i + 1; Kind = 'markdown'; Text = $text.Trim() }
    }
}

# --- XAML ---------------------------------------------------------------------------------------

$script:XamlAttrRegex = New-Object System.Text.RegularExpressions.Regex(
    '(?<![\w:])(?:[A-Za-z]+\.)?(?<n>Content|Header|Text|ToolTip|Title|Watermark|PlaceholderText|PlaceholderHeader|Label)\s*=\s*"(?<v>[^"]*)"|AutomationProperties\.(?<n>Name|HelpText)\s*=\s*"(?<v>[^"]*)"')

function Get-XamlTextItemFromString {
    param([string]$Content, [string]$Rel, [int]$LineOffset, [string]$Kind)
    $lines = $Content -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        foreach ($m in $script:XamlAttrRegex.Matches($lines[$i])) {
            $v = $m.Groups['v'].Value
            if ($v.StartsWith('{}')) { $v = $v.Substring(2) }
            elseif ($v.StartsWith('{')) { continue }
            $v = [System.Net.WebUtility]::HtmlDecode($v)
            $v = $v -replace '\s+', ' '
            if (-not (Test-ProseCandidate -Text $v -Strict:$false)) { continue }
            [pscustomobject]@{ File = $Rel; Line = $LineOffset + $i; Kind = $Kind; Text = $v.Trim() }
        }
    }
}

# --- PowerShell ---------------------------------------------------------------------------------

$script:MessageCommandPattern = '^(Show-\w*(Message|Confirm|Warning|Error|Dialog)\w*|Set-Status\w*|Add-Log\w*|Write-(Warning|Error|Host)|Write-Log\w*|Write-Status\w*)$'
$script:TextMemberPattern = '^(Text|Content|Header|ToolTip|Title|Watermark|Description|Message|StatusText|Caption)$'

function Get-StringNodeText {
    # Yields the prose strings under one AST node.
    param($Node)
    $strings = $Node.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
            $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
        }, $true)
    if ($Node -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
        $Node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { $strings = @($Node) + @($strings) }
    foreach ($s in $strings) {
        $p = $s.Parent
        # skip member names, hashtable keys, index keys, parameter names
        if ($p -is [System.Management.Automation.Language.MemberExpressionAst] -and $p.Member -eq $s) { continue }
        if ($p -is [System.Management.Automation.Language.IndexExpressionAst] -and $p.Index -eq $s) { continue }
        if ($p -is [System.Management.Automation.Language.CommandExpressionAst] -and $p.Parent -is [System.Management.Automation.Language.PipelineAst] -and $false) { continue }
        if ($p -is [System.Management.Automation.Language.CommandAst] -and $p.CommandElements[0] -eq $s) { continue }
        if ($p -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $p.Member -eq $s) { continue }
        if ($p -is [System.Management.Automation.Language.TypeExpressionAst]) { continue }
        $isKey = $false
        $q = $s
        while ($q -and $q.Parent) {
            if ($q.Parent -is [System.Management.Automation.Language.HashtableAst]) {
                foreach ($kv in $q.Parent.KeyValuePairs) { if ($kv.Item1 -eq $q) { $isKey = $true } }
                break
            }
            if ($q.Parent -is [System.Management.Automation.Language.StatementBlockAst]) { break }
            $q = $q.Parent
        }
        if ($isKey) { continue }
        $text = $s.Value
        if ($s -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { $text = $s.Value }
        [pscustomobject]@{ Text = $text; Line = $s.Extent.StartLineNumber; Here = ($s.StringConstantType -like 'Here*') }
    }
}

function Get-PowerShellTextItem {
    param([string]$Path, [string]$Rel)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (-not $ast) { return }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'

    $emit = {
        param($Text, $Line, $Kind)
        $t = ConvertTo-FormatFreeText -Text $Text
        $t = ($t -replace '\s+', ' ').Trim()
        if (-not (Test-ProseCandidate -Text $t -Strict:$false)) { return }
        $key = '{0}|{1}' -f $Line, $t
        if ($seen.Add($key)) { [pscustomobject]@{ File = $Rel; Line = $Line; Kind = $Kind; Text = $t } }
    }

    # 1. message, dialog, status, log and console commands
    $cmds = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($c in $cmds) {
        $name = $c.GetCommandName()
        if (-not $name -or $name -notmatch $script:MessageCommandPattern) { continue }
        foreach ($s in (Get-StringNodeText -Node $c)) {
            if ($s.Here) { continue }
            & $emit $s.Text $s.Line ('message:' + $name)
        }
    }
    # 2. assignments to a text property
    $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
    foreach ($a in $assigns) {
        $left = $a.Left
        if ($left -is [System.Management.Automation.Language.MemberExpressionAst] -and $left.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
            $left.Member.Value -match $script:TextMemberPattern) {
            foreach ($s in (Get-StringNodeText -Node $a.Right)) { if (-not $s.Here) { & $emit $s.Text $s.Line ('property:' + $left.Member.Value) } }
        }
    }
    # 3. throw
    $throws = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst] }, $true)
    foreach ($t in $throws) {
        if (-not $t.Pipeline) { continue }
        foreach ($s in (Get-StringNodeText -Node $t.Pipeline)) { if (-not $s.Here) { & $emit $s.Text $s.Line 'throw' } }
    }
    # 3b. any other sentence-like string literal: a status text kept in a variable, a hashtable value, a method argument
    $allStrings = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
            $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
        }, $true)
    foreach ($sc in $allStrings) {
        if ($sc.StringConstantType -like 'Here*') { continue }
        $p = $sc.Parent
        if ($p -is [System.Management.Automation.Language.MemberExpressionAst] -and $p.Member -eq $sc) { continue }
        if ($p -is [System.Management.Automation.Language.IndexExpressionAst] -and $p.Index -eq $sc) { continue }
        if ($p -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $p.Member -eq $sc) { continue }
        if ($p -is [System.Management.Automation.Language.CommandAst] -and $p.CommandElements[0] -eq $sc) { continue }
        $raw = ($sc.Value -replace '\s+', ' ').Trim()
        if ($raw -notmatch '^[A-Z].{10,}[.!?:]$') { continue }
        if (-not (Test-ProseCandidate -Text (ConvertTo-FormatFreeText -Text $raw) -Strict:$true)) { continue }
        & $emit $sc.Value $sc.Extent.StartLineNumber 'string-literal'
    }
    # 3c. display text of any shape: a hashtable value under a text key, and an argument of a call that takes display text
    $textKeys = '^(Label|Tip|ToolTip|Text|Title|Hint|Watermark|Placeholder|Description|Message|Caption|Header|Content)$'
    foreach ($ht in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)) {
        foreach ($kv in $ht.KeyValuePairs) {
            $k = $kv.Item1
            if ($k -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $k.Value -match $textKeys) {
                foreach ($s in (Get-StringNodeText -Node $kv.Item2)) { if (-not $s.Here) { & $emit $s.Text $s.Line ('hashtable:' + $k.Value) } }
            }
        }
    }
    foreach ($im in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
        $mn = $im.Member
        if ($mn -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $mn.Value -match 'Watermark|ToolTip|Placeholder|SetText|SetHint|ShowMessage') {
            foreach ($a in $im.Arguments) {
                foreach ($s in (Get-StringNodeText -Node $a)) { if (-not $s.Here) { & $emit $s.Text $s.Line ('call:' + $mn.Value) } }
            }
        }
    }
    foreach ($cmd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $cn = $cmd.GetCommandName()
        if ($cn -and $cn -match 'Watermark|ToolTip|Placeholder|Hint') {
            foreach ($s in (Get-StringNodeText -Node $cmd)) { if (-not $s.Here) { & $emit $s.Text $s.Line ('call:' + $cn) } }
        }
    }
    # 4. XAML in here-strings
    $here = $ast.FindAll({
            param($n)
            ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -like 'Here*') -or
            ($n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] -and $n.StringConstantType -like 'Here*')
        }, $true)
    foreach ($h in $here) {
        if ($h.Value -notmatch '<(Window|Grid|StackPanel|UserControl|Button|TextBlock|Border|DockPanel)\b') { continue }
        foreach ($x in (Get-XamlTextItemFromString -Content $h.Value -Rel $Rel -LineOffset ($h.Extent.StartLineNumber + 1) -Kind 'xaml-in-script')) { $x }
    }
    # 5. comment-based help
    foreach ($tok in $tokens) {
        if ($tok.Kind -ne [System.Management.Automation.Language.TokenKind]::Comment) { continue }
        $txt = $tok.Text
        if ($txt -notmatch '(?m)^\s*\.(SYNOPSIS|DESCRIPTION|PARAMETER|EXAMPLE|NOTES|LINK)\b') { continue }
        $lines = $txt -split "`r?`n"
        $section = ''
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $l = $lines[$i]
            if ($l -match '^\s*\.(SYNOPSIS|DESCRIPTION|PARAMETER|EXAMPLE|NOTES|LINK|INPUTS|OUTPUTS|COMPONENT|ROLE|FUNCTIONALITY|FORWARDHELPTARGETNAME|EXTERNALHELP)\b') {
                $section = $Matches[1]
                continue
            }
            if ($section -notin 'SYNOPSIS', 'DESCRIPTION', 'PARAMETER', 'EXAMPLE') { continue }
            if ($l -match '^\s*(<#|#>)\s*$') { continue }
            if ($section -eq 'EXAMPLE' -and $l -match '^\s*(PS[^>]*>|\.\\|\./|[A-Za-z]+-[A-Za-z]+\b|\$|&|\.\s)') { continue }
            if ($section -eq 'PARAMETER' -and $l -match '^\s*\w+\s*$') { continue }
            $t = ConvertTo-ProseText -Line $l
            $t = ($t -replace '\s+', ' ').Trim()
            if (-not (Test-ProseCandidate -Text $t -Strict:$false)) { continue }
            [pscustomobject]@{ File = $Rel; Line = $tok.Extent.StartLineNumber + $i; Kind = 'help:' + $section; Text = $t }
        }
    }
}

# --- C# -----------------------------------------------------------------------------------------

function Get-CSharpTextItem {
    param([string]$Path, [string]$Rel)
    $lines = [IO.File]::ReadAllLines($Path)
    $rx = New-Object System.Text.RegularExpressions.Regex('(?<!\w)\$?@?\$?"(?<s>(?:[^"\\]|\\.)*)"')
    $inBlockComment = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]
        $trim = $l.TrimStart()
        if ($inBlockComment) { if ($l -match '\*/') { $inBlockComment = $false }; continue }
        if ($trim.StartsWith('/*')) { if ($l -notmatch '\*/') { $inBlockComment = $true }; continue }
        if ($trim.StartsWith('//')) { continue }
        if ($trim -match '^\[(Obsolete|JsonPropertyName|SuppressMessage|DllImport|Fact|Theory|InlineData|Description|Category)') { continue }
        if ($l -match '\b(Regex|GeneratedRegex|RegexOptions|nameof|Environment\.GetEnvironmentVariable|ArgumentNullException|ArgumentException|Path\.Combine|File\.|Directory\.|Process\.Start|ProcessStartInfo|WqlQuery|CimQuery|SqlCommand|CommandText)\b') { continue }
        foreach ($m in $rx.Matches($l)) {
            $s = $m.Groups['s'].Value
            $s = $s.Replace('\"', '"').Replace('\\', '\')
            $s = ConvertTo-FormatFreeText -Text $s
            $s = ($s -replace '\s+', ' ').Trim()
            if (-not (Test-ProseCandidate -Text $s -Strict:$true)) { continue }
            [pscustomobject]@{ File = $Rel; Line = $i + 1; Kind = 'csharp-string'; Text = $s }
        }
    }
}

# --- repository walk ----------------------------------------------------------------------------

function Get-ManifestTextItem {
    <#
    .SYNOPSIS
        Reads the user-facing text of a module manifest (.psd1): Description, Synopsis and ReleaseNotes.
    .DESCRIPTION
        The manifest Description shows in Get-Module and in the PowerShell Gallery, so it is user-facing text.
        A value in single or double quotes on one line is read; a doubled quote inside the value is one quote.
    #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Rel)
    $lines = [IO.File]::ReadAllLines($Path)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $m = [regex]::Match($lines[$i], "^\s*(?:Description|Synopsis|ReleaseNotes)\s*=\s*(?:'(?<s>(?:[^']|'')*)'|""(?<d>(?:[^""]|"""")*)"")")
        if (-not $m.Success) { continue }
        $text = if ($m.Groups['s'].Success) { $m.Groups['s'].Value.Replace("''", "'") } else { $m.Groups['d'].Value.Replace('""', '"') }
        if ($text.Trim()) { [pscustomobject]@{ File = $Rel; Line = $i + 1; Kind = 'manifest'; Text = $text.Trim() } }
    }
}

function Get-LexiconTextItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [ValidateSet('Auto', 'PowerShell', 'Apd')][string]$RepoType = 'Auto'
    )
    $Root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\', '/')
    if ($RepoType -eq 'Auto') {
        $RepoType = 'PowerShell'
        if (Test-Path -LiteralPath (Join-Path $Root 'src')) {
            if (@(Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Recurse -Filter *.csproj -File -ErrorAction SilentlyContinue).Count -gt 0) { $RepoType = 'Apd' }
        }
    }
    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($Root.Length + 1).Replace('\', '/')
        if ($rel -match $script:SkipDirPattern -and $rel -notmatch '^(src|docs)/') { continue }
        if ($rel -match '(^|/)(bin|obj|node_modules|\.git|images)/') { continue }
        $ext = $f.Extension.ToLowerInvariant()
        if ($ext -eq '.md') {
            if ($rel -eq 'README.md' -or $rel -like 'docs/*') { Get-MarkdownTextItem -Path $f.FullName -Rel $rel }
            elseif ($rel -eq 'CHANGELOG.md') { Get-MarkdownTextItem -Path $f.FullName -Rel $rel -UnreleasedOnly }
            continue
        }
        if ($ext -eq '.xaml') {
            if ($RepoType -eq 'Apd' -and $rel -notlike 'src/*') { continue }
            $content = [IO.File]::ReadAllText($f.FullName)
            Get-XamlTextItemFromString -Content $content -Rel $rel -LineOffset 1 -Kind 'xaml'
            continue
        }
        if ($RepoType -eq 'PowerShell' -and $ext -eq '.psd1') {
            Get-ManifestTextItem -Path $f.FullName -Rel $rel
            continue
        }
        if ($RepoType -eq 'PowerShell' -and ($ext -eq '.ps1' -or $ext -eq '.psm1')) {
            if ($rel -match '^(tools|installer)/') { continue }
            Get-PowerShellTextItem -Path $f.FullName -Rel $rel
            continue
        }
        if ($RepoType -eq 'Apd' -and $ext -eq '.cs' -and $rel -like 'src/*') {
            Get-CSharpTextItem -Path $f.FullName -Rel $rel
            continue
        }
        if ($RepoType -eq 'Apd' -and ($ext -eq '.ps1' -or $ext -eq '.psm1') -and $rel -like 'src/*') {
            Get-PowerShellTextItem -Path $f.FullName -Rel $rel
        }
    }
}

# --- matching -----------------------------------------------------------------------------------

function Test-AllowListed {
    param($Allow, [string]$File, [string]$RuleId, [string]$Phrase, [int]$Line)
    foreach ($a in $Allow) {
        if ($a.file -and $File -notlike $a.file) { continue }
        if ($a.rule -and $a.rule -ne '*' -and $a.rule -ne $RuleId) { continue }
        if ($a.phrase -and $Phrase -notmatch $a.phrase) { continue }
        if ($a.PSObject.Properties['line'] -and $a.line -and [int]$a.line -ne $Line) { continue }
        return $true
    }
    $false
}

function Find-LexiconHit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)]$Lexicon,
        $AllowList = @()
    )
    foreach ($it in $Item) {
        foreach ($e in @($Lexicon.entries)) {
            if (-not $e.FlaggedRegex -or $e.FlaggedRegex.Count -eq 0) { continue }
            if ($e.PSObject.Properties['scope'] -and $e.scope -eq 'tracked-files') { continue }
            $allowed = $false
            foreach ($ar in $e.AllowRegex) { if ($ar.IsMatch($it.Text)) { $allowed = $true; break } }
            if ($allowed) { continue }
            foreach ($rx in $e.FlaggedRegex) {
                foreach ($m in $rx.Matches($it.Text)) {
                    if (Test-AllowListed -Allow $AllowList -File $it.File -RuleId $e.id -Phrase $m.Value -Line $it.Line) { continue }
                    [pscustomobject]@{
                        File        = $it.File
                        Line        = $it.Line
                        Kind        = $it.Kind
                        Phrase      = $m.Value
                        RuleId      = $e.id
                        Severity    = $e.severity
                        Status      = $e.status
                        Replacement = $e.replacement
                        Text        = $it.Text
                    }
                }
            }
        }
    }
}


# --- banned names in every tracked file -----------------------------------------------------------

$script:BinaryExtensions = '.png', '.jpg', '.jpeg', '.gif', '.ico', '.bmp', '.dll', '.exe', '.zip', '.7z', '.gz', '.pdf', '.pfx', '.cer', '.snk', '.woff', '.woff2', '.ttf', '.otf', '.nupkg', '.msi', '.cab', '.bin', '.mp4', '.mov', '.wav', '.mp3', '.pdb', '.lib', '.obj', '.xz', '.tgz'

function Get-TrackedFileList {
    # Tracked files of a repository: git ls-files when git and a .git folder exist, else every file outside .git.
    param([string]$Root)
    $files = $null
    if ((Test-Path -LiteralPath (Join-Path $Root '.git')) -and (Get-Command git -ErrorAction SilentlyContinue)) {
        $out = & git -C $Root ls-files -z 2>$null
        if ($LASTEXITCODE -eq 0 -and $out) { $files = @(($out -join '') -split "`0" | Where-Object { $_ }) }
    }
    if (-not $files) {
        $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName.Substring($Root.Length + 1) -notmatch '^(\.git|node_modules|bin|obj)[\/]' } |
                ForEach-Object { $_.FullName.Substring($Root.Length + 1).Replace('\', '/') })
    }
    $files
}

function Get-BannedNameHit {
    # Matches the entries with scope 'tracked-files' (banned names) against every line of every tracked text file.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)]$Lexicon,
        $AllowList = @()
    )
    $Root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\', '/')
    $entries = @($Lexicon.entries | Where-Object { $_.PSObject.Properties['scope'] -and $_.scope -eq 'tracked-files' -and $_.FlaggedRegex.Count -gt 0 })
    foreach ($rel in (Get-TrackedFileList -Root $Root)) {
        if ($rel -match '^(lexicon|Lib/Lexicon)/' -or $rel -eq 'Tests/Lexicon.Guard.Tests.ps1') { continue }   # the lexicon data and its guard test hold the banned patterns
        if ($script:BinaryExtensions -contains [IO.Path]::GetExtension($rel).ToLowerInvariant()) { continue }
        $full = Join-Path $Root $rel
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { continue }
        if ((Get-Item -LiteralPath $full).Length -gt 4MB) { continue }
        $bytes = [IO.File]::ReadAllBytes($full)
        if ([Array]::IndexOf($bytes, [byte]0) -ge 0 -and $bytes.Length -gt 1 -and $bytes[1] -ne 0) { continue }   # binary (UTF-16 text has a zero in the second byte)
        $lines = [IO.File]::ReadAllLines($full)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            foreach ($e in $entries) {
                $spans = New-Object System.Collections.Generic.List[int[]]
                foreach ($rx in $e.FlaggedRegex) {
                    foreach ($m in $rx.Matches($lines[$i])) {
                        $overlap = $false
                        foreach ($sp in $spans) { if ($m.Index -lt $sp[1] -and ($m.Index + $m.Length) -gt $sp[0]) { $overlap = $true; break } }
                        if ($overlap) { continue }
                        $spans.Add(@($m.Index, ($m.Index + $m.Length)))
                        if (Test-AllowListed -Allow $AllowList -File $rel -RuleId $e.id -Phrase $m.Value -Line ($i + 1)) { continue }
                        $t = $lines[$i].Trim()
                        if ($t.Length -gt 200) { $t = $t.Substring(0, 200) }
                        [pscustomobject]@{
                            File = $rel; Line = $i + 1; Kind = 'tracked-file'; Phrase = $m.Value; RuleId = $e.id
                            Severity = $e.severity; Status = $e.status; Replacement = $e.replacement; Class = $e.category; Text = $t
                        }
                    }
                }
            }
        }
    }
}

function Invoke-LexiconCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [string]$LexiconPath,
        [string]$AllowListPath,
        [ValidateSet('Auto', 'PowerShell', 'Apd')][string]$RepoType = 'Auto'
    )
    if (-not $LexiconPath) { $LexiconPath = Join-Path $PSScriptRoot 'configmgr-lexicon.json' }
    if (-not $AllowListPath) { $AllowListPath = Join-Path $PSScriptRoot 'lexicon-allowlist.json' }
    $lex = Import-LexiconData -Path $LexiconPath
    $allow = Import-LexiconAllowList -Path $AllowListPath
    $items = @(Get-LexiconTextItem -Root $Root -RepoType $RepoType)
    Find-LexiconHit -Item $items -Lexicon $lex -AllowList $allow
}

Export-ModuleMember -Function Import-LexiconData, Import-LexiconAllowList, Get-LexiconTextItem, Find-LexiconHit, Invoke-LexiconCheck, Get-BannedNameHit
