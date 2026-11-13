#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Footer guard. It parses the XAML files, the XAML here-strings and the PowerShell scripts of the app.
# It creates no WPF control and opens no window.
#
# Rules:
#   1. The footer status element exists, and it sits in a Border or a DockPanel (Initialize-SuiteFooter needs one of them).
#   2. No footer element sets FontFamily, FontSize, FontWeight, FontStyle or TextDecorations, in XAML or in code.
#      The font of the footer comes from SuiteCommon (Get-SuiteFooterStyle). Only the colors change with the theme.
#   3. The app calls Initialize-SuiteFooter once, in the Loaded handler, and refreshes the footer with Update-SuiteFooter
#      after the load, after the Options dialog closes and after a connect.
#   4. A site code box or a site server box is editable only in the Options dialog (and in a first-run dialog).
#      The footer and every other window show the values as read-only text.

BeforeAll {
    $script:FtRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:FtNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $script:FtKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
    $script:FtExcluded = '^(\.claude|\.git|Tests|tools|Logs|Lib)([\\/]|$)'
    $script:FtFontProperty = @('FontFamily', 'FontSize', 'FontWeight', 'FontStyle', 'TextDecorations')
    $script:FtEditableType = @('TextBox', 'ComboBox', 'PasswordBox', 'NumericUpDown')

    # ---- Settings of this repository ----------------------------------------------------------------------------
    # StartScript     : the script that builds the window and calls the footer helpers.
    # ExtraScript     : more script files that hold a refresh place (the Options dialog may live in its own file).
    # StatusName      : x:Name of the footer status element.
    # FooterDocument  : the XAML document of the main window, as a pattern for its source name: a XAML file name, or
    #                   'file here-string N' for a here-string. A dialog may have its own status line. The footer rules do not apply to it.
    # StatusVariable  : the variable names that hold the status element in code.
    # NoSite          : true for an app without a ConfigMgr connection. The app calls Initialize-SuiteFooter for the shared
    #                   status style and hides the site text. It has no wrapper, no refresh and no site box.
    # Wrapper         : the function of the app that calls Update-SuiteFooter with the saved site code and site server.
    # Update          : the places that must refresh the footer. Kind Handler: a call of the event member (Member) on the
    #                   target (Target, empty for any). Kind Function: the function Name. After: a command that must come
    #                   before the refresh in that place (empty for none).
    # OptionsFunction : the functions that may hold an editable site code box or site server box (the Options dialog
    #                   and a first-run dialog).
    # SiteNamePattern : the names of the site code boxes and the site server boxes.
    # ReadOnlyName    : names that match the pattern and are read-only text. A string constant with such a name is not a site
    #                   box in code. XAML relies on the element type and on IsReadOnly.
    $script:FtConfig = @{
        StartScript     = 'start-apppackager.ps1'
        ExtraScript     = @()
        StatusName      = 'txtStatus'
        FooterDocument  = '^MainWindow[.]xaml$'
        StatusVariable  = @('txtStatus')
        NoSite          = $false
        Wrapper         = 'Update-ApFooter'
        Update          = @(
            @{ Label = 'after the load'; Kind = 'Handler'; Member = 'Add_Loaded'; Target = '$window'; After = '' }
            @{ Label = 'after the Options dialog closes'; Kind = 'Handler'; Member = 'Add_Click'; Target = '$btnOptions'; After = 'Show-OptionsDialog' }
            @{ Label = 'after the first-run wizard'; Kind = 'Handler'; Member = 'Add_ContentRendered'; Target = '$window'; After = 'Show-FirstRunWizard' }
            @{ Label = 'after a connect check'; Kind = 'Handler'; Member = 'Add_Click'; Target = '$btnCheckMECM'; After = '' }
        )
        OptionsFunction = @('Show-OptionsDialog', 'Show-FirstRunWizard', 'Connect-OptionsFieldLink', 'Get-OptionsSectionDefinition', 'Initialize-OptionsTestSection')
        SiteNamePattern = '(?i)^((txt|cbo|cmb|tb)\w*(SiteCode|SiteServer|SiteSrv|SmsProvider|Provider|Site)\d*|cm\.(siteCode|provider))$'
        ReadOnlyName    = @('txtCmSite', 'txtFooterSite')
    }
    # --------------------------------------------------------------------------------------------------------------

    function Get-FtRelativePath {
        param([string]$FullName)
        return $FullName.Substring($script:FtRepoRoot.Length).TrimStart('\', '/')
    }

    function Get-FtScriptFile {
        return @(Get-ChildItem -LiteralPath $script:FtRepoRoot -File -Recurse | Where-Object {
                $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' -and (Get-FtRelativePath $_.FullName) -notmatch $script:FtExcluded
            })
    }

    function Get-FtXamlFile {
        return @(Get-ChildItem -LiteralPath $script:FtRepoRoot -Filter '*.xaml' -File -Recurse | Where-Object {
                (Get-FtRelativePath $_.FullName) -notmatch $script:FtExcluded
            })
    }

    # Every XAML document of the app: the XAML files and the XAML here-strings of the scripts.
    function Get-FtDocument {
        $docs = @()
        foreach ($f in @(Get-FtXamlFile)) { $docs += @{ Source = (Get-FtRelativePath $f.FullName); Text = (Get-Content -LiteralPath $f.FullName -Raw) } }
        foreach ($f in @(Get-FtScriptFile)) {
            $i = 0
            foreach ($h in @(Get-FtHereString -Text (Get-Content -LiteralPath $f.FullName -Raw))) {
                $i++
                $docs += @{ Source = ('{0} here-string {1}' -f (Get-FtRelativePath $f.FullName), $i); Text = $h.Text }
            }
        }
        return $docs
    }

    function Get-FtAst {
        param([string]$Text)
        $tokens = $null; $errors = $null
        return [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
    }

    function Get-FtFunctionName {
        param($Node)
        $p = $Node.Parent
        while ($null -ne $p) {
            if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $p.Name }
            $p = $p.Parent
        }
        return ''
    }

    # The XAML here-strings of a script: the text and the function that holds it.
    function Get-FtHereString {
        param([string]$Text)
        $ast = Get-FtAst -Text $Text
        $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -in 'SingleQuotedHereString', 'DoubleQuotedHereString' }, $true)
        foreach ($n in @($nodes)) {
            if ($n.Value -notmatch ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:FtNamespace) + '"')) { continue }
            [PSCustomObject]@{ Text = $n.Value; Function = (Get-FtFunctionName $n) }
        }
    }

    function Test-FtFontName {
        param([string]$Name)
        foreach ($f in $script:FtFontProperty) { if ($Name -match ('(^|\.)' + $f + '$')) { return $true } }
        return $false
    }

    # The footer status element of a XAML document: its bar problems and its font problems.
    # Found is false when the document has no element of that name.
    function Get-FtFooterXamlProblem {
        param([string]$Text, [string]$Source, [string]$StatusName)
        [xml]$x = $Text
        $problems = New-Object System.Collections.Generic.List[string]
        $status = $null
        foreach ($el in @($x.SelectNodes('//*'))) {
            if ($el.GetAttribute('Name', $script:FtKeyNamespace) -eq $StatusName -or $el.GetAttribute('Name') -eq $StatusName) { $status = $el; break }
        }
        if ($null -eq $status) { return [PSCustomObject]@{ Found = $false; Problems = @() } }
        $bar = $status.ParentNode
        if ($null -eq $bar -or $bar.NodeType -ne 'Element' -or $bar.LocalName -notin 'Border', 'DockPanel') {
            $kind = if ($bar -and $bar.NodeType -eq 'Element') { $bar.LocalName } else { 'nothing' }
            [void]$problems.Add(('{0}: the footer status element {1} sits in {2}, not in a Border or a DockPanel' -f $Source, $StatusName, $kind))
        }
        $scope = @($status)
        if ($bar -and $bar.NodeType -eq 'Element') { $scope = @($bar) + @($bar.SelectNodes('.//*')) }
        foreach ($el in $scope) {
            foreach ($a in @($el.Attributes)) {
                if (Test-FtFontName $a.Name) { [void]$problems.Add(('{0}: <{1}> in the footer sets {2}' -f $Source, $el.LocalName, $a.Name)) }
            }
            if (Test-FtFontName $el.LocalName) { [void]$problems.Add(('{0}: the footer has the property element <{1}>' -f $Source, $el.LocalName)) }
            if ($el.LocalName -eq 'Setter' -and (Test-FtFontName $el.GetAttribute('Property'))) {
                [void]$problems.Add(('{0}: a style in the footer sets {1}' -f $Source, $el.GetAttribute('Property')))
            }
        }
        return [PSCustomObject]@{ Found = $true; Problems = @($problems) }
    }

    # An editable box with a site code name or a site server name in a XAML document.
    function Get-FtSiteBoxXamlProblem {
        param([string]$Text, [string]$Source)
        [xml]$x = $Text
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($el in @($x.SelectNodes('//*'))) {
            if ($el.LocalName -notin $script:FtEditableType) { continue }
            $name = $el.GetAttribute('Name', $script:FtKeyNamespace)
            if (-not $name) { $name = $el.GetAttribute('Name') }
            if (-not $name -or $name -notmatch $script:FtConfig.SiteNamePattern) { continue }
            if ($el.GetAttribute('IsReadOnly') -eq 'True') { continue }
            [void]$found.Add(('{0}: the editable <{1}> {2} sits outside the Options dialog' -f $Source, $el.LocalName, $name))
        }
        return @($found)
    }

    # Problems of one script text: a font set on the status variable, a site box name outside the Options functions,
    # an editable site box in a XAML here-string outside the Options functions.
    function Get-FtCodeProblem {
        param([string]$Text, [string]$Source)
        $found = New-Object System.Collections.Generic.List[string]
        $ast = Get-FtAst -Text $Text
        $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
        foreach ($a in @($assigns)) {
            $left = $a.Left
            if ($left -isnot [System.Management.Automation.Language.MemberExpressionAst]) { continue }
            $member = [string]$left.Member.Extent.Text
            if (-not (Test-FtFontName $member)) { continue }
            $target = [string]$left.Expression.Extent.Text
            $isStatus = $false
            foreach ($v in $script:FtConfig.StatusVariable) { if ($target -eq ('$' + $v)) { $isStatus = $true } }
            if ($target -match '(?i)footer') { $isStatus = $true }
            if ($isStatus) { [void]$found.Add(('{0}: line {1} sets {2} on the footer element {3}' -f $Source, $a.Extent.StartLineNumber, $member, $target)) }
        }
        $strings = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -notin 'SingleQuotedHereString', 'DoubleQuotedHereString' }, $true)
        foreach ($s in @($strings)) {
            if ($s.Value -notmatch $script:FtConfig.SiteNamePattern) { continue }
            if ($script:FtConfig.ReadOnlyName -contains $s.Value) { continue }
            $fn = Get-FtFunctionName $s
            if ($script:FtConfig.OptionsFunction -contains $fn) { continue }
            $where = if ($fn) { 'function ' + $fn } else { 'script scope' }
            [void]$found.Add(('{0}: line {1} uses the site box name {2} in {3}, outside the Options dialog' -f $Source, $s.Extent.StartLineNumber, $s.Value, $where))
        }
        foreach ($h in @(Get-FtHereString -Text $Text)) {
            if ($script:FtConfig.OptionsFunction -contains $h.Function) { continue }
            $where = if ($h.Function) { $h.Function } else { 'script scope' }
            foreach ($p in @(Get-FtSiteBoxXamlProblem -Text $h.Text -Source ($Source + ' here-string in ' + $where))) { [void]$found.Add($p) }
        }
        return @($found)
    }

    # The footer calls of the start script: Initialize-SuiteFooter once in the Loaded handler, a wrapper that calls
    # Update-SuiteFooter, and a refresh in every place of the settings.
    function Get-FtWiringProblem {
        param([string]$Text, [string]$Source)
        $found = New-Object System.Collections.Generic.List[string]
        $ast = Get-FtAst -Text $Text
        $wrapper = $script:FtConfig.Wrapper
        if ($script:FtConfig.NoSite) {
            $initNo = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Initialize-SuiteFooter' }, $true))
            if ($initNo.Count -ne 1) { [void]$found.Add(('{0}: Initialize-SuiteFooter is called {1} times, expected 1' -f $Source, $initNo.Count)); return @($found) }
            $assign = $initNo[0].Parent
            while ($null -ne $assign -and $assign -isnot [System.Management.Automation.Language.AssignmentStatementAst]) { $assign = $assign.Parent }
            if ($null -eq $assign) { [void]$found.Add(('{0}: the result of Initialize-SuiteFooter is not kept, so the site text cannot be hidden' -f $Source)); return @($found) }
            $var = [string]$assign.Left.Extent.Text
            $hides = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and [string]$n.Left.Member.Value -eq 'Visibility' -and $n.Left.Expression.Extent.Text -eq $var -and $n.Right.Extent.Text -match 'Collapsed' }, $true))
            if ($hides.Count -eq 0) { [void]$found.Add(('{0}: the site text of the footer ({1}) is not collapsed, and the app has no ConfigMgr connection' -f $Source, $var)) }
            $loadedNo = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$n.Member.Value -eq 'Add_Loaded' }, $true))
            $inLoaded = $false
            foreach ($l in $loadedNo) { if ($initNo[0].Extent.StartOffset -ge $l.Extent.StartOffset -and $initNo[0].Extent.EndOffset -le $l.Extent.EndOffset) { $inLoaded = $true } }
            if (-not $inLoaded) { [void]$found.Add(('{0}: Initialize-SuiteFooter is not called in the Loaded handler' -f $Source)) }
            return @($found)
        }
        $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))
        $init = @($commands | Where-Object { $_.GetCommandName() -eq 'Initialize-SuiteFooter' })
        if ($init.Count -ne 1) { [void]$found.Add(('{0}: Initialize-SuiteFooter is called {1} times, expected 1' -f $Source, $init.Count)) }
        $wrapperAst = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $wrapper }, $true)
        $calls = @('Update-SuiteFooter', $wrapper)
        if ($null -eq $wrapperAst) { [void]$found.Add(('{0}: the function {1} is missing' -f $Source, $wrapper)) }
        else {
            $inner = @($wrapperAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Update-SuiteFooter' }, $true))
            if ($inner.Count -eq 0) { [void]$found.Add(('{0}: the function {1} does not call Update-SuiteFooter' -f $Source, $wrapper)) }
        }
        foreach ($u in $script:FtConfig.Update) {
            $kind = [string]$u.Kind; $uname = [string]$u.Name; $member = [string]$u.Member; $target = [string]$u.Target; $after = [string]$u.After
            if ($kind -eq 'Function') {
                $contexts = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $uname }, $true))
            }
            else {
                $contexts = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$n.Member.Value -eq $member -and ((-not $target) -or $n.Expression.Extent.Text -eq $target) }, $true))
            }
            if ($contexts.Count -eq 0) { [void]$found.Add(('{0}: the place "{1}" does not exist' -f $Source, $u.Label)); continue }
            $ok = $false
            foreach ($c in $contexts) {
                $cmds = @($c.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))
                $first = -1
                if ($after) {
                    foreach ($cmd in $cmds) { if ($cmd.GetCommandName() -eq $after -and ($first -lt 0 -or $cmd.Extent.StartOffset -lt $first)) { $first = $cmd.Extent.StartOffset } }
                }
                foreach ($cmd in $cmds) {
                    if ($calls -notcontains $cmd.GetCommandName()) { continue }
                    if ($after -and ($first -lt 0 -or $cmd.Extent.StartOffset -lt $first)) { continue }
                    $ok = $true
                }
            }
            if (-not $ok) { [void]$found.Add(('{0}: the footer is not refreshed {1}' -f $Source, $u.Label)) }
        }
        $loaded = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$n.Member.Value -eq 'Add_Loaded' }, $true))
        $initInLoaded = $false
        foreach ($l in $loaded) { foreach ($i in $init) { if ($i.Extent.StartOffset -ge $l.Extent.StartOffset -and $i.Extent.EndOffset -le $l.Extent.EndOffset) { $initInLoaded = $true } } }
        if ($init.Count -eq 1 -and -not $initInLoaded) { [void]$found.Add(('{0}: Initialize-SuiteFooter is not called in the Loaded handler' -f $Source)) }
        return @($found)
    }

    $script:FtSampleNs = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
}

Describe 'Footer guard' {
    Context 'the app' {
        It 'finds the XAML, the scripts and the footer status element (the scan is not empty)' {
            @(Get-FtXamlFile).Count | Should -BeGreaterThan 0
            @(Get-FtScriptFile).Count | Should -BeGreaterThan 0
            $found = 0
            foreach ($d in @(Get-FtDocument | Where-Object { $_.Source -match $script:FtConfig.FooterDocument })) {
                if ((Get-FtFooterXamlProblem -Text $d.Text -Source $d.Source -StatusName $script:FtConfig.StatusName).Found) { $found++ }
            }
            $found | Should -Be 1
        }

        It 'sets no font on a footer element in the XAML, and the status element sits in a Border or a DockPanel' {
            $v = @()
            foreach ($d in @(Get-FtDocument | Where-Object { $_.Source -match $script:FtConfig.FooterDocument })) {
                $v += @((Get-FtFooterXamlProblem -Text $d.Text -Source $d.Source -StatusName $script:FtConfig.StatusName).Problems)
            }
            ($v -join [Environment]::NewLine) | Should -BeNullOrEmpty
        }

        It 'sets no font on a footer element in code' {
            $v = @()
            foreach ($f in @(Get-FtScriptFile)) { $v += @(Get-FtCodeProblem -Text (Get-Content -LiteralPath $f.FullName -Raw) -Source (Get-FtRelativePath $f.FullName) | Where-Object { $_ -match 'sets Font|sets TextDecorations' }) }
            ($v -join [Environment]::NewLine) | Should -BeNullOrEmpty
        }

        It 'calls Initialize-SuiteFooter after the load and refreshes the footer after the load, after the Options dialog and after a connect' {
            $path = Join-Path $script:FtRepoRoot $script:FtConfig.StartScript
            Test-Path -LiteralPath $path | Should -BeTrue
            $text = Get-Content -LiteralPath $path -Raw
            foreach ($extra in $script:FtConfig.ExtraScript) {
                $extraPath = Join-Path $script:FtRepoRoot $extra
                Test-Path -LiteralPath $extraPath | Should -BeTrue
                $text += [Environment]::NewLine + (Get-Content -LiteralPath $extraPath -Raw)
            }
            $v = @(Get-FtWiringProblem -Text $text -Source $script:FtConfig.StartScript)
            ($v -join [Environment]::NewLine) | Should -BeNullOrEmpty
        }

        It 'has an editable site code box or site server box only in the Options dialog' {
            $v = @()
            foreach ($f in @(Get-FtXamlFile)) { $v += @(Get-FtSiteBoxXamlProblem -Text (Get-Content -LiteralPath $f.FullName -Raw) -Source (Get-FtRelativePath $f.FullName)) }
            foreach ($f in @(Get-FtScriptFile)) { $v += @(Get-FtCodeProblem -Text (Get-Content -LiteralPath $f.FullName -Raw) -Source (Get-FtRelativePath $f.FullName) | Where-Object { $_ -match 'site box|editable' }) }
            ($v -join [Environment]::NewLine) | Should -BeNullOrEmpty
        }
    }

    Context 'bad samples' {
        It 'fails a font on the footer status element, on a sibling in the bar and in a property element' {
            $bad = '<Grid ' + $script:FtSampleNs + '><Border><TextBlock x:Name="txtStatus" FontSize="11"/></Border></Grid>'
            $r = Get-FtFooterXamlProblem -Text $bad -Source 'sample' -StatusName 'txtStatus'
            $r.Found | Should -BeTrue
            @($r.Problems).Count | Should -Be 1
            $bad2 = '<Grid ' + $script:FtSampleNs + '><DockPanel><TextBlock Text="Site" FontWeight="Bold" DockPanel.Dock="Right"/><TextBlock x:Name="txtStatus"/></DockPanel></Grid>'
            @((Get-FtFooterXamlProblem -Text $bad2 -Source 'sample' -StatusName 'txtStatus').Problems).Count | Should -Be 1
            $bad3 = '<Grid ' + $script:FtSampleNs + '><Border><TextBlock x:Name="txtStatus"><TextBlock.FontFamily>Consolas</TextBlock.FontFamily></TextBlock></Border></Grid>'
            @((Get-FtFooterXamlProblem -Text $bad3 -Source 'sample' -StatusName 'txtStatus').Problems).Count | Should -Be 1
        }

        It 'fails a status element in a layout that the footer helper cannot use' {
            $bad = '<Grid ' + $script:FtSampleNs + '><StackPanel><TextBlock x:Name="txtStatus"/></StackPanel></Grid>'
            @((Get-FtFooterXamlProblem -Text $bad -Source 'sample' -StatusName 'txtStatus').Problems).Count | Should -Be 1
        }

        It 'passes a status element without a font in a Border' {
            $good = '<Grid ' + $script:FtSampleNs + '><Border Padding="8"><TextBlock x:Name="txtStatus" Text="Ready."/></Border></Grid>'
            $r = Get-FtFooterXamlProblem -Text $good -Source 'sample' -StatusName 'txtStatus'
            $r.Found | Should -BeTrue
            @($r.Problems) | Should -BeNullOrEmpty
        }

        It 'fails a font that code sets on the status variable' {
            $name = $script:FtConfig.StatusVariable[0]
            $bad = '$' + $name + '.FontSize = 11'
            @(Get-FtCodeProblem -Text $bad -Source 'sample').Count | Should -Be 1
            @(Get-FtCodeProblem -Text ('$' + $name + '.Text = "Ready"') -Source 'sample') | Should -BeNullOrEmpty
        }

        It 'fails an editable site code box outside the Options dialog and passes a read-only one' {
            $bad = '<StackPanel ' + $script:FtSampleNs + '><TextBox x:Name="txtSiteCode"/></StackPanel>'
            $bad2 = '<StackPanel ' + $script:FtSampleNs + '><ComboBox x:Name="cboSiteServer"/></StackPanel>'
            $good = '<StackPanel ' + $script:FtSampleNs + '><TextBox x:Name="txtSiteCode" IsReadOnly="True"/><TextBlock x:Name="txtSiteCodeInfo"/></StackPanel>'
            @(Get-FtSiteBoxXamlProblem -Text $bad -Source 'sample').Count | Should -Be 1
            @(Get-FtSiteBoxXamlProblem -Text $bad2 -Source 'sample').Count | Should -Be 1
            @(Get-FtSiteBoxXamlProblem -Text $good -Source 'sample') | Should -BeNullOrEmpty
        }

        It 'fails a site box name in code outside the Options functions and passes one inside' {
            $outside = 'function Show-Other { $t = New-Object System.Windows.Controls.TextBox; $t.Name = ''txtWizSiteCode'' }'
            @(Get-FtCodeProblem -Text $outside -Source 'sample').Count | Should -Be 1
            if ($script:FtConfig.OptionsFunction.Count -gt 0) {
                $inside = 'function ' + $script:FtConfig.OptionsFunction[0] + ' { $t = New-Object System.Windows.Controls.TextBox; $t.Name = ''txtWizSiteCode'' }'
                @(Get-FtCodeProblem -Text $inside -Source 'sample') | Should -BeNullOrEmpty
            }
        }

        It 'fails an editable site box in a XAML here-string outside the Options functions' {
            $nl = [Environment]::NewLine
            $bad = 'function Show-Other {' + $nl + '$x = @''' + $nl + '<StackPanel ' + $script:FtSampleNs + '><TextBox x:Name="txtProvider"/></StackPanel>' + $nl + '''@' + $nl + '}'
            @(Get-FtCodeProblem -Text $bad -Source 'sample').Count | Should -Be 1
        }

        It 'fails a start script that does not refresh the footer, and passes one that does' {
            $cfg = $script:FtConfig
            if ($cfg.NoSite) {
                $okText = '$window.Add_Loaded({ $script:FooterSite = Initialize-SuiteFooter -Window $window -StatusName ''txtStatus''; $script:FooterSite.Visibility = [System.Windows.Visibility]::Collapsed })'
                @(Get-FtWiringProblem -Text $okText -Source 'sample') | Should -BeNullOrEmpty
                @(Get-FtWiringProblem -Text $okText.Replace('Collapsed', 'Visible') -Source 'sample').Count | Should -Be 1
                @(Get-FtWiringProblem -Text '$window.Add_Loaded({ Write-Host 1 })' -Source 'sample').Count | Should -Be 1
                @(Get-FtWiringProblem -Text 'Initialize-SuiteFooter -Window $window' -Source 'sample').Count | Should -BeGreaterThan 0
                return
            }
            $loadedCall = 'Initialize-SuiteFooter -Window $window'
            $good = New-Object System.Text.StringBuilder
            [void]$good.AppendLine('function ' + $cfg.Wrapper + ' { Update-SuiteFooter -Window $window -SiteCode $a -SiteServer $b }')
            $hasLoaded = $false
            foreach ($u in $cfg.Update) {
                if ($u.Kind -eq 'Function') {
                    $pre = ''
                    if ($u.After) { $pre = $u.After + '; ' }
                    [void]$good.AppendLine('function ' + $u.Name + ' { ' + $pre + $cfg.Wrapper + ' }')
                }
                else {
                    $t = '$window'
                    if ($u.Target) { $t = $u.Target }
                    $body = ''
                    if ($u.Member -eq 'Add_Loaded') { $body = $loadedCall + '; '; $hasLoaded = $true }
                    if ($u.After) { $body += $u.After + '; ' }
                    [void]$good.AppendLine($t + '.' + $u.Member + '({ ' + $body + $cfg.Wrapper + ' })')
                }
            }
            if (-not $hasLoaded) { [void]$good.AppendLine('$window.Add_Loaded({ ' + $loadedCall + ' })') }
            $goodText = $good.ToString()
            @(Get-FtWiringProblem -Text $goodText -Source 'sample') | Should -BeNullOrEmpty
            @(Get-FtWiringProblem -Text '$window.Add_Loaded({ Write-Host 1 })' -Source 'sample').Count | Should -BeGreaterThan 2
            @(Get-FtWiringProblem -Text $goodText.Replace($loadedCall, 'Write-Host') -Source 'sample').Count | Should -BeGreaterThan 0
            @(Get-FtWiringProblem -Text $goodText.Replace(' ' + $cfg.Wrapper + ' }', ' Write-Host }') -Source 'sample').Count | Should -BeGreaterThan 0
        }
    }
}
