#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the tooltips (Q-103). They parse the files, so they run on Windows PowerShell 5.1 and PowerShell 7.
#   (a) Every named control of every window, dialog and panel has a tooltip, except the controls of the exempt list.
#   (b) Every column header in the tooltip table of the start script exists, and its tip is one short sentence.

BeforeDiscovery {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
}

BeforeAll {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocs = @(Get-XamlDocument)
    $script:ControlType = 'Button', 'CheckBox', 'ToggleButton', 'RadioButton', 'TextBox', 'ComboBox', 'PasswordBox', 'DataGrid', 'ListBox', 'ListView', 'TreeView', 'ToggleSwitch', 'Slider', 'NumericUpDown'

    # A control that explains itself by its label, or that holds only line numbers, needs no tooltip.
    $script:Exempt = @{
        btnClose        = 'The label says what it does.'
        btnCancel       = 'The label says what it does.'
        btnOk           = 'The label says what it does.'
        btnContinue     = 'The label says what it does.'
        btnNameOk       = 'The label says what it does.'
        btnNameCancel   = 'The label says what it does.'
        btnDropCancel   = 'The label says what it does.'
        btnCmCancel     = 'The label says what it does.'
        btnReviewCancel = 'The label says what it does.'
        btnUnsavedCancel = 'The label says what it does.'
        txtNameValue    = 'The prompt above the box says what to type.'
        txtScriptGutter = 'Line numbers of the script. It takes no input.'
        txtDetectGutter = 'Line numbers of the detection script. It takes no input.'
    }

    function Get-ControlWithoutTip {
        param($Doc)
        [xml]$x = $Doc.Text
        foreach ($el in @($x.SelectNodes('//*'))) {
            if ($script:ControlType -notcontains $el.LocalName) { continue }
            $name = $el.GetAttribute('Name', $script:XamlKeyNamespace)
            if (-not $name) { continue }
            if ($el.HasAttribute('ToolTip')) { continue }
            $property = $el.LocalName + '.ToolTip'
            if (@($el.ChildNodes | Where-Object { $_.LocalName -eq $property }).Count -gt 0) { continue }
            [pscustomobject]@{ Source = $Doc.Source; Name = $name; Type = $el.LocalName }
        }
    }
}

Describe 'Every named control has a tooltip' {
    It 'finds the documents' {
        $script:XamlDocs.Count | Should -BeGreaterOrEqual 20
    }

    It 'has no named control without a tooltip, except the exempt list' {
        $missing = foreach ($doc in $script:XamlDocs) { Get-ControlWithoutTip -Doc $doc | Where-Object { -not $script:Exempt.ContainsKey($_.Name) } }
        @($missing | ForEach-Object { '{0}: {1} {2}' -f $_.Source, $_.Type, $_.Name }) | Should -BeNullOrEmpty
    }

    It 'keeps only exempt controls that still exist without a tooltip' {
        $seen = @{}
        foreach ($doc in $script:XamlDocs) { foreach ($c in @(Get-ControlWithoutTip -Doc $doc)) { $seen[$c.Name] = $true } }
        $stale = @($script:Exempt.Keys | Where-Object { -not $seen.ContainsKey($_) })
        $stale | Should -BeNullOrEmpty
    }

    It 'finds a control without a tooltip (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
        $doc = [pscustomobject]@{ Source = 'probe'; Text = "<StackPanel $ns><Button x:Name=`"a`"/><Button x:Name=`"b`" ToolTip=`"Tip.`"/><TextBox x:Name=`"c`"><TextBox.ToolTip>Tip.</TextBox.ToolTip></TextBox></StackPanel>" }
        @(Get-ControlWithoutTip -Doc $doc | ForEach-Object { $_.Name }) | Should -Be @('a')
    }
}

Describe 'The column header tips' {
    BeforeAll {
        $path = Join-Path $script:RepoRoot 'start-apppackager.ps1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        $assign = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$script:GridHeaderTip' }, $true))
        $script:HeaderTip = & ([scriptblock]::Create($assign[0].Right.Extent.Text))
    }

    It 'has a table in the start script' {
        $script:HeaderTip.Keys.Count | Should -BeGreaterOrEqual 10
    }

    It 'names only grids and headers that exist' {
        $bad = foreach ($grid in $script:HeaderTip.Keys) {
            $headers = @()
            foreach ($doc in $script:XamlDocs) {
                [xml]$x = $doc.Text
                foreach ($g in @($x.SelectNodes("//*[local-name()='DataGrid']") | Where-Object { $_.GetAttribute('Name', $script:XamlKeyNamespace) -eq $grid })) {
                    $headers += @($g.SelectNodes(".//*[starts-with(local-name(), 'DataGrid') and @Header]") | ForEach-Object { $_.GetAttribute('Header') })
                }
            }
            if ($headers.Count -eq 0) { 'grid {0} not found' -f $grid; continue }
            foreach ($h in $script:HeaderTip[$grid].Keys) { if ($headers -notcontains $h) { '{0}: header {1} not found' -f $grid, $h } }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives every header tip one short sentence' {
        $bad = foreach ($grid in $script:HeaderTip.Keys) { foreach ($h in $script:HeaderTip[$grid].Keys) { if (-not (Test-OneSentence $script:HeaderTip[$grid][$h])) { '{0}/{1}' -f $grid, $h } } }
        @($bad) | Should -BeNullOrEmpty
    }
}
