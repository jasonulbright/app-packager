#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every grid column is wide enough for its header.

.DESCRIPTION
    The grid header style writes the header in capitals, bold, and reserves
    room for padding and the sort arrow. A column with a fixed width and no
    MinWidth cut the header (CURRENT (CONFIGM, MAX CVS, IN LA, NEWES). The
    check measures the capitals with the header font and adds the room the
    style keeps.
#>

BeforeAll {
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
    $script:HeaderRoom = 55
    $xaml = New-Object System.Xml.XmlDocument
    $xaml.Load((Join-Path (Split-Path -Parent $PSScriptRoot) 'MainWindow.xaml'))
    $ns = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
    $ns.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    $script:Columns = @($xaml.SelectNodes('//p:DataGrid[@x:Name="dataGrid"]/p:DataGrid.Columns/*', $ns) | ForEach-Object { $_ })
    if ($script:Columns.Count -eq 0) {
        $script:Columns = @($xaml.SelectNodes('//p:DataGridTextColumn | //p:DataGridTemplateColumn', $ns) | ForEach-Object { $_ })
    }
    $script:Measure = {
        param([string]$Text)
        $face = New-Object System.Windows.Media.Typeface((New-Object System.Windows.Media.FontFamily('Segoe UI')), [System.Windows.FontStyles]::Normal, [System.Windows.FontWeights]::Bold, [System.Windows.FontStretches]::Normal)
        $ft = New-Object System.Windows.Media.FormattedText($Text, [cultureinfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $face, 13, [System.Windows.Media.Brushes]::Black, 1.0)
        $ft.WidthIncludingTrailingWhitespace
    }
}

Describe 'Grid column headers' {
    It 'finds the grid columns' {
        @($script:Columns | Where-Object { $_.GetAttribute('Header') }).Count | Should -BeGreaterThan 8
    }

    It 'gives every titled column a MinWidth that holds its header in capitals' {
        $short = @()
        foreach ($c in $script:Columns) {
            $header = $c.GetAttribute('Header')
            if ([string]::IsNullOrWhiteSpace($header)) { continue }
            $need = [math]::Ceiling((& $script:Measure $header.ToUpperInvariant()) + $script:HeaderRoom)
            $min = 0.0
            [void][double]::TryParse($c.GetAttribute('MinWidth'), [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$min)
            $width = 0.0
            $isFixed = [double]::TryParse($c.GetAttribute('Width'), [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$width)
            if ($min -lt $need) { $short += ('{0}: MinWidth {1}, needs {2}' -f $header, $min, $need) }
            elseif ($isFixed -and $width -lt $min) { $short += ('{0}: Width {1} is below MinWidth {2}' -f $header, $width, $min) }
        }
        $short | Should -BeNullOrEmpty
    }
}

Describe 'Grid column headers in the dialogs of the start script' {
    BeforeAll {
        # The dialog grids are written as XAML text inside start-apppackager.ps1. Read each column tag that has a literal header.
        # The header style of a dialog grid keeps 12 on each side and 16 for the sort arrow: 40 beside the text. The main grid keeps more.
        $script:DialogHeaderRoom = 40
        $script:ScriptText = [IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1'))
        $script:DialogColumns = @([regex]::Matches($script:ScriptText, '<DataGrid(?:Text|Template)Column\b[^>]*>') | ForEach-Object { $_.Value } | Where-Object { $_ -match '\sHeader="[^"{][^"]*"' })
        function Get-ColumnAttribute([string]$Tag, [string]$Name) { $m = [regex]::Match($Tag, '\s' + $Name + '="([^"]*)"'); if ($m.Success) { $m.Groups[1].Value } else { '' } }
    }

    It 'finds the column tags of the dialogs' {
        $script:DialogColumns.Count | Should -BeGreaterThan 10
    }

    It 'gives every titled column of a dialog grid a MinWidth that holds its header in capitals' {
        $short = @()
        foreach ($tag in $script:DialogColumns) {
            $header = Get-ColumnAttribute $tag 'Header'
            if ([string]::IsNullOrWhiteSpace($header)) { continue }
            $need = [math]::Ceiling((& $script:Measure $header.ToUpperInvariant()) + $script:DialogHeaderRoom)
            $min = 0.0
            [void][double]::TryParse((Get-ColumnAttribute $tag 'MinWidth'), [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$min)
            $width = 0.0
            $isFixed = [double]::TryParse((Get-ColumnAttribute $tag 'Width'), [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$width)
            if ($min -lt $need) { $short += ('{0}: MinWidth {1}, needs {2}' -f $header, $min, $need) }
            elseif ($isFixed -and $width -lt $min) { $short += ('{0}: Width {1} is below MinWidth {2}' -f $header, $width, $min) }
        }
        $short | Should -BeNullOrEmpty
    }
}
