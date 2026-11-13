#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard test for the weight of static labels. It parses the XAML files, the XAML here-strings of the scripts
# and the run-time FontWeight assignments of the scripts. It runs on Windows PowerShell 5.1 and PowerShell 7.
# Only two kinds of text are bold:
#   (1) a field title: the text label of an input control (Site Code, SMS Provider, Name, and so on). In XAML a field title is a
#       TextBlock or Label that sits in the same grid row next to an input (or in the row above it) or that is the sibling before
#       or after an input in a panel. In code it is a TextBlock whose text ends in a colon, or one that the helper settings list.
#   (2) an app title: the x:Name of the settings (txtAppTitle in the sidebar).
# A page header, a section header, a status text, an about text and a log header keep their weight from before the theme change.

BeforeDiscovery {
    . "$PSScriptRoot\FontWeightGuardHelpers.ps1"
    $script:FwDocsD = @(Get-FwXamlDocument)
    $script:FwScriptsD = @(Get-FwScriptFile | ForEach-Object { @{ Name = (Get-FwRelativePath $_.FullName); Path = $_.FullName } })
    $script:FwStaticD = @($script:FwConfig.StaticName | ForEach-Object { @{ Name = $_ } })
}

BeforeAll {
    . "$PSScriptRoot\FontWeightGuardHelpers.ps1"
    $script:FwDocs = @(Get-FwXamlDocument)
}

Describe 'Static labels are not bold' {
    It 'finds the XAML of the app (the scan is not empty)' {
        $script:FwDocs.Count | Should -BeGreaterThan 0
        ($script:FwDocs | ForEach-Object { $_.Source }) -join ',' | Should -Match 'xaml|here-string'
    }

    It 'finds text labels and field titles in the XAML (the scan is not empty)' {
        $labels = 0; $titles = 0
        foreach ($d in $script:FwDocs) {
            $c = Get-FwLabelCount -Text $d.Text
            $labels += $c.Labels; $titles += $c.Titles
        }
        $labels | Should -BeGreaterThan 5
        $titles | Should -BeGreaterOrEqual 0
    }

    It 'finds the script files of the app (the scan is not empty)' {
        @(Get-FwScriptFile).Count | Should -BeGreaterThan 0
    }

    It 'has no bold style or element that is not a field title or an app title in <Source>' -ForEach $script:FwDocsD {
        @(Get-FwXamlViolation -Text $Text -Source $Source) | Should -BeNullOrEmpty
    }

    It 'sets no bold weight at run time on a label that is not a field title in <Name>' -ForEach $script:FwScriptsD {
        @(Get-FwCodeViolation -Path $Path -Label $Name) | Should -BeNullOrEmpty
    }

    It 'finds the element <Name> in the XAML and keeps it out of bold' -ForEach $script:FwStaticD {
        $found = $false
        foreach ($d in $script:FwDocs) {
            $s = Get-FwNamedElementState -Text $d.Text -Name $Name
            if ($s.Found) {
                $found = $true
                $s.Bold | Should -BeFalse -Because ('{0} in {1}' -f $Name, $d.Source)
            }
        }
        $found | Should -BeTrue -Because ('{0} must exist, or the settings list is out of date' -f $Name)
    }

    It 'keeps the app title bold in the XAML that names it' {
        foreach ($name in $script:FwConfig.AppTitleName) {
            $hits = @($script:FwDocs | ForEach-Object { Get-FwNamedElementState -Text $_.Text -Name $name } | Where-Object { $_.Found })
            $hits.Count | Should -BeGreaterThan 0 -Because $name
        }
    }
}

Describe 'The font weight check itself' {
    BeforeAll {
        $script:BadHeader = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <StackPanel>
    <TextBlock x:Name="txtHeaderSample" Text="Page header" FontSize="18" FontWeight="Bold"/>
    <TextBlock x:Name="txtStatusSample" Text="Ready" FontWeight="Bold"/>
  </StackPanel>
</Grid>
'@
        $script:GoodSample = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
  <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Name" FontWeight="Bold"/>
  <TextBox Grid.Row="0" Grid.Column="1"/>
  <TextBlock Grid.Row="1" Grid.Column="0" Text="Collection:" FontWeight="Bold"/>
  <ComboBox Grid.Row="1" Grid.Column="1"/>
  <TextBlock x:Name="txtModuleTitle" Text="Page header" FontSize="18" FontWeight="SemiBold"/>
</Grid>
'@
        $script:GoodSample = $script:GoodSample.Replace('txtModuleTitle', $script:FwConfig.StaticName[0])
        $script:StyleSample = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style x:Key="HeaderStyle" TargetType="TextBlock"><Setter Property="FontWeight" Value="Bold"/></Style>
</ResourceDictionary>
'@
    }

    It 'fails a bold page header and a bold status text (a bad sample)' {
        $v = @(Get-FwXamlViolation -Text $script:BadHeader -Source 'sample')
        $v.Count | Should -Be 2
        ($v -join ' ') | Should -Match 'txtHeaderSample'
        ($v -join ' ') | Should -Match 'txtStatusSample'
    }

    It 'fails a bold style that no field title uses (a bad sample)' {
        @(Get-FwXamlViolation -Text $script:StyleSample -Source 'sample').Count | Should -Be 1
    }

    It 'fails a bold key label that has a text value and no input (a bad sample)' {
        $text = $script:GoodSample.Replace('<ComboBox Grid.Row="1" Grid.Column="1"/>', '<TextBlock Grid.Row="1" Grid.Column="1" Text="value"/>')
        @(Get-FwXamlViolation -Text $text -Source 'sample').Count | Should -Be 1
    }

    It 'passes a bold field title next to an input and a SemiBold header' {
        @(Get-FwXamlViolation -Text $script:GoodSample -Source 'sample') | Should -BeNullOrEmpty
    }

    It 'fails a bold element that the settings name as a static element, also next to an input' {
        $text = $script:GoodSample.Replace('Text="Page header" FontSize="18" FontWeight="SemiBold"', 'Text="Page header" FontWeight="Bold"')
        @(Get-FwXamlViolation -Text $text -Source 'sample').Count | Should -Be 1
        $named = '<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="txtModuleHeader" FontWeight="Bold"/>'
        $s = Get-FwNamedElementState -Text $named -Name 'txtModuleHeader'
        $s.Found | Should -BeTrue
        $s.Bold | Should -BeTrue
    }

    It 'fails a run-time bold header and passes a run-time bold field title' {
        $bad = '$hdr = New-Object System.Windows.Controls.TextBlock' + "`n" + '$hdr.Text = ''About''' + "`n" + '$hdr.FontWeight = ''Bold'''
        $good = '$lbl = New-Object System.Windows.Controls.TextBlock' + "`n" + '$lbl.Text = ''Name:''' + "`n" + '$lbl.FontWeight = ''Bold'''
        @(Get-FwCodeViolation -ScriptText $bad -Label 'sample').Count | Should -Be 1
        @(Get-FwCodeViolation -ScriptText $good -Label 'sample') | Should -BeNullOrEmpty
    }

    It 'fails a run-time bold weight in a form that the scan does not parse' {
        $odd = '$tb.SetValue([System.Windows.Controls.TextBlock]::FontWeightProperty, [System.Windows.FontWeights]::Bold)'
        @(Get-FwCodeViolation -ScriptText $odd -Label 'sample').Count | Should -Be 1
    }

    It 'fails a call that asks for a bold header and passes one for a field title' {
        $bad = 'New-Label ''About'' 18 -Bold'
        $badWeight = 'New-Label ''About'' 18 -Weight Bold'
        $good = 'New-Label ''Name:'' 14 -Bold'
        @(Get-FwCodeViolation -ScriptText $bad -Label 'sample').Count | Should -Be 1
        @(Get-FwCodeViolation -ScriptText $badWeight -Label 'sample').Count | Should -Be 1
        @(Get-FwCodeViolation -ScriptText $good -Label 'sample') | Should -BeNullOrEmpty
    }
}
