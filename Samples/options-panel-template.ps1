# Options panel factory template.
#
# Copy this function body into start-apppackager.ps1 near the other
# New-*Panel functions, rename New-MyPanel -> New-<YourThing>Panel,
# and wire it into the $panels array in Show-OptionsDialog.
#
# This file is a standalone template and is NOT imported by the GUI.
# Read Samples/OPTIONS_AUTHORING.md for the full walkthrough, including
# the three closure / scope traps that will bite you.

function New-MyPanel {
    # ---------------------------------------------------------------
    # XAML - panel body only. No MetroWindow wrapper; no OK / Cancel
    # (the master Options dialog owns those).
    #
    # Critical: avoid BasedOn="{StaticResource ...}". The panel is
    # loaded standalone before it is attached to the master's visual
    # tree, so StaticResource lookups fail at XamlReader.Load time.
    # Use DynamicResource for brushes, or drop BasedOn entirely.
    # ---------------------------------------------------------------
    $xaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="140"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Text="Foo:"
               FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"
               ToolTip="Explain what Foo does"/>
    <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtFoo"
               FontSize="13" MaxLength="200" Margin="0,0,0,8"/>

    <TextBlock Grid.Row="1" Grid.Column="0" Text="Bar:"
               FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <CheckBox  Grid.Row="1" Grid.Column="1" x:Name="chkBar"
               Content="Enable Bar" FontSize="13" VerticalAlignment="Center" Margin="0,0,0,8"
               Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

    <TextBlock Grid.Row="2" Grid.Column="0" Text="Baz:"
               FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,0,8"/>
    <ComboBox  Grid.Row="2" Grid.Column="1" x:Name="cmbBaz"
               FontSize="13" Width="180" HorizontalAlignment="Left">
        <ComboBoxItem Content="Option A"/>
        <ComboBoxItem Content="Option B"/>
        <ComboBoxItem Content="Option C"/>
    </ComboBox>
</Grid>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    # ---------------------------------------------------------------
    # FindName each named control into a local variable.
    # ---------------------------------------------------------------
    $txtFoo = $element.FindName('txtFoo')
    $chkBar = $element.FindName('chkBar')
    $cmbBaz = $element.FindName('cmbBaz')

    # ---------------------------------------------------------------
    # Pre-populate from current prefs.
    #
    # Be sure $script:Prefs.MySection exists. Add the defaults to
    # Read-Preferences's $defaults block + the merge block so the
    # first-run case has something to bind to.
    # ---------------------------------------------------------------
    $txtFoo.Text      = [string]$script:Prefs.MySection.Foo
    $chkBar.IsChecked = [bool]  $script:Prefs.MySection.Bar

    $bazMap = @{ 'A' = 0; 'B' = 1; 'C' = 2 }
    $idx = $bazMap[[string]$script:Prefs.MySection.Baz]
    if ($null -eq $idx) { $idx = 0 }
    $cmbBaz.SelectedIndex = $idx

    # ---------------------------------------------------------------
    # Commit closure.
    #
    # Capture $script:Prefs once as a LOCAL variable ($prefsRef). The
    # closure will be invoked later from the master OK handler, and
    # $script: scope + function lookup are not reachable from inside a
    # GetNewClosure()d scriptblock. Mutate captured refs only.
    # ---------------------------------------------------------------
    $prefsRef = $script:Prefs
    $commit = {
        $prefsRef.MySection.Foo = $txtFoo.Text.Trim()
        $prefsRef.MySection.Bar = [bool]$chkBar.IsChecked

        $bazReverse = @{ 0 = 'A'; 1 = 'B'; 2 = 'C' }
        $selBaz = $bazReverse[$cmbBaz.SelectedIndex]
        if (-not $selBaz) { $selBaz = 'A' }
        $prefsRef.MySection.Baz = $selBaz
    }.GetNewClosure()

    return @{
        Name    = 'My Panel'
        Element = $element
        Commit  = $commit
    }
}
