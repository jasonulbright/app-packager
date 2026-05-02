# Authoring an Options panel

The Options window follows a master/detail pattern: a left ListBox picks a panel name, a right ContentControl renders the panel body, and a single OK commits every panel in one action. Adding a fifth (or fifteenth) panel is a three-step change in `start-apppackager.ps1`.

## Architecture

Each panel is built by a **factory function** that returns a hashtable with a fixed shape:

```powershell
@{
    Name        = 'My Panel'             # shown in the nav list
    Element     = $wpfElement            # WPF UIElement rendered on select
    Commit      = { ... }.GetNewClosure()   # mutates captured state, called on master OK
    CwaSwitches = $optional                # optional: sibling JSON configs to save after Save-Preferences
    TvConfig    = $optional
}
```

Four factories ship today:

- `New-MecmPreferencesPanel`
- `New-PackagerPreferencesPanel`
- `New-AppFlowPanel` (the "One Click Settings" panel)
- `New-ProductFilterPanel`

All four follow the same shape. Copy any one as a starting point. `New-MecmPreferencesPanel` is the simplest - flat form, no sub-dialogs.

The master `Show-OptionsDialog` hosts the ListBox + ContentControl + OK/Cancel. It builds the panel array at dialog open:

```powershell
$panels = @(
    (New-MecmPreferencesPanel),
    (New-PackagerPreferencesPanel),
    (New-AppFlowPanel),
    (New-ProductFilterPanel)
)
```

On OK, it runs each panel's `Commit`, then `Save-Preferences`, then (for panels that expose sibling JSON configs) their save helpers, then `Invoke-RefreshGrid`.

## Preferences JSON schema

Panels read and write `$script:Prefs`, a `pscustomobject` that round-trips through `AppPackager.preferences.json` via `Read-Preferences` / `Save-Preferences`. To persist a new setting:

1. Add the default to the `$defaults` pscustomobject in `Read-Preferences`. Nest related settings inside a sub-object (like `AppFlow`, `ContentDistribution`, `DetectedTools`).
2. Add a merge block below that reads `$data.YourSection` from disk, validates each field, and copies into `$defaults.YourSection`. Bad values fall back to defaults silently.
3. `Save-Preferences` already serializes the whole `$Prefs` at `-Depth 5`; no change needed unless your shape needs deeper nesting.

Inline configs (Citrix switches, TeamViewer Host) that are too complex for the main prefs file live in sibling JSON under `Packagers/` and have dedicated `Read-<Thing>` / `Save-<Thing>` helpers in the shared module. Follow that pattern for similarly-sized state.

## Panel factory anatomy

A minimal factory looks like this (annotated):

```powershell
function New-MyPanel {
    # 1. XAML shell. Panel body only - no MetroWindow, no OK/Cancel
    # (the master owns those). No BasedOn="{StaticResource ...}" either -
    # StaticResource resolves at XamlReader.Load time, and the panel is
    # loaded standalone before it is attached to the master's visual tree,
    # so static-resource references fail. Use DynamicResource for brushes.
    $xaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    ...
    <TextBox x:Name="txtFoo"/>
    <CheckBox x:Name="chkBar"/>
    ...
</Grid>
'@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $element = [System.Windows.Markup.XamlReader]::Load($reader)

    # 2. FindName each x:Name-bearing control into a local variable.
    $txtFoo = $element.FindName('txtFoo')
    $chkBar = $element.FindName('chkBar')

    # 3. Pre-populate controls from current prefs.
    $txtFoo.Text     = [string]$script:Prefs.MySection.Foo
    $chkBar.IsChecked = [bool]$script:Prefs.MySection.Bar

    # 4. Commit closure. Captures the panel-local control refs + a local
    #    ref to $script:Prefs. Uses .GetNewClosure() because the handler
    #    fires from the master OK click AFTER the factory has returned,
    #    so local scope needs to be snapshotted.
    #
    #    Inside GetNewClosure() you CANNOT call script-scope functions or
    #    read $script:Foo directly (PowerShell strips function-table
    #    visibility and $script: scope inside closures). Only touch
    #    captured locals. If you need to call a module function, either
    #    expose the state on the returned hashtable and let the master
    #    handle it (the CwaSwitches/TvConfig pattern), or use a $script:
    #    scriptblock variable and invoke via & $script:MyFunc.
    $prefsRef = $script:Prefs
    $commit = {
        $prefsRef.MySection.Foo = $txtFoo.Text.Trim()
        $prefsRef.MySection.Bar = [bool]$chkBar.IsChecked
    }.GetNewClosure()

    return @{ Name = 'My Panel'; Element = $element; Commit = $commit }
}
```

## Wiring the panel into the master

Add one line to the `$panels` array inside `Show-OptionsDialog`:

```powershell
$panels = @(
    (New-MecmPreferencesPanel),
    (New-PackagerPreferencesPanel),
    (New-AppFlowPanel),
    (New-ProductFilterPanel),
    (New-MyPanel)                     # <- insert here
)
```

No other code change is needed. The ListBox auto-populates from `$panels[i].Name`; the ContentControl swaps on selection; OK runs every `Commit` in order.

## Three things that will bite you

**1. StaticResource + BasedOn fails at load time.** Panels load standalone, without the parent window's merged resource dictionaries. `<Style TargetType="..." BasedOn="{StaticResource MahApps.Styles.DataGridCell}">` throws at `XamlReader.Load`. Workarounds: use `DynamicResource` (for brushes), drop the `BasedOn` (accept plain WPF defaults, or inline the setters you need), or import a dictionary inside `<Panel.Resources>`. `DynamicResource` is the usual right answer.

**2. `.GetNewClosure()` strips script scope and function lookup.** Variables in the closure are captured as panel-local snapshots. Script-scope `$script:Foo` and script-scope functions like `Save-Preferences` become unreachable from inside the closure. Mutate captured references only; let the master OK handler do the cross-module calls.

**3. Inner functions do not inherit outer scope via GetNewClosure.** If you define a helper function inside the factory, its captured variables come from the helper's scope, not the factory's. Pass any outer-scope dependency in as an explicit parameter. Mistake version (broken):

```powershell
function New-MyPanel {
    $ownerRef = $someOwner
    function Do-Thing {
        $ownerRef.DoStuff()   # $ownerRef is null at call time
    }
    $btn.Add_Click({ Do-Thing }.GetNewClosure())
    ...
}
```

Fixed version:

```powershell
function New-MyPanel {
    $ownerRef = $someOwner
    function Do-Thing {
        param($Owner)
        $Owner.DoStuff()
    }
    $btn.Add_Click({ Do-Thing -Owner $ownerRef }.GetNewClosure())
    ...
}
```

## Where to put owner-specific things

Most panels never need to open sub-dialogs. If yours does (like Packager Preferences opening the CWA / M365 previews), the owner window is the master Options dialog. You do not have a direct reference to it at factory-build time, but you can get it from the panel element at click time:

```powershell
$btn.Add_Click({
    $ownerWin = [System.Windows.Window]::GetWindow($element)
    Show-PreviewDialog -Owner $ownerWin -Title "Preview" -Content $stuff
}.GetNewClosure())
```

`GetWindow` walks the visual tree to find the enclosing window, which is the master Options dialog once the panel is attached.

## Saving to sibling JSON configs

If your panel mutates state that lives outside `AppPackager.preferences.json` - its own JSON file, a registry key, etc. - expose the mutated state on the returned hashtable:

```powershell
$myState = Read-MyState
...
$commit = {
    $myState.Foo = $txtFoo.Text.Trim()
    $myState.Bar = [bool]$chkBar.IsChecked
}.GetNewClosure()

return @{
    Name    = 'My Panel'
    Element = $element
    Commit  = $commit
    MyState = $myState                # <- expose for master to save
}
```

Then teach the master OK handler to recognize your property and call its save helper. Open `Show-OptionsDialog`, find the block that handles `p.CwaSwitches` and `p.TvConfig`, and add a case for `p.MyState`. That keeps the save-helper function call in the non-closure'd master handler where function resolution works normally.

## Testing a new panel

1. Parse-check `start-apppackager.ps1` after your edits:
   ```powershell
   $null = [System.Management.Automation.Language.Parser]::ParseFile('start-apppackager.ps1', [ref]$null, [ref]$errors)
   ```
2. Headless-load the panel's XAML to catch brush/style issues before GUI launch (see the pattern the existing AppPackager test harness uses).
3. Launch the GUI, open Options, select your panel. Verify all controls render.
4. Change values, click OK. Check that changes landed in `AppPackager.preferences.json` (and any sibling configs).
5. Relaunch the GUI - your new defaults should load back in from disk.
