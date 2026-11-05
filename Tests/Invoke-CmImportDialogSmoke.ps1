#Requires -Version 5.1
# Headless probe for Convert ConfigMgr App to Intune: the dialog and its two
# buttons, the checks that keep it open, the publish and package switches,
# and the sidebar button. ConfigMgr is not contacted and no window is shown.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('Show-CmImportDialog', 'Get-CmImportRunSpec', 'Get-CmImportOutcomeText')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw ('function not found: ' + $name) }
    # Global, the way the shell publishes its functions: a GetNewClosure
    # handler runs in a module whose scope chain ends at global.
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}
function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }

$checks = 0
$failures = @()
$assert = {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { $script:checks += 1; Write-Host ('PASS  ' + $Message) }
    else { $script:failures += $Message; Write-Host ('FAIL  ' + $Message) }
}
$click = { param($Control) $Control.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }
$script1 = 'C:\AP\Convert-CMApplicationToIntune.ps1'

# --- The window --------------------------------------------------------------
$result = Show-CmImportDialog -Owner $null -SiteCode 'MCM' -OutputRoot 'D:\Converted' -IntuneReady $true -ScriptPath $script1 -ToolPath 'C:\T\IntuneWinAppUtil.exe' -Probe {
    param($p)
    & $assert ($p.Site.Text -eq 'MCM' -and $p.Output.Text -eq 'D:\Converted') 'the site code and the output folder start from the saved values'
    & $assert ($p.App.Text -eq '') 'the application name starts empty'
    & $assert ([string]$p.Strategy.SelectedItem.Tag -eq 'PerDeploymentType') 'the first-match strategy is the default'
    & $assert (-not $p.Native.IsChecked -and -not $p.Untranslated.IsChecked -and -not $p.Package.IsChecked -and -not $p.Publish.IsChecked) 'no option starts checked'
    & $assert ($p.Publish.IsEnabled) 'publish is available when the Intune settings are saved'
    & $assert ($p.Preview.Content -eq 'Preview' -and $p.Convert.Content -eq 'Convert' -and $p.Cancel.Content -eq 'Cancel') 'the three buttons carry their labels'

    & $click $p.Preview
    & $assert ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -like 'Enter the exact name*' -and $p.State.Action -eq 'Cancel') 'Preview with no name asks for the name and keeps the window'
    $p.App.Text = 'Contoso Viewer'
    $p.Site.Text = 'MC'
    & $click $p.Convert
    & $assert ($p.Note.Text -like 'The site code is three*' -and $p.State.Action -eq 'Cancel') 'Convert with a short site code says so and keeps the window'
    $p.Site.Text = 'MCM'

    $p.Publish.IsChecked = $true
    & $assert ($p.Package.IsChecked -and -not $p.Package.IsEnabled) 'checking publish checks package and locks it'
    $p.Publish.IsChecked = $false
    & $assert ($p.Package.IsEnabled) 'unchecking publish frees package'

    $p.Strategy.SelectedIndex = 1
    & $assert ($p.StrategyHelp.Text -like 'Each app carries only its own requirements*') 'the group strategy changes the help line'
    $p.Native.IsChecked = $true
    & $click $p.Convert
}
& $assert ($result.Action -eq 'Convert' -and $result.Spec.Ok) 'Convert returns the action and the run spec'
& $assert ($result.Spec.Arguments -contains '-NativeDetectionOnly' -and $result.Spec.Arguments[($result.Spec.Arguments.IndexOf('-VariantStrategy') + 1)] -eq 'AssignByGroup') 'the spec carries the checked switch and the strategy'
& $assert ($result.Spec.Arguments -notcontains '-SkipContent') 'a conversion copies the content'
& $assert ($result.OutputRoot -eq 'D:\Converted') 'the result carries the output folder'

$result = Show-CmImportDialog -Owner $null -SiteCode 'MCM' -OutputRoot 'D:\Converted' -IntuneReady $true -ScriptPath $script1 -Probe {
    param($p)
    $p.App.Text = 'Contoso Viewer'
    $p.Package.IsChecked = $true
    & $click $p.Preview
}
& $assert ($result.Action -eq 'Preview' -and $result.Spec.Arguments -contains '-SkipContent' -and $result.Spec.Arguments -notcontains '-Package') 'Preview copies no content and builds nothing'

$result = Show-CmImportDialog -Owner $null -SiteCode 'MCM' -OutputRoot 'D:\Converted' -IntuneReady $true -ScriptPath $script1 -ToolPath 'C:\T\IntuneWinAppUtil.exe' -Probe {
    param($p)
    $p.App.Text = 'Contoso Viewer'
    $p.Publish.IsChecked = $true
    & $click $p.Convert
}
& $assert ($result.Spec.Arguments -contains '-Publish' -and $result.Spec.Arguments -contains '-UseSavedSettings' -and $result.Spec.Arguments -contains '-Package' -and $result.Spec.Arguments -contains 'C:\T\IntuneWinAppUtil.exe') 'publish passes package, the saved settings and the tool path'

$result = Show-CmImportDialog -Owner $null -SiteCode 'MCM' -OutputRoot 'D:\Converted' -IntuneReady $false -ScriptPath $script1 -Probe {
    param($p)
    & $assert (-not $p.Publish.IsEnabled) 'publish is off when the Intune settings are not saved'
    & $click $p.Cancel
}
& $assert ($result.Action -eq 'Cancel' -and $null -eq $result.Spec) 'Cancel returns no spec'

# --- The sidebar button in the main window XAML -------------------------------------
[xml]$main = Get-Content -LiteralPath (Join-Path $root 'MainWindow.xaml') -Raw
$ns = New-Object System.Xml.XmlNamespaceManager($main.NameTable)
$ns.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
$ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
$button = $main.SelectSingleNode("//p:Button[@x:Name='btnConvertCm']", $ns)
& $assert ($null -ne $button -and $button.FirstChild.GetAttribute('Text') -eq 'Convert ConfigMgr App to Intune') 'the main window has the sidebar button'

Write-Host ("`n{0} passed, {1} failed." -f $checks, $failures.Count)
exit ([int]($failures.Count -gt 0))
