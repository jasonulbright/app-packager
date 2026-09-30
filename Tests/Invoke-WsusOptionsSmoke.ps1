#Requires -Version 5.1
# Options -> WSUS Publishing round trip: defaults, panel build, the SSL port
# follow, commit and reload, and the background-context settings per
# deployment target. No WSUS server is contacted and no dialog opens.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Packagers\AppPackagerWsus.psd1') -Force -Global

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('New-WsusPanel', 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-DeploymentTargetNames',
                    'Test-DeploymentTargetSkipsSite', 'Test-DeploymentTargetPublishesToWsus', 'Get-WsusClassificationNameList',
                    'Get-WsusClassificationLabel', 'Get-WsusPublishConfigForContext',
                    'Get-IntunePublishConfigForContext')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw ('function not found: ' + $name) }
    # Global, the way the shell publishes its functions: a GetNewClosure
    # handler runs in a module whose scope chain ends at global.
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$fixtureDir = Join-Path $env:TEMP ('ap-wsus-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
$script:fixturePath = Join-Path $fixtureDir 'preferences.json'
function global:Get-PreferencesPath { $script:fixturePath }

$checks = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:checks += 1
    Write-Host ('  ok  ' + $Message)
}
$click = { param($Control) $Control.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }

try {
    $script:Prefs = Read-Preferences
    $wsus = $script:Prefs.Wsus
    & $assert ($null -ne $wsus) 'preferences carry a Wsus block'
    & $assert ([string]$wsus.ServerName -eq '' -and [int]$wsus.PortNumber -eq 8530 -and -not [bool]$wsus.UseSsl) 'the server defaults to unset on HTTP port 8530'
    & $assert ([string]$wsus.Classification -eq 'Updates' -and -not $wsus.PSObject.Properties['PackageType']) 'publishing defaults to the Updates classification and stores no package type'
    & $assert ($null -eq (Get-WsusPublishConfigForContext)) 'a ConfigMgr One Click destination hands no WSUS settings to the background run'

    $panel = New-WsusPanel
    & $assert ([string]$panel.Name -eq 'WSUS Publishing') 'the panel registers under its own Options section'
    & $assert ($null -eq $panel.Element.FindName('cboWsusPackageType') -and [string]$panel.Element.FindName('txtWsusPolicy').Text -match 'older version') 'the panel offers no package type and states that updates need an older installed version'
    & $assert ($panel.Element.FindName('cboWsusClassification').Items.Count -eq @(Get-WsusClassificationNames).Count) 'every classification is offered'
    & $assert ($null -eq $panel.Element.FindName('btnWsusManage') -and $null -eq $panel.Element.FindName('btnWsusCatalog')) 'update management lives in the WSUS Updates window, not in Options'

    $api = Test-WsusAdministrationApi
    if (-not $api.Available) {
        & $assert (-not $panel.Element.FindName('btnWsusTest').IsEnabled) 'without the WSUS API the server buttons are disabled'
        & $assert ([string]$panel.Element.FindName('txtWsusStatus').Text -match 'not installed') 'the status line says the API is missing'
    }

    $ssl = $panel.Element.FindName('chkWsusSsl')
    $ssl.IsChecked = $true
    & $assert ([string]$panel.Element.FindName('txtWsusPort').Text -eq '8531') 'turning SSL on moves a default port to 8531, also without a mouse click'
    $ssl.IsChecked = $false
    & $assert ([string]$panel.Element.FindName('txtWsusPort').Text -eq '8530') 'turning SSL off moves the port back to 8530'
    $ssl.IsChecked = $true

    $panel.Element.FindName('txtWsusServer').Text = ' wsus01.contoso.com '
    $panel.Element.FindName('cboWsusClassification').SelectedIndex = @(Get-WsusClassificationNames).IndexOf('SecurityUpdates')
    $panel.Element.FindName('cboWsusGroup').Text = ' Pilot '
    $panel.Element.FindName('chkWsusDecline').IsChecked = $true
    & $panel.Commit
    $script:Prefs.Intune.DeploymentTarget = 'WSUSOnly'
    $script:Prefs | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:fixturePath -Encoding UTF8

    $script:Prefs = Read-Preferences
    $w = $script:Prefs.Wsus
    & $assert ([string]$w.ServerName -eq 'wsus01.contoso.com') 'the server name is trimmed and stored'
    & $assert ([int]$w.PortNumber -eq 8531 -and [bool]$w.UseSsl) 'SSL and its port survive save and reload'
    & $assert (-not ((Get-Content -LiteralPath $script:fixturePath -Raw | ConvertFrom-Json).Wsus.PSObject.Properties['PackageType'])) 'the saved file carries no package type'
    & $assert ([string]$w.Classification -eq 'SecurityUpdates') 'the classification survives'
    & $assert ([string]$w.ApprovalGroup -eq 'Pilot' -and [bool]$w.DeclineSuperseded) 'the approval group and the decline choice survive'

    $config = Get-WsusPublishConfigForContext
    & $assert ($config -is [hashtable] -and [string]$config.ServerName -eq 'wsus01.contoso.com' -and [int]$config.PortNumber -eq 8531 -and -not $config.ContainsKey('PackageType')) 'the WSUS-only target hands the saved settings, without a package type, to the background run'
    & $assert ($null -eq (Get-IntunePublishConfigForContext)) 'a WSUS target never hands out Intune credentials'
    $script:Prefs.Intune.DeploymentTarget = 'MECM'
    $buttonConfig = Get-WsusPublishConfigForContext -Target 'WSUSOnly'
    & $assert ($buttonConfig -is [hashtable] -and [string]$buttonConfig.ServerName -eq 'wsus01.contoso.com') 'the Publish to WSUS button hands the saved settings to its run whatever One Click publishes to'
    & $assert ($null -eq (Get-WsusPublishConfigForContext -Target 'MECM') -and $null -eq (Get-WsusPublishConfigForContext -Target 'IntuneOnly')) 'the ConfigMgr and Intune buttons hand no WSUS settings to their runs'
    & $assert ($null -eq (Get-IntunePublishConfigForContext -Target 'WSUSOnly')) 'the Publish to WSUS button never hands out Intune credentials'
    $script:Prefs.Intune.DeploymentTarget = 'WSUSOnly'

    # A stale Intune toggle and unknown values from a hand-edited file.
    $edited = Get-Content -LiteralPath $script:fixturePath -Raw | ConvertFrom-Json
    $edited.Intune.PublishToIntune = $true
    $edited.Wsus.Classification = 'Drivers'
    $edited.Wsus.PortNumber = 70000
    $edited.Wsus | Add-Member -NotePropertyName PackageType -NotePropertyValue 'Application' -Force
    $edited | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:fixturePath -Encoding UTF8
    $reloaded = Read-Preferences
    & $assert (-not [bool]$reloaded.Intune.PublishToIntune) 'a WSUS target clears a stale Intune publish toggle'
    & $assert ([string]$reloaded.Wsus.Classification -eq 'Updates' -and -not $reloaded.Wsus.PSObject.Properties['PackageType']) 'an unknown classification falls back and a stored retired package type is dropped'
    & $assert ([int]$reloaded.Wsus.PortNumber -eq 8530) 'an out-of-range port falls back to the default'
    & $assert ([string]$reloaded.Wsus.ServerName -eq 'wsus01.contoso.com') 'the valid values next to the bad ones are kept'

    # A bad port typed in the panel keeps the stored one.
    $script:Prefs = $reloaded
    $again = New-WsusPanel
    $again.Element.FindName('txtWsusPort').Text = 'abc'
    & $again.Commit
    & $assert ([int]$script:Prefs.Wsus.PortNumber -eq 8530) 'a non-numeric port does not replace the stored port'

    $script:Prefs.Wsus.ServerName = ''
    & $assert ($null -eq (Get-WsusPublishConfigForContext)) 'no server means no WSUS settings for the background run'

    # The dialogs only build when opened, so their XAML is loaded here
    # without showing a window.
    $dialogs = @{
        'Show-WsusPublishedUpdatesDialog' = @('gridUpdates', 'cboGroup', 'btnApprove', 'btnDecline', 'btnExpire', 'btnRemove', 'btnRefresh', 'chkOtherPublishers', 'btnCatalog')
        'Show-WsusCatalogImportDialog'    = @('txtInput', 'txtResults', 'btnImport', 'btnLoadFile', 'btnOpenCatalog', 'txtStatus')
        'Show-WsusPasswordPrompt'         = @('pwdValue', 'btnOk', 'btnCancel')
    }
    foreach ($dialogName in $dialogs.Keys) {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $dialogName }, $false)
        $xamlNode = $fn.Find({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -like '<Controls:MetroWindow*' }, $true)
        [xml]$xml = $xamlNode.Value
        $window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
        $missing = @($dialogs[$dialogName] | Where-Object { -not $window.FindName($_) })
        & $assert ($missing.Count -eq 0) ('{0} XAML loads with every control its handlers use' -f $dialogName)
        if ($dialogName -eq 'Show-WsusPublishedUpdatesDialog') {
            # A DockPanel clips the children it cannot fit instead of wrapping.
            # Measured without a constraint, because a constrained measure
            # clamps DesiredSize to the space offered and hides the overflow.
            $content = $window.Content
            $footer = $content.Children[$content.Children.Count - 1]
            $footer.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
            & $assert ($footer.DesiredSize.Width -le $window.MinWidth) ('the WSUS Updates button row fits at the minimum width ({0} of {1})' -f [int]$footer.DesiredSize.Width, [int]$window.MinWidth)
            # Fixed columns take their width first; the Title column gets the rest.
            $columns = $window.FindName('gridUpdates').Columns
            $fixed = ($columns | Where-Object { -not $_.Width.IsStar } | ForEach-Object { $_.Width.Value } | Measure-Object -Sum).Sum
            $title = $columns | Where-Object { $_.Width.IsStar } | Select-Object -First 1
            & $assert (($window.Width - 32 - $fixed) -ge $title.MinWidth -and $title.MinWidth -ge 280) ('the Title column keeps at least {0} px at the default width' -f [int]$title.MinWidth)
        }
    }

    # A host where the WSUS module failed to load still opens Options.
    Remove-Module AppPackagerWsus -Force
    $script:Prefs = Read-Preferences
    $orphan = New-WsusPanel
    & $assert (-not $orphan.Element.FindName('btnWsusTest').IsEnabled) 'without the WSUS module the panel builds with the server buttons disabled'
    & $assert ([string]$orphan.Element.FindName('txtWsusStatus').Text -match 'did not load') 'the status line says the WSUS module did not load'

    Write-Host ''
    Write-Host ('PASS: WSUS options round trip, {0} checks' -f $script:checks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
