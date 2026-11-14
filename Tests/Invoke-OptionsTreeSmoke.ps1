#Requires -Version 5.1
# Window probe of the Options window: every section of the tree opens, the search finds a field and
# jumps to it, a save and a reload keep every setting, the SQL server name splits into its fields, the
# imported Microsoft 365 configuration.xml survives save and reload, and the Setup wizard builds
# its fields from the same definitions as the Connections sections. No server is contacted.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixtureDir = Join-Path $env:TEMP ('ap-options-tree-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
. (Join-Path $PSScriptRoot 'OptionsTestSupport.ps1') -Root $root -FixtureDir $fixtureDir

$script:PassedChecks = 0
$script:LayoutChecked = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:PassedChecks += 1
    Write-Host ('  ok  ' + $Message)
}

# The text boxes, password boxes and combo boxes of a page, found through the logical tree.
function Get-OptionsInputControl {
    param([Parameter(Mandatory)]$Element)
    foreach ($child in @([System.Windows.LogicalTreeHelper]::GetChildren($Element))) {
        if ($child -is [System.Windows.Controls.DataGrid] -or $child -is [System.Windows.Controls.TreeView]) { continue }
        if ($child -is [System.Windows.Controls.TextBox] -or $child -is [System.Windows.Controls.PasswordBox] -or $child -is [System.Windows.Controls.ComboBox]) { $child }
        if ($child -is [System.Windows.DependencyObject]) { Get-OptionsInputControl -Element $child }
    }
}

# Lays the element out at a width and fails on any input that is cut off at the right edge or is narrower than its floor.
function Test-OptionsLayout {
    param([Parameter(Mandatory)]$Element, [object[]]$Fields = @(), [Parameter(Mandatory)][string]$Label, [double]$NativeWidth = 900, [double]$SmallestWidth = 560)
    foreach ($width in @($NativeWidth, $SmallestWidth)) {
        $Element.Measure((New-Object System.Windows.Size($width, 6000)))
        $Element.Arrange((New-Object System.Windows.Rect(0, 0, $width, [Math]::Max($Element.DesiredSize.Height, 10))))
        $Element.UpdateLayout()
        foreach ($c in @(Get-OptionsInputControl -Element $Element)) {
            if ($c.ActualWidth -le 0) { continue }
            $script:LayoutChecked += 1
            $id = [System.Windows.Automation.AutomationProperties]::GetAutomationId($c)
            $box = $c.TransformToAncestor($Element).TransformBounds((New-Object System.Windows.Rect(0, 0, $c.ActualWidth, $c.ActualHeight)))
            if ($box.Right -gt $width + 0.5) { throw ('FAIL: {0} / {1} ends at {2} in a window with {3} of width' -f $Label, $id, [int]$box.Right, [int]$width) }
            $field = @($Fields | Where-Object { $_ -and [string]$_.Id -eq $id }) | Select-Object -First 1
            if ($field -and $width -eq $NativeWidth) {
                $floor = Get-OptionsInputFloor -Field $field
                if ($c.ActualWidth + 0.5 -lt $floor) { throw ('FAIL: {0} / {1} is {2} wide at the native size; the floor is {3}' -f $Label, $id, [int]$c.ActualWidth, $floor) }
            }
            if (-not $field -and $width -eq $NativeWidth -and $c -is [System.Windows.Controls.TextBox] -and $c.ActualWidth + 0.5 -lt 360) { throw ('FAIL: {0} / a text box without a field definition is {1} wide at the native size' -f $Label, [int]$c.ActualWidth) }
            if ($c.ActualWidth + 0.5 -lt $c.MinWidth) { throw ('FAIL: {0} / {1} is narrower than its minimum' -f $Label, $id) }
        }
    }
}

# A preference file with a non-default value in every setting that the Options window shows.
$fullPrefs = @'
{
  "SiteCode": "ABC",
  "ProviderMachineName": "cm01.contoso.com",
  "FileShareRoot": "\\\\cm01\\sources$",
  "ContentLayout": "Flat",
  "DownloadRoot": "D:\\ap",
  "EstimatedRuntimeMins": 20,
  "MaximumRuntimeMins": 45,
  "ContentFallback": "Deny",
  "SlowNetworkDeploymentMode": "DoNothing",
  "CompanyName": "Contoso",
  "M365Channel": "Current",
  "M365DeployMode": "Online",
  "M365ExcludeApps": ["Access", "Excel", "Word"],
  "M365Options": { "Languages": ["de-de", "en-us"], "OfficeMgmtCom": false, "UpdatesEnabled": false, "ForceAppShutdown": false, "RemoveMsi": false, "MigrateArch": true, "SharedComputerLicensing": true, "DeviceBasedLicensing": true },
  "SSMSInstallOptions": { "UIMode": "Passive", "DownloadThenInstall": false, "NoUpdateInstaller": false, "IncludeRecommended": true, "IncludeOptional": true, "RemoveOos": false, "ForceClose": true, "InstallPath": "D:\\SSMS", "Workloads": ["Microsoft.SqlServer.Workload.SSMS.AI", "Microsoft.SqlServer.Workload.SSMS.BI"], "Languages": ["de-de", "fr-fr"], "NoCache": true },
  "DBeaverInstallOptions": { "InstallScope": "User", "DisableAI": true },
  "AdobeReaderInstallOptions": { "Edition": "MUI", "Languages": ["de_DE", "fr_FR"], "AcceptEula": true, "DisableDesktopShortcut": true, "DisableChromeExtension": true, "DisableBrowserIntegration": true, "UpdateMode": "3" },
  "BeyondCompareKeyFile": "D:\\keys\\BC5Key.txt",
  "BeyondCompareOptions": { "DisableRemoteProfiles": true, "DisableSavedPasswords": true },
  "HiddenApplications": ["package-7zip.ps1"],
  "DiscoverySqlServer": "sql01",
  "DiscoverySqlInstance": "INST",
  "DiscoverySqlPort": 1444,
  "DiscoverySqlDatabase": "CM_ABC",
  "FirstRunCompleted": true,
  "IncludeVersionInTitle": false,
  "RemoveVersionFromTitle": true,
  "Systems": { "ConfigMgr": true, "Intune": true, "Wsus": true },
  "AppFlow": { "Tracked": ["package-7zip"], "Action": "Stage", "CadenceOverrides": {}, "ForceOnLaunch": true, "Destinations": {}, "DefaultDestinations": { "ConfigMgr": true, "WSUS": true, "Intune": false }, "OnExisting": "Overwrite" },
  "ContentDistribution": { "AutoDistribute": true, "DPGroupName": "All DPs", "DeployToTestCollection": true, "TestCollectionName": "Pilot", "CreateTestCollectionIfMissing": true },
  "Intune": { "DeploymentTarget": "MECMAndWSUS", "PublishToIntune": false, "TenantId": "contoso.onmicrosoft.com", "ClientId": "11111111-2222-3333-4444-555555555555" },
  "Wsus": { "ServerName": "wsus01.contoso.com", "PortNumber": 8531, "UseSsl": true, "Classification": "SecurityUpdates", "ApprovalGroup": "Pilot", "DeclineSuperseded": true },
  "Vulnerability": { "CacheTtlMinutes": 90, "LookupAfterCheckCm": true }
}
'@

$cwaFixture = @'
{ "Store": { "Name": "Work", "Url": "https://store.contoso.com/Citrix/Store/discovery" },
  "Installation": { "CleanInstall": false, "IncludeSSON": false, "EnableSSON": false, "AppProtection": true, "SessionPreLaunch": true, "SelfServiceMode": false, "ForceInstall": true, "BidirContentRedirection": true, "ClientHostedAppsUrl": true, "ForceLaa": true, "LegacyFtaIcons": true, "DisableSetting": "3", "ClientName": "KIOSK01" },
  "Plugins": { "MSTeamsPlugin": false, "ZoomPlugin": false, "WebExPlugin": true, "UberAgent": true, "UberAgentSkipUpgrade": true, "EPAClient": false, "SessionRecording": true, "SkipSRAgentUpgrade": true },
  "UpdateAndTelemetry": { "AutoUpdateCheck": "auto", "AutoUpdateStream": "LTSR", "AURolloutPriority": "Slow", "EnableCEIP": true, "EnableTracing": true },
  "StorePolicy": { "AllowAddStore": "A", "AllowSavePwd": "N" },
  "Components": { "Customize": true, "ReceiverInside": true, "ICA_Client": true, "AM": true, "SelfService": false, "DesktopViewer": false, "WebHelper": true, "BCR_Client": false, "USB": true, "SSON": true, "CitrixEnterpriseBrowser": true } }
'@
$tvFixture = '{ "ApiToken": "token-1", "CustomConfigId": "cfg-9", "AssignmentOptions": "--reassign", "RemoveDesktopShortcut": false }'

try {
    Set-Content -LiteralPath $global:FixturePrefsPath -Value $fullPrefs -Encoding UTF8
    Set-Content -LiteralPath (Get-CwaSwitchesPath) -Value $cwaFixture -Encoding UTF8
    Set-Content -LiteralPath (Get-TvHostConfigPath) -Value $tvFixture -Encoding UTF8
    $script:Prefs = Read-Preferences

    # --- the tree -----------------------------------------------------------------
    $definition = @(Get-OptionsSectionDefinition)
    $titles = @($definition | Where-Object { -not $_.Parent } | ForEach-Object { $_.Title })
    & $assert (($titles -join '|') -eq 'Connections|Packaging|Applications|Updates and CVEs|Security|Tools|About') 'the tree has the seven top-level nodes in order'
    $applications = @($definition | Where-Object { $_.Parent -eq 'applications' } | ForEach-Object { $_.Title })
    & $assert (($applications -join '|') -eq 'Adobe Acrobat Reader|Beyond Compare 5|Citrix Workspace App|DBeaver Community|Microsoft 365 Apps|SQL Server Management Studio|TeamViewer Host|Local Installer Sources') 'the applications are in alphabetical order and Local Installer Sources comes last'
    $citrixChildren = @($definition | Where-Object { $_.Parent -eq 'cwa' } | ForEach-Object { $_.Title })
    & $assert ($citrixChildren.Count -eq 0) 'Citrix Workspace App is one page with no child nodes'
    & $assert (@($definition | Where-Object { $_.Id -eq 'github' }).Count -eq 0) 'Security has no GitHub API node'
    & $assert (@($definition | Where-Object { $_.Id -eq 'tools' })[0].Fields.Id -contains 'gh.token') 'Tools holds the GitHub token field'

    $owner = New-OptionsTestOwner
    $dialog = New-OptionsDialog -Owner $owner
    & $assert ($dialog.Tree.Items.Count -eq 7) 'the window builds the tree'
    & $assert ($dialog.Window.Title -eq 'Options: Connections > ConfigMgr') 'the window opens on ConfigMgr and the title names the section'

    $opened = 0
    foreach ($section in $definition) {
        if ($section.Kind -eq 'Group') { continue }
        & $dialog.Navigate ([string]$section.Id) ''
        $page = $dialog.Context.Sections[[string]$section.Id]
        if (-not $page -or -not $page.Element) { throw ('FAIL: the section ' + $section.Title + ' did not build') }
        if ([string]$dialog.Window.Title -notmatch [regex]::Escape($section.Title)) { throw ('FAIL: the title does not name ' + $section.Title) }
        $opened++

        # Every input has the shared height; the layout probe below checks its width at the native size and at the smallest size.
        foreach ($tb in @(Get-OptionsInputControl -Element $page.Element)) {
            $label = $section.Title + ' / ' + [System.Windows.Automation.AutomationProperties]::GetAutomationId($tb) + ' ' + $tb.GetType().Name
            if ($tb.Height -ne 28) { throw ('FAIL: the height of ' + $label + ' is ' + $tb.Height) }
        }
        Test-OptionsLayout -Element $page.Element -Fields @($section.Fields) -Label $section.Title -NativeWidth 900 -SmallestWidth 560
    }
    & $assert ($opened -eq @($definition | Where-Object { $_.Kind -ne 'Group' }).Count -and $opened -ge 20) ('every section opens ({0} sections)' -f $opened)
    & $assert (@($global:LogLines | Where-Object { $_ -match 'could not be built' }).Count -eq 0) 'no section failed to build'

    # Every field has a tooltip on its title and its editor.
    $missingTip = @(foreach ($section in $definition) { foreach ($f in @($section.Fields)) { if ($f.Id -and $f.Kind -notin 'Group', 'Note' -and [string]::IsNullOrWhiteSpace([string]$f.Tip)) { $f.Id } } })
    & $assert ($missingTip.Count -eq 0) 'every field has a tooltip'

    # --- search -------------------------------------------------------------------
    $hits = @(Search-OptionsIndex -Index $dialog.Index -Query 'sql port')
    & $assert ($hits.Count -ge 1 -and $hits[0].FieldId -eq 'cm.sqlPort' -and $hits[0].SectionId -eq 'cm') 'the search finds SQL port in ConfigMgr'
    $hits = @(Search-OptionsIndex -Index $dialog.Index -Query 'client secret')
    & $assert ($hits[0].FieldId -eq 'int.secret') 'the search finds Client secret'
    $hits = @(Search-OptionsIndex -Index $dialog.Index -Query 'not offered')
    & $assert ($hits[0].SectionId -eq 'productfilter') 'the search finds Not offered by Discover in Product Filter'
    & $assert (@(Search-OptionsIndex -Index $dialog.Index -Query 'zzzz no such setting').Count -eq 0) 'the search returns nothing for an unknown name'

    & $dialog.Navigate 'about' ''
    $dialog.SearchBox.Text = 'use ssl'
    & $assert ($dialog.SearchList.Visibility -eq 'Visible' -and $dialog.SearchList.Items.Count -ge 1) 'typing in the search box lists the matches in place of the tree'
    & $assert ([string]$dialog.SearchList.Items[0].Tag.FieldId -eq 'wsus.ssl') 'the first match of Use SSL is the WSUS field'
    & $dialog.Jump $dialog.SearchList.Items[0].Tag
    & $assert ([string]$dialog.Tree.SelectedItem.Tag -eq 'wsus' -and $dialog.Tree.Visibility -eq 'Visible') 'the jump selects the WSUS section and shows the tree again'
    & $assert ($dialog.Window.Title -eq 'Options: Connections > WSUS') 'the jump changes the window title'
    $dialog.SearchBox.Text = ''
    & $assert ($dialog.SearchList.Visibility -eq 'Collapsed') 'an empty search box hides the match list'

    # --- the first save keeps every setting ---------------------------------------
    $before = $script:Prefs | ConvertTo-Json -Depth 8
    $cwaBefore = $dialog.Context.Aux.cwa | ConvertTo-Json -Depth 8
    $tvBefore = $dialog.Context.Aux.tv | ConvertTo-Json -Depth 8
    $saved = & $dialog.Apply
    & $assert ($saved -eq $true -and $dialog.Result.Saved) 'OK saves with no message when every value is valid'
    $after = $script:Prefs | ConvertTo-Json -Depth 8
    & $assert ($before -ceq $after) 'saving the unchanged window changes no preference'
    & $assert (($dialog.Context.Aux.cwa | ConvertTo-Json -Depth 8) -ceq $cwaBefore) 'saving the unchanged window changes no Citrix switch'
    & $assert (($dialog.Context.Aux.tv | ConvertTo-Json -Depth 8) -ceq $tvBefore) 'saving the unchanged window changes no TeamViewer Host setting'
    $reloaded = Read-Preferences
    & $assert (($reloaded | ConvertTo-Json -Depth 8) -ceq $before) 'the saved file loads back unchanged'

    # --- values that cannot be stored stop the save --------------------------------
    $script:Prefs = Read-Preferences
    $dialog2 = New-OptionsDialog -Owner (New-OptionsTestOwner) -InitialSection 'wsus'
    $portEditor = $dialog2.Context.Sections['wsus'].Editors['wsus.port']
    & $portEditor.Set '70000'
    $global:Messages = @()
    $global:SavedCount = 0
    & $assert ((& $dialog2.Apply) -eq $false) 'a port above 65535 stops the save'
    & $assert ($global:SavedCount -eq 0 -and @($global:Messages | Where-Object { $_ -match 'Port must be a whole number from 1 to 65535' }).Count -eq 1) 'the message names the field and the range'
    & $assert ([string]$dialog2.Tree.SelectedItem.Tag -eq 'wsus') 'the window stays on the field with the problem'
    & $portEditor.Set 'abc'
    & $assert ((& $dialog2.Apply) -eq $false) 'a port that is not a number stops the save'
    & $portEditor.Set '8532'
    & $assert ((& $dialog2.Apply) -eq $true -and [int]$script:Prefs.Wsus.PortNumber -eq 8532) 'a custom port in range is stored, with SSL on'
    & $assert ([bool]$script:Prefs.Wsus.UseSsl) 'the Use SSL switch keeps its value with a custom port'

    # --- edits survive a save and a reload ------------------------------------------
    $script:Prefs = Read-Preferences
    $dialog3 = New-OptionsDialog -Owner (New-OptionsTestOwner)
    $cm = $dialog3.Context.Sections['cm']
    & $cm.Editors['cm.siteCode'].Set 'XYZ'
    & $cm.Editors['cm.sqlInstance'].Set 'SQL2'
    & $cm.Editors['cm.sqlPort'].Set ''
    & $dialog3.Navigate 'dtdefaults' ''
    & $dialog3.Context.Sections['dtdefaults'].Editors['pkg.estimated'].Set '25'
    & $dialog3.Navigate 'm365' ''
    & $dialog3.Context.Sections['m365'].Editors['m365.languages'].Set 'fr-fr, de-de'
    & $dialog3.Navigate 'cwa' ''
    & $dialog3.Context.Sections['cwa'].Editors['cwa.rollout'].Set 'Fast'
    & $dialog3.Navigate 'tools' ''
    & $assert ((& $dialog3.Apply) -eq $true) 'OK saves the edits'
    $script:Prefs = Read-Preferences
    & $assert ($script:Prefs.SiteCode -eq 'XYZ' -and $script:Prefs.DiscoverySqlInstance -eq 'SQL2' -and $script:Prefs.DiscoverySqlPort -eq 0) 'the site code, the SQL instance and an empty SQL port reload'
    & $assert ($script:Prefs.EstimatedRuntimeMins -eq 25 -and (@($script:Prefs.M365Options.Languages) -join ',') -eq 'fr-fr,de-de') 'the runtime and the Office languages reload'
    & $assert ((Read-CwaSwitches).UpdateAndTelemetry.AURolloutPriority -eq 'Fast') 'the Citrix rollout priority reloads from its own file'

    # --- a SQL server name with an instance and a port shows split ---------------------
    $legacy = $fullPrefs | ConvertFrom-Json
    $legacy.DiscoverySqlServer = 'sql07\INST7,1450'
    $legacy.PSObject.Properties.Remove('DiscoverySqlInstance')
    $legacy.PSObject.Properties.Remove('DiscoverySqlPort')
    $legacy | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $global:FixturePrefsPath -Encoding UTF8
    $script:Prefs = Read-Preferences
    & $assert ($script:Prefs.DiscoverySqlServer -eq 'sql07\INST7,1450' -and $script:Prefs.DiscoverySqlPort -eq 0) 'a file with server\instance,port in the old field loads as written'
    $page = New-OptionsTestPage -Id 'cm'
    & $assert ((& $page.Editors['cm.sqlServer'].Get) -eq 'sql07' -and (& $page.Editors['cm.sqlInstance'].Get) -eq 'INST7' -and (& $page.Editors['cm.sqlPort'].Get) -eq '1450') 'the old field shows split into the SQL server, SQL instance and SQL port fields'
    & $assert ($page.Editors['cm.sqlTarget'].Text.Text -match 'sql07\\INST7,1450, database CM_ABC') 'the target line shows server\instance,port'
    & $page.Commit
    & $assert ($script:Prefs.DiscoverySqlServer -eq 'sql07' -and $script:Prefs.DiscoverySqlInstance -eq 'INST7' -and $script:Prefs.DiscoverySqlPort -eq 1450) 'the save writes the three parts'
    & $assert ((Join-CMSqlServerName -Server $script:Prefs.DiscoverySqlServer -Instance $script:Prefs.DiscoverySqlInstance -Port $script:Prefs.DiscoverySqlPort) -eq 'sql07\INST7,1450') 'the parts join to the name that Discover connects to'

    # --- an imported Microsoft 365 configuration.xml survives a save and a reload -----------
    Set-Content -LiteralPath $global:FixturePrefsPath -Value $fullPrefs -Encoding UTF8
    $script:Prefs = Read-Preferences
    $sourceXml = Join-Path $fixtureDir 'mine.xml'
    $caf = [string][char]0x00E9
    Set-Content -LiteralPath $sourceXml -Encoding UTF8 -Value @(
        '<Configuration>',
        '  <Add OfficeClientEdition="32" Channel="Current" Version="16.0.1.1" SourcePath="D:\x">',
        '    <Product ID="O365ProPlusRetail"><Language ID="en-us" /></Product>',
        '  </Add>',
        ('  <AppSettings><Setup Name="Company" Value="Caf' + $caf + '" /></AppSettings>'),
        '</Configuration>')
    $dialog4 = New-OptionsDialog -Owner (New-OptionsTestOwner) -InitialSection 'm365'
    $m365 = $dialog4.Context.Sections['m365']
    & $assert ($m365.Editors['m365.useCustom'].Focus.IsEnabled -eq $false) 'the box Use my own configuration.xml waits for an imported file'
    & $assert ((& $m365.Actions.ImportM365 $sourceXml) -eq $true) 'a well-formed configuration.xml with Add and Product imports'
    & $assert ($m365.Editors['m365.file'].Text.Text -match 'mine\.xml') 'the section shows the name of the imported file'
    & $assert ($m365.Editors['m365.file'].Text.Text -match 'bytes|KB') 'the section shows the size of the imported file'
    & $assert ($m365.Editors['m365.useCustom'].Focus.IsEnabled) 'the box Use my own configuration.xml is available after the import'
    $m365.Editors['m365.useCustom'].Focus.IsChecked = $true
    & $assert (-not $m365.Editors['m365.channel'].Element.IsEnabled -and -not $m365.Editors['m365.deploy'].Element.IsEnabled -and -not $m365.Editors['m365.exclude'].Element.IsEnabled -and -not $m365.Editors['m365.company'].Element.IsEnabled) 'the generated settings are disabled'
    & $assert ($m365.Editors['m365.reason'].Text.Visibility -eq 'Visible' -and $m365.Editors['m365.reason'].Text.Text -match 'imported file sets') 'a line says why they are disabled'
    & $assert ((Test-Path -LiteralPath (Get-M365CustomConfigPath)) -eq $false) 'nothing is stored before OK'
    & $assert ((& $dialog4.Apply) -eq $true) 'OK saves'
    & $assert ((Test-Path -LiteralPath (Get-M365CustomConfigPath)) -and $script:Prefs.M365CustomConfig.Enabled -and $script:Prefs.M365CustomConfig.FileName -eq 'mine.xml' -and $script:Prefs.M365CustomConfig.SizeBytes -gt 100) 'OK stores the file and records its name and size'
    $script:Prefs = Read-Preferences
    $dialog5 = New-OptionsDialog -Owner (New-OptionsTestOwner) -InitialSection 'm365'
    $again = $dialog5.Context.Sections['m365']
    & $assert ($again.Editors['m365.useCustom'].Focus.IsChecked -eq $true -and $again.Editors['m365.file'].Text.Text -match 'mine\.xml' -and -not $again.Editors['m365.channel'].Element.IsEnabled) 'the stored file and the box come back after a reload'
    $again.Editors['m365.file'].Buttons['remove'].RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    & $assert ($again.Editors['m365.channel'].Element.IsEnabled -and -not $again.Editors['m365.useCustom'].Focus.IsEnabled -and $again.Editors['m365.file'].Text.Text -match 'No file') 'Remove returns the generated settings and clears the box'
    & $assert ((& $dialog5.Apply) -eq $true -and -not (Test-Path -LiteralPath (Get-M365CustomConfigPath)) -and -not $script:Prefs.M365CustomConfig.Enabled -and $script:Prefs.M365CustomConfig.FileName -eq '') 'OK deletes the stored file and clears the record'
    $broken = Join-Path $fixtureDir 'broken.xml'
    Set-Content -LiteralPath $broken -Value '<Configuration><Add></Configuration>' -Encoding UTF8
    $script:Prefs = Read-Preferences
    $dialog6 = New-OptionsDialog -Owner (New-OptionsTestOwner) -InitialSection 'm365'
    $global:Messages = @()
    & $assert ((& $dialog6.Context.Sections['m365'].Actions.ImportM365 $broken) -eq $false -and @($global:Messages | Where-Object { $_ -match 'Import Failed' }).Count -eq 1) 'a file that is not well-formed XML is refused with a message'
    $noAdd = Join-Path $fixtureDir 'noadd.xml'
    Set-Content -LiteralPath $noAdd -Value '<Configuration><Display Level="None" /></Configuration>' -Encoding UTF8
    & $assert ((& $dialog6.Context.Sections['m365'].Actions.ImportM365 $noAdd) -eq $false) 'a configuration without an Add element is refused'
    $wrongRoot = Join-Path $fixtureDir 'wrongroot.xml'
    Set-Content -LiteralPath $wrongRoot -Value '<Settings><Add><Product ID="x" /></Add></Settings>' -Encoding UTF8
    & $assert ((& $dialog6.Context.Sections['m365'].Actions.ImportM365 $wrongRoot) -eq $false) 'a file without a Configuration root is refused'

    # --- the Setup wizard and the Connections sections share the fields ------------------------
    $table = Get-OptionsFieldTable
    $wizard = $script:LauncherAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-FirstRunWizard' }, $false)
    $wizardIds = @([regex]::Matches($wizard.Extent.Text, '(?<![\w$.])(?:sys|cm|int|wsus|pkg)\.[A-Za-z]+') | ForEach-Object { $_.Value } | Sort-Object -Unique)
    & $assert ((@('cm.provider', 'cm.siteCode', 'int.client', 'int.secret', 'int.tenant', 'pkg.downloadRoot', 'wsus.port', 'wsus.server', 'wsus.ssl') | Where-Object { $wizardIds -notcontains $_ }).Count -eq 0) 'the wizard asks for the Connections fields by their ids'
    & $assert (@($wizardIds | Where-Object { -not $table.ContainsKey($_) }).Count -eq 0) 'every field id of the wizard is a field of the Options definitions'
    & $assert ($wizard.Extent.Text -notmatch 'txtWiz|chkWiz(Mecm|Intune|Wsus)\b') 'the wizard keeps no copy of a Connections field'
    Set-Content -LiteralPath $global:FixturePrefsPath -Value $fullPrefs -Encoding UTF8
    $script:Prefs = Read-Preferences
    $group = New-OptionsFieldGroup -Context (New-OptionsContext -Prefs $script:Prefs) -Field @($table['cm.siteCode'], $table['wsus.port'], $table['wsus.ssl'])
    & $assert ((& $group.Editors['cm.siteCode'].Get) -eq 'ABC' -and (& $group.Editors['wsus.port'].Get) -eq '8531') 'a group of the wizard fields shows the stored values'
    # The wizard column is 700 wide at the native size and 570 at the smallest size: no field row is cut off.
    $wizardGroup = New-OptionsFieldGroup -Context (New-OptionsContext -Prefs $script:Prefs) -Field @($table['cm.siteCode'], $table['cm.provider'], $table['cm.fileShare'], $table['pkg.downloadRoot'], $table['int.tenant'], $table['int.client'], $table['int.secret'], $table['wsus.server'], $table['wsus.port'], $table['wsus.ssl'])
    Test-OptionsLayout -Element $wizardGroup.Element -Fields @($table['cm.siteCode'], $table['cm.provider'], $table['cm.fileShare'], $table['pkg.downloadRoot'], $table['int.tenant'], $table['int.client'], $table['int.secret'], $table['wsus.server'], $table['wsus.port']) -Label 'Setup wizard' -NativeWidth 700 -SmallestWidth 570
    & $assert $true 'no field row of the Setup wizard is cut off at the native size or at the smallest size'
    Set-Content -LiteralPath $global:FixturePrefsPath -Value '{}' -Encoding UTF8
    $script:Prefs = Read-Preferences
    $fresh = New-OptionsFieldGroup -Context (New-OptionsContext -Prefs $script:Prefs) -Field @($table['wsus.port'], $table['wsus.ssl'])
    $fresh.Editors['wsus.ssl'].Focus.IsChecked = $true
    & $assert ((& $fresh.Editors['wsus.port'].Get) -eq '8531') 'in any group, Use SSL moves a default port to 8531'

    & $assert ($script:LayoutChecked -ge 80) ('the layout probe measured {0} inputs' -f $script:LayoutChecked)
    Write-Host ''
    Write-Host ('PASS: Options tree probe, {0} checks' -f $script:PassedChecks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
