#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The grid of all applications in Options (Applications, All Applications): the registry of settings, what each cell offers, what each
    bulk action does, what the grid writes, and that the packaging and One Click code read and use the value that the grid wrote.

.DESCRIPTION
    Windows PowerShell 5.1 and PowerShell 7. The cases open no window and read no installed product. They build the working copy and the
    rows from the catalog headers, run the pure functions of the grid, write the preferences the way Save-Preferences does, read them back
    with Read-Preferences, and follow each value into the code that uses it.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerWorkbench.psd1') -Force -DisableNameChecking -ErrorAction SilentlyContinue
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerOneClick.psd1') -Force -DisableNameChecking

    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $script:GridFunction = @()
    $script:LauncherText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-apppackager.ps1') -Raw
    foreach ($name in 'Get-ApplicationGridLabel', 'Initialize-ApplicationGridType', 'Get-ApplicationGridCapability', 'Get-ApplicationGridItemList', 'ConvertTo-ApplicationGridMap',
        'New-ApplicationGridWork', 'Get-ApplicationGridEntry', 'Get-ApplicationGridOwn', 'Get-ApplicationGridEffectiveSupersedence', 'Get-ApplicationGridSetting',
        'Get-ApplicationGridCell', 'Set-ApplicationGridCell', 'Update-ApplicationGridRow', 'New-ApplicationGridRow', 'Set-ApplicationGridBulkValue',
        'ConvertTo-ApplicationGridObject', 'Save-ApplicationGridWork', 'Get-ApplicationGridNotice',
        'ConvertTo-ApplicationGridTree', 'Get-ApplicationGridProfileContext', 'Set-ApplicationGridProfileValue', 'Save-ApplicationGridProfile',
        'Get-WorkbenchOverrideValue', 'Test-WorkbenchOverridePresent', 'Set-WorkbenchOverrideValue', 'Clear-WorkbenchOverrideValue',
        'Resolve-FirstRunCompleted', 'Read-Preferences', 'Get-WsusClassificationNameList',
        'ConvertTo-InstallModeValue', 'ConvertTo-VariantsJson', 'ConvertTo-RequirementsJson', 'ConvertTo-CommandsJson',
        'Get-InstallModesMapForContext', 'Get-VariantsMapForContext', 'Get-RequirementsMapForContext', 'Get-CommandsMapForContext', 'Get-TitleModesMapForContext',
        'Set-PackagerEnvironment', 'Get-SupersedenceLockedTitleMode') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "start-apppackager.ps1 no longer defines $name." }
        # The functions are global, because the script blocks of the registry are closures and look their functions up from the global scope,
        # the way they do in the app. AfterAll removes them again.
        $body = $fn.Body.Extent.Text
        Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($body.Substring(1, $body.Length - 2)))
        $script:GridFunction += $name
    }
    $script:PrefPath = Join-Path $TestDrive 'preferences.json'
    function Get-PreferencesPath { $script:PrefPath }

    $script:Label = Get-ApplicationGridLabel
    $script:Setting = @(Get-ApplicationGridSetting)
    function Get-GridSetting { param([string]$Id) @($script:Setting | Where-Object { $_.Id -eq $Id })[0] }

    # A catalog row from the header lines of one packager file, the way Get-Packagers reads them.
    function Get-CatalogRow {
        foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1' -File | Sort-Object Name)) {
            $head = @(Get-Content -LiteralPath $file.FullName -TotalCount 60)
            $variants = @(); $modes = @(); $app = $file.BaseName; $vendor = 'Vendor'; $wsus = ''; $reason = ''
            foreach ($line in $head) {
                if ($line -match '^\s*(?:#\s*)?SupportsVariants\s*:\s*(.+?)\s*$') { $variants = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() }) }
                if ($line -match '^\s*(?:#\s*)?SupportsInstallModes\s*:\s*(.+?)\s*$') { $modes = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() }) }
                if ($line -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$') { $app = $Matches[1] }
                if ($line -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$') { $vendor = $Matches[1] }
                if ($line -match '^\s*(?:#\s*)?WsusSupport\s*:\s*(Yes|No)\s*(?:\(\s*([A-Za-z|]+)\s*\))?\s*$') { $wsus = $Matches[1]; $reason = $Matches[2] }
            }
            [pscustomobject]@{ Script = $file.Name; Application = $app; Vendor = $vendor; SupportsVariants = $variants; SupportsInstallModes = $modes; WsusSupport = $wsus; WsusSupportReason = $reason }
        }
    }
    function New-TestPackager {
        param([string]$Base = 'package-demo', [string[]]$Variants = @(), [string[]]$Modes = @(), [string]$Application = 'Demo App', [string]$Wsus = 'Yes', [string]$WsusReason = '')
        [pscustomobject]@{ Script = ($Base + '.ps1'); Application = $Application; Vendor = 'Demo Vendor'; SupportsVariants = $Variants; SupportsInstallModes = $Modes; WsusSupport = $Wsus; WsusSupportReason = $WsusReason }
    }
    function New-TestPrefs {
        param([string]$Json = '{}')
        Set-Content -LiteralPath $script:PrefPath -Value $Json -Encoding UTF8
        return (Read-Preferences)
    }
    function Save-TestPrefs {
        param($Prefs)
        $Prefs | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:PrefPath -Encoding UTF8
        return (Read-Preferences)
    }
    # Sets columns of one application, saves the preferences and reads them back, as the Options window and the next start do.
    function Invoke-Grid {
        param([System.Collections.IDictionary]$Set, [string]$Json = '{}', $Packager = (New-TestPackager -Variants @('Architecture', 'Network') -Modes @('CurrentUser', 'AllUsers')))
        $prefs = New-TestPrefs $Json
        $work = New-ApplicationGridWork -Prefs $prefs
        $initial = New-ApplicationGridWork -Prefs $prefs
        foreach ($id in $Set.Keys) { [void](Set-ApplicationGridCell -Setting (Get-GridSetting $id) -Work $work -Packager $Packager -Value $Set[$id]) }
        Save-ApplicationGridWork -Prefs $prefs -Work $work -Initial $initial
        $script:Prefs = Save-TestPrefs -Prefs $prefs
        return $script:Prefs
    }
    function Get-SavedJson { Get-Content -LiteralPath $script:PrefPath -Raw }
}

AfterAll { foreach ($name in $script:GridFunction) { Remove-Item -Path ('function:global:' + $name) -ErrorAction SilentlyContinue } }

Describe 'The registry' {
    It 'has one entry for each setting, with the columns in the order of the design' {
        (@($script:Setting | ForEach-Object { $_.Id }) -join ',') | Should -Be 'source,scope,supersedence,title,rule,network,netsplit,install,uninstall,oneclick,cadence,toconfigmgr,tointune,towsus,estimated,maximum,reboot'
        (@($script:Setting | Where-Object { $_.Visible } | ForEach-Object { $_.Id }) -join ',') | Should -Be 'source,scope,supersedence,title,rule'
    }
    It 'gives every entry a header, a tip of one sentence, and the functions that the grid calls' {
        foreach ($s in $script:Setting) {
            $s.Header | Should -Match '^[A-Z]'
            $s.Tip.Length | Should -BeLessThan 101 -Because $s.Id
            $s.Tip | Should -Match '\.$'
            $s.Kind | Should -BeIn 'Choice', 'Text'
            foreach ($f in 'Get', 'Set', 'NotApplicable') { $s[$f] | Should -BeOfType [scriptblock] -Because ($s.Id + ' ' + $f) }
            if ($s.Kind -eq 'Choice') { @($s.Choices).Count | Should -BeGreaterThan 1 }
        }
    }
    It 'resets a column with its first choice' {
        (Get-GridSetting 'source').Choices[0] | Should -Be 'Packager default'
        (Get-GridSetting 'supersedence').Choices[0] | Should -Be 'Follow the default'
        (Get-GridSetting 'oneclick').Choices[0] | Should -Be 'No'
        (Get-GridSetting 'network').Choices[0] | Should -Be 'Any'
    }
}

Describe 'What each cell offers' {
    It 'opens Package Source only for a packager that declares an architecture split' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $s = Get-GridSetting 'source'
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Variants @('Architecture'))).Reason | Should -BeNullOrEmpty
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Variants @('Language'))).Reason | Should -Match 'one package source'
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager)).Reason | Should -Not -BeNullOrEmpty
    }
    It 'shows the own architecture of a sibling packager, closed, and says why' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        foreach ($case in @(@('package-temurin-jdk8-x64', 'x64'), @('package-temurin-jdk8-x86', 'x86'), @('package-citrixworkspace-ltsr-arm64', 'ARM64'))) {
            $cell = Get-ApplicationGridCell -Setting (Get-GridSetting 'source') -Work $work -Packager (New-TestPackager -Base $case[0])
            $cell.Value | Should -Be $case[1]
            $cell.Reason | Should -Match 'another packager'
        }
    }
    It 'opens Install Scope only for a packager that declares both modes' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $s = Get-GridSetting 'scope'
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Modes @('CurrentUser', 'AllUsers'))).Reason | Should -BeNullOrEmpty
        foreach ($modes in @('AllUsers'), @('CurrentUser'), @()) { (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Modes $modes)).Reason | Should -Match 'one install scope' }
    }
    It 'opens Network Split only for a packager that declares a network split' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $s = Get-GridSetting 'netsplit'
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Variants @('Network'))).Reason | Should -BeNullOrEmpty
        (Get-ApplicationGridCell -Setting $s -Work $work -Packager (New-TestPackager -Variants @('Architecture'))).Reason | Should -Match 'network split'
    }
    It 'closes Publish to WSUS with the reason of the packager, and leaves the other destinations open' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $p = New-TestPackager -Wsus 'No' -WsusReason 'PerUserInstall|CustomInstall'
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'towsus') -Work $work -Packager $p).Reason | Should -Be 'WSUS cannot carry Demo App. Reason: PerUserInstall, CustomInstall.'
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'tointune') -Work $work -Packager $p).Reason | Should -BeNullOrEmpty
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'towsus') -Work $work -Packager (New-TestPackager)).Reason | Should -BeNullOrEmpty
    }
    It 'follows the header tags of the real catalog' {
        $catalog = @(Get-CatalogRow)
        $catalog.Count | Should -BeGreaterThan 300
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $rows = @($catalog | ForEach-Object { New-ApplicationGridRow -Packager $_ -Setting $script:Setting -Work $work })
        $rows.Count | Should -Be $catalog.Count
        @($rows | Where-Object { $_['e_source'] }).Count | Should -Be @($catalog | Where-Object { $_.SupportsVariants -contains 'Architecture' }).Count
        @($rows | Where-Object { $_['e_scope'] }).Count | Should -Be @($catalog | Where-Object { $_.SupportsInstallModes -contains 'AllUsers' -and $_.SupportsInstallModes -contains 'CurrentUser' }).Count
        @($rows | Where-Object { $_['e_netsplit'] }).Count | Should -Be @($catalog | Where-Object { $_.SupportsVariants -contains 'Network' }).Count
        @($rows | Where-Object { -not $_['e_towsus'] }).Count | Should -Be @($catalog | Where-Object { $_.WsusSupport -eq 'No' }).Count
        @($rows | Where-Object { -not $_['e_towsus'] }).Count | Should -BeGreaterThan 0
        @($rows | Where-Object { -not $_['e_source'] })[0]['t_source'] | Should -Not -BeNullOrEmpty
    }
    It 'gives each row the cell values, the flags, the tips and the lists' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-chrome' -Application 'Google Chrome' -Variants @('Architecture')) -Setting $script:Setting -Work $work
        $row.Id | Should -Be 'catalog:package-chrome'
        $row['v_source'] | Should -Be 'Packager default'
        $row['e_source'] | Should -BeTrue
        $row['t_source'] | Should -Match '^Google Chrome: '
        $row['i_source'] | Should -Be @('Packager default', 'x64 and ARM64')
        $row['v_supersedence'] | Should -Be 'Follow the default'
        $row['v_oneclick'] | Should -Be 'No'
        $other = New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b') -Setting $script:Setting -Work $work
        [object]::ReferenceEquals($row['i_title'], $other['i_title']) | Should -BeTrue
    }
    It 'ignores a null that a combo box writes while it changes its items, and raises an event for a value that changes' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture')) -Setting $script:Setting -Work $work
        $script:Events = @()
        $row.add_CellChanged([Action[AppPackagerGridRow, string]]{ param($r, $id) $script:Events += $id })
        # A binding writes a real null; PowerShell turns $null into an empty string, so the indexer is called by reflection.
        $row.GetType().GetProperty('Item').SetValue($row, $null, @('v_source'))
        $row['v_source'] | Should -Be 'Packager default'
        $row['v_source'] = 'x64 and ARM64'
        $row['v_source'] = 'x64 and ARM64'
        $row.Load('v_scope', 'User')
        @($script:Events) | Should -Be @('source')
    }
}

Describe 'The title lock' {
    It 'closes the title and shows Include version while supersedence is on for the application' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $p = New-TestPackager
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p -Value 'No version')
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $p -Value 'Replace')
        $cell = Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p
        $cell.Shown | Should -Be 'Include version'
        $cell.Value | Should -Be 'No version' -Because 'the lock never writes the stored choice'
        $cell.Reason | Should -Match 'separate applications'
        Set-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p -Value 'Follow the default' | Should -BeFalse
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p).Value | Should -Be 'No version'
    }
    It 'shows the stored choice again when supersedence is off' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $p = New-TestPackager
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p -Value 'No version')
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $p -Value 'Replace and uninstall')
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $p -Value 'Off')
        $cell = Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p
        $cell.Reason | Should -BeNullOrEmpty
        $cell.Shown | Should -Be 'No version'
    }
    It 'follows the default for an application that follows the default' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs '{ "Supersedence": { "Mode": "Replace" } }')
        $p = New-TestPackager
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p).Reason | Should -Not -BeNullOrEmpty
        $work.DefaultSupersedence = 'Off'
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p).Reason | Should -BeNullOrEmpty
        $work.DefaultSupersedence = 'Replace'
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $p -Value 'Off')
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'title') -Work $work -Packager $p).Reason | Should -BeNullOrEmpty -Because 'an own value of Off outranks a default that is on'
    }
    It 'tells the user how many applications are locked' {
        $work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $rows = @(1..3 | ForEach-Object { New-ApplicationGridRow -Packager (New-TestPackager -Base "package-$_") -Setting $script:Setting -Work $work })
        Get-ApplicationGridNotice -Row $rows -Work $work | Should -BeNullOrEmpty
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $rows[0].Packager -Value 'Replace')
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $work -Packager $rows[1].Packager -Value 'Replace and uninstall')
        Get-ApplicationGridNotice -Row $rows -Work $work | Should -Match '^Supersedence is on for 2 application\(s\)\.'
    }
}

Describe 'The bulk actions' {
    BeforeAll {
        function Get-CellValue { param($Id, $Row) (Get-ApplicationGridCell -Setting (Get-GridSetting $Id) -Work $script:Work -Packager $Row.Packager).Value }
    }
    BeforeEach {
        $script:Work = New-ApplicationGridWork -Prefs (New-TestPrefs)
        $script:Rows = @(
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a' -Application 'Alpha Tool' -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')) -Setting $script:Setting -Work $script:Work)
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b' -Application 'Beta Tool' -Modes @('CurrentUser', 'AllUsers')) -Setting $script:Setting -Work $script:Work)
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-c' -Application 'Gamma Editor') -Setting $script:Setting -Work $script:Work)
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-d-x64' -Application 'Delta Editor' -Wsus 'No' -WsusReason 'PerUserInstall') -Setting $script:Setting -Work $script:Work)
        )
    }

    It 'changes only the checked applications for the scope Checked' {
        $script:Rows[0].Checked = $true; $script:Rows[2].Checked = $true
        $r = Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'supersedence') -Scope Checked -Value 'Replace' -Row $script:Rows -Work $script:Work
        $r.Changed | Should -Be 2
        @($script:Rows | ForEach-Object { Get-CellValue 'supersedence' $_ }) | Should -Be @('Replace', 'Follow the default', 'Replace', 'Follow the default')
    }
    It 'changes the applications that pass the search for the scope Filtered' {
        $filter = { param($row) $row.Application -like '*Editor*' }
        $r = Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'rule') -Scope Filtered -Value 'Exact' -Row $script:Rows -Work $script:Work -Filter $filter
        $r.Matched | Should -Be 2
        @($script:Rows | ForEach-Object { Get-CellValue 'rule' $_ }) | Should -Be @('Follow the default', 'Follow the default', 'Exact', 'Exact')
    }
    It 'changes every application for the scope All, and resets the column for the scope None' {
        (Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'rule') -Scope All -Value 'Minimum' -Row $script:Rows -Work $script:Work).Changed | Should -Be 4
        @($script:Rows | ForEach-Object { Get-CellValue 'rule' $_ } | Select-Object -Unique) | Should -Be @('Minimum')
        (Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'rule') -Scope None -Row $script:Rows -Work $script:Work).Changed | Should -Be 4
        @($script:Rows | ForEach-Object { Get-CellValue 'rule' $_ } | Select-Object -Unique) | Should -Be @('Follow the default')
    }
    It 'skips a cell that does not apply, counts it, and never writes it' {
        $r = Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'scope') -Scope All -Value 'User' -Row $script:Rows -Work $script:Work
        @($script:Rows | ForEach-Object { Get-CellValue 'scope' $_ }) | Should -Be @('User', 'User', 'Packager default', 'Packager default')
        $r.Changed | Should -Be 2
        $r.Skipped | Should -Be 2
        [string]$script:Work.Apps['package-c']['InstallMode'] | Should -Be '' -Because 'a skipped cell writes no value'
        $r = Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'towsus') -Scope All -Value 'No' -Row $script:Rows -Work $script:Work
        $r.Skipped | Should -Be 1
        $script:Work.Destinations.Contains('package-d-x64') | Should -BeFalse
    }
    It 'sets a title for an application that is not locked and skips one that is' {
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'supersedence') -Work $script:Work -Packager $script:Rows[0].Packager -Value 'Replace')
        $r = Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'title') -Scope All -Value 'No version' -Row $script:Rows -Work $script:Work
        $r.Changed | Should -Be 3
        $r.Skipped | Should -Be 1
    }
    It 'writes a typed value for a text column, and refuses a value that is not a choice' {
        (Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'cadence') -Scope All -Value '14' -Row $script:Rows -Work $script:Work).Changed | Should -Be 4
        (Get-CellValue 'cadence' $script:Rows[0]) | Should -Be '14'
        (Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'cadence') -Scope None -Row $script:Rows -Work $script:Work).Changed | Should -Be 4
        (Get-CellValue 'cadence' $script:Rows[0]) | Should -Be ''
        { Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'scope') -Scope All -Value 'Replace' -Row $script:Rows -Work $script:Work } | Should -Throw
    }
    It 'accepts an empty list of rows' {
        (Set-ApplicationGridBulkValue -Setting (Get-GridSetting 'title') -Scope All -Value 'No version' -Row @() -Work $script:Work).Changed | Should -Be 0
    }
}

Describe 'What the grid writes' {
    It 'stores each column under the key that the code reads' {
        $first = Invoke-Grid -Set ([ordered]@{ title = 'No version' })
        $set = [ordered]@{ source = 'x64 and ARM64'; scope = 'User'; supersedence = 'Replace and uninstall'; rule = 'Exact'; network = 'VPN only'
            install = 'setup.exe /silent'; uninstall = 'setup.exe /remove'; oneclick = 'Yes'; cadence = '30'; toconfigmgr = 'No'; tointune = 'Yes' }
        $prefs = Invoke-Grid -Json (Get-SavedJson) -Set $set
        $e = $prefs.DeploymentConditions.Apps.'package-demo'
        $e.Split | Should -Be 'Architecture'
        $e.InstallMode | Should -Be 'CurrentUser'
        $e.Network | Should -Be 'VpnOnly'
        $e.TitleMode | Should -Be 'NoVersion'
        $prefs.Supersedence.TitleOverrides.'catalog:package-demo'.Mode | Should -Be 'Replace and uninstall'
        $prefs.Supersedence.TitleOverrides.'catalog:package-demo'.VersionRule | Should -Be 'Exact'
        $prefs.CommandOverrides.Apps.'package-demo'.Install | Should -Be 'setup.exe /silent'
        $prefs.CommandOverrides.Apps.'package-demo'.Uninstall | Should -Be 'setup.exe /remove'
        @($prefs.AppFlow.Tracked) | Should -Contain 'package-demo'
        [int]$prefs.AppFlow.CadenceOverrides.'package-demo' | Should -Be 30
        $prefs.AppFlow.Destinations.'package-demo'.ConfigMgr | Should -BeFalse
        $prefs.AppFlow.Destinations.'package-demo'.Intune | Should -BeTrue
    }
    It 'stores nothing for a value that resets, and removes what it wrote before' {
        [void](Invoke-Grid -Set ([ordered]@{ scope = 'User'; supersedence = 'Replace'; install = 'a'; oneclick = 'Yes'; cadence = '9'; tointune = 'Yes' }))
        $reset = [ordered]@{ scope = 'Packager default'; supersedence = 'Follow the default'; install = ''; oneclick = 'No'; cadence = ''; tointune = 'Follow the default' }
        $prefs = Invoke-Grid -Json (Get-SavedJson) -Set $reset
        $prefs.DeploymentConditions.Apps.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty
        $prefs.Supersedence.TitleOverrides.PSObject.Properties['catalog:package-demo'] | Should -BeNullOrEmpty
        $prefs.CommandOverrides.Apps.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty
        @($prefs.AppFlow.Tracked) | Should -Not -Contain 'package-demo'
        $prefs.AppFlow.CadenceOverrides.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty
        $prefs.AppFlow.Destinations.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty
    }
    It 'ignores an interval that is not a whole number from 1 to 365' {
        foreach ($bad in 'abc', '0', '366', '-3') { (Invoke-Grid -Set ([ordered]@{ cadence = $bad })).AppFlow.CadenceOverrides.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty }
    }
    It 'keeps what the grid does not own' {
        $json = '{ "DeploymentConditions": { "Apps": { "package-demo": { "Architecture": "ARM64", "Languages": ["de-DE"], "Network": "VpnOnly", "Split": "Language", "TitleMode": "", "InstallMode": "" },' +
        ' "package-gone": { "Architecture": "x64" } } }, "Supersedence": { "TitleOverrides": { "catalog:package-demo": { "Mode": "Replace" }, "custom:other": { "Mode": "Off" } } } }'
        $prefs = Invoke-Grid -Json $json -Set ([ordered]@{ rule = 'Exact' }) -Packager (New-TestPackager -Variants @('Language', 'Architecture'))
        $e = $prefs.DeploymentConditions.Apps.'package-demo'
        $e.Architecture | Should -Be 'ARM64'
        $e.Network | Should -Be 'VpnOnly'
        @($e.Languages) | Should -Be @('de-DE')
        $e.Split | Should -Be 'Language' -Because 'a Language split is not a package source'
        $prefs.DeploymentConditions.Apps.'package-gone'.Architecture | Should -Be 'x64'
        $prefs.Supersedence.TitleOverrides.'catalog:package-demo'.Mode | Should -Be 'Replace'
        $prefs.Supersedence.TitleOverrides.'catalog:package-demo'.VersionRule | Should -Be 'Exact'
        $prefs.Supersedence.TitleOverrides.'custom:other'.Mode | Should -Be 'Off'
    }
    It 'keeps one split at a time: a package source and a network split replace each other' {
        $prefs = Invoke-Grid -Set ([ordered]@{ source = 'x64 and ARM64' })
        $prefs.DeploymentConditions.Apps.'package-demo'.Split | Should -Be 'Architecture'
        $prefs = Invoke-Grid -Json (Get-SavedJson) -Set ([ordered]@{ netsplit = 'Network split' })
        $prefs.DeploymentConditions.Apps.'package-demo'.Split | Should -Be 'Network'
        $prefs = Invoke-Grid -Json (Get-SavedJson) -Set ([ordered]@{ source = 'Packager default' })
        $prefs.DeploymentConditions.Apps.'package-demo'.Split | Should -Be 'Network' -Because 'a source of Packager default must not remove a network split'
    }
    It 'writes the One Click keys as changes on top of what the preferences hold now' {
        $prefs = New-TestPrefs '{ "AppFlow": { "Tracked": ["package-a"] } }'
        $work = New-ApplicationGridWork -Prefs $prefs
        $initial = New-ApplicationGridWork -Prefs $prefs
        [void](Set-ApplicationGridCell -Setting (Get-GridSetting 'oneclick') -Work $work -Packager (New-TestPackager -Base 'package-b') -Value 'Yes')
        $prefs.AppFlow.Tracked = @('package-a', 'package-c')
        Save-ApplicationGridWork -Prefs $prefs -Work $work -Initial $initial
        @($prefs.AppFlow.Tracked) | Should -Be @('package-a', 'package-c', 'package-b')
    }
    It 'reads the keys that the earlier grid, the One Click page and the command line wrote' {
        $json = '{ "DeploymentConditions": { "Apps": { "package-demo": { "Split": "Architecture", "InstallMode": "CurrentUser", "TitleMode": "IncludeVersion", "Network": "OnSiteOnly" } } },' +
        ' "Supersedence": { "TitleOverrides": { "catalog:package-demo": { "VersionRule": "Exact", "Mode": "Off" } } }, "CommandOverrides": { "Apps": { "package-demo": { "Install": "i", "Uninstall": "u" } } },' +
        ' "AppFlow": { "Tracked": ["package-demo"], "CadenceOverrides": { "package-demo": 21 }, "DefaultDestinations": { "ConfigMgr": true, "WSUS": false, "Intune": false }, "Destinations": { "package-demo": { "ConfigMgr": true, "WSUS": false, "Intune": true } } } }'
        $prefs = New-TestPrefs $json
        $work = New-ApplicationGridWork -Prefs $prefs
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')) -Setting $script:Setting -Work $work
        $row['v_source'] | Should -Be 'x64 and ARM64'
        $row['v_scope'] | Should -Be 'User'
        $row['v_title'] | Should -Be 'Include version'
        $row['v_supersedence'] | Should -Be 'Off'
        $row['v_rule'] | Should -Be 'Exact'
        $row['v_network'] | Should -Be 'On site only'
        $row['v_install'] | Should -Be 'i'
        $row['v_uninstall'] | Should -Be 'u'
        $row['v_oneclick'] | Should -Be 'Yes'
        $row['v_cadence'] | Should -Be '21'
        $row['v_toconfigmgr'] | Should -Be 'Follow the default'
        $row['v_tointune'] | Should -Be 'Yes'
        $row['v_towsus'] | Should -Be 'Follow the default'
    }
}

Describe 'The code that packages and runs One Click uses the value that the grid wrote' {
    BeforeAll {
        $script:SavedEnv = @{}
        foreach ($n in 'APP_PACKAGER_INSTALL_MODE', 'APP_PACKAGER_VARIANTS', 'APP_PACKAGER_COMMANDS', 'APP_PACKAGER_REQUIREMENTS') { $script:SavedEnv[$n] = [Environment]::GetEnvironmentVariable($n) }
    }
    BeforeEach { foreach ($n in $script:SavedEnv.Keys) { [Environment]::SetEnvironmentVariable($n, $null) } }
    AfterAll { foreach ($n in $script:SavedEnv.Keys) { [Environment]::SetEnvironmentVariable($n, $script:SavedEnv[$n]) } }

    It 'passes Install Scope User to the child process as APP_PACKAGER_INSTALL_MODE=CurrentUser' {
        [void](Invoke-Grid -Set ([ordered]@{ scope = 'User' }))
        $map = Get-InstallModesMapForContext
        $map['package-demo'] | Should -Be 'CurrentUser'
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -InstallMode ([string]$map['package-demo'])
        $start.EnvironmentVariables['APP_PACKAGER_INSTALL_MODE'] | Should -Be 'CurrentUser'
        $env:APP_PACKAGER_INSTALL_MODE = $start.EnvironmentVariables['APP_PACKAGER_INSTALL_MODE']
        Get-RequestedInstallMode | Should -Be 'CurrentUser'
    }
    It 'passes Install Scope System as AllUsers, and the default as nothing' {
        [void](Invoke-Grid -Set ([ordered]@{ scope = 'System' }))
        (Get-InstallModesMapForContext)['package-demo'] | Should -Be 'AllUsers'
        [void](Invoke-Grid -Set ([ordered]@{ scope = 'Packager default' }) -Json (Get-SavedJson))
        (Get-InstallModesMapForContext).ContainsKey('package-demo') | Should -BeFalse
    }
    It 'passes the Package Source x64 and ARM64 as an Architecture split, and Network Split as a Network split' {
        [void](Invoke-Grid -Set ([ordered]@{ source = 'x64 and ARM64' }))
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -VariantsJson ((Get-VariantsMapForContext)['package-demo'])
        $env:APP_PACKAGER_VARIANTS = $start.EnvironmentVariables['APP_PACKAGER_VARIANTS']
        (Get-RequestedPackagerVariants).Split | Should -Be 'Architecture'
        [void](Invoke-Grid -Set ([ordered]@{ netsplit = 'Network split' }))
        $env:APP_PACKAGER_VARIANTS = (Get-VariantsMapForContext)['package-demo']
        (Get-RequestedPackagerVariants).Split | Should -Be 'Network'
    }
    It 'passes Network as a requirement rule of the deployment type' {
        [void](Invoke-Grid -Set ([ordered]@{ network = 'VPN only' }))
        $rules = (Get-RequirementsMapForContext)['package-demo'] | ConvertFrom-Json
        $rules.Rules[0].ConditionId | Should -Be 'vpn-connected'
        $rules.Rules[0].Value | Should -BeTrue
        [void](Invoke-Grid -Set ([ordered]@{ network = 'On site only' }))
        ((Get-RequirementsMapForContext)['package-demo'] | ConvertFrom-Json).Rules[0].Value | Should -BeFalse
        [void](Invoke-Grid -Set ([ordered]@{ network = 'Any' }))
        (Get-RequirementsMapForContext).ContainsKey('package-demo') | Should -BeFalse
    }
    It 'passes the install and uninstall commands as APP_PACKAGER_COMMANDS' {
        [void](Invoke-Grid -Set ([ordered]@{ install = 'setup.exe /quiet'; uninstall = 'setup.exe /remove' }))
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -CommandsJson ((Get-CommandsMapForContext)['package-demo'])
        $env:APP_PACKAGER_COMMANDS = $start.EnvironmentVariables['APP_PACKAGER_COMMANDS']
        $commands = Get-RequestedCommandOverrides
        $commands.Install | Should -Be 'setup.exe /quiet'
        $commands.Uninstall | Should -Be 'setup.exe /remove'
    }
    It 'resolves the application title from Version in Title' {
        [void](Invoke-Grid -Set ([ordered]@{ title = 'No version' }))
        $mode = (Get-TitleModesMapForContext)['package-demo']
        $mode | Should -Be 'NoVersion'
        Get-PackagedApplicationName -AppName 'Demo App 2.4.1' -Version '2.4.1' -Mode $mode | Should -Be 'Demo App'
        [void](Invoke-Grid -Set ([ordered]@{ title = 'Include version' }))
        $mode = (Get-TitleModesMapForContext)['package-demo']
        Get-PackagedApplicationName -AppName 'Demo App' -Version '2.4.1' -Mode $mode | Should -Be 'Demo App - 2.4.1'
    }
    It 'locks the title to the version while supersedence is on, whatever the stored choice is' {
        Mock Get-SupersedenceSettings { $script:Prefs.Supersedence }
        Mock Get-SupersedenceCommandLine { $null }
        [void](Invoke-Grid -Set ([ordered]@{ title = 'No version' }))
        [void](Invoke-Grid -Set ([ordered]@{ supersedence = 'Replace' }) -Json (Get-SavedJson))
        $stored = (Get-TitleModesMapForContext)['package-demo']
        $stored | Should -Be 'NoVersion' -Because 'the choice of the operator stays stored'
        Get-SupersedenceLockedTitleMode -PackagerPath 'C:\Packagers\package-demo.ps1' -TitleMode $stored | Should -Be 'IncludeVersion'
        [void](Invoke-Grid -Set ([ordered]@{ supersedence = 'Off' }) -Json (Get-SavedJson))
        Get-SupersedenceLockedTitleMode -PackagerPath 'C:\Packagers\package-demo.ps1' -TitleMode 'NoVersion' | Should -Be 'NoVersion'
    }
    It 'resolves the supersedence of the application, and the default for the others' {
        [void](Invoke-Grid -Set ([ordered]@{ supersedence = 'Replace and uninstall'; rule = 'Exact' }))
        $resolved = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        $resolved.Mode.Value | Should -Be 'ReplaceAndUninstall'
        $resolved.Mode.Source | Should -Be 'TitleOverride'
        $other = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-other'
        $other.Mode.Value | Should -Be 'Off'
        $other.ConfigMgrRule.Value | Should -Be 'Minimum'
    }
    It 'writes the operator of the detection from Version Rule' {
        [void](Invoke-Grid -Set ([ordered]@{ rule = 'Exact' }))
        $resolved = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        $resolved.ConfigMgrRule.Value | Should -Be 'Exact'
        $resolved.IntuneRule.Value | Should -Be 'Exact'
        $clause = [pscustomobject]@{ Type = 'File'; FilePath = '%ProgramFiles%\Demo'; FileName = 'demo.exe'; ExpectedValue = '2.4.1'; Operator = 'GreaterEquals'; PropertyType = 'Version' }
        $allow = { param($c) [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null } }
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion '2.4.1' -Rule $resolved.ConfigMgrRule.Value -Guard $allow).Detection.Operator | Should -Be 'IsEquals'
        [void](Invoke-Grid -Set ([ordered]@{ rule = 'Minimum' }))
        $minimum = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion '2.4.1' -Rule $minimum.ConfigMgrRule.Value -Guard $allow).Detection.Operator | Should -Be 'GreaterEquals'
    }
    It 'tracks the application in One Click, with its interval and its destinations' {
        [void](Invoke-Grid -Set ([ordered]@{ oneclick = 'Yes'; cadence = '30'; tointune = 'Yes'; toconfigmgr = 'No' }))
        $apps = @([pscustomobject]@{ Packager = 'package-demo'; Application = 'Demo App'; Vendor = 'Demo Vendor' }, [pscustomobject]@{ Packager = 'package-other'; Application = 'Other'; Vendor = 'Demo Vendor' })
        $plan = @(Get-OneClickPlan -Apps $apps -Prefs $script:Prefs)
        $plan.Count | Should -Be 1 -Because 'only the tracked application is in the plan'
        $plan[0].Packager | Should -Be 'package-demo'
        $plan[0].Intune | Should -BeTrue
        $plan[0].ConfigMgr | Should -BeFalse
        $plan[0].Source | Should -Be 'Row'
        Get-OneClickCadenceDays -Prefs $script:Prefs -PackagerName 'package-demo' | Should -Be 30
    }
    It 'follows the default destinations when a destination is set to Follow the default' {
        [void](Invoke-Grid -Set ([ordered]@{ oneclick = 'Yes'; towsus = 'Yes' }))
        (Get-OneClickDestinations -Prefs $script:Prefs -PackagerName 'package-demo').WSUS | Should -BeTrue
        [void](Invoke-Grid -Set ([ordered]@{ towsus = 'Follow the default' }) -Json (Get-SavedJson))
        $set = Get-OneClickDestinations -Prefs $script:Prefs -PackagerName 'package-demo'
        $set.WSUS | Should -BeFalse
        $set.Source | Should -Be 'Default'
    }
}

Describe 'The Workbench profile columns' {
    BeforeAll {
        $script:SavedRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
        $env:APP_PACKAGER_WORKBENCH_ROOT = Join-Path $TestDrive 'workbench'
        function Get-ProfileWork { New-ApplicationGridWork -Prefs (New-TestPrefs) }
        function Save-ProfileCell {
            param($Work, $Packager, [string]$Id, [string]$Value)
            $ok = Set-ApplicationGridCell -Setting (Get-GridSetting $Id) -Work $Work -Packager $Packager -Value $Value
            return $ok
        }
    }
    AfterAll { $env:APP_PACKAGER_WORKBENCH_ROOT = $script:SavedRoot }

    It 'creates the profile in the shape of the migration and makes it active while the application is on the default profile' {
        $p = New-TestPackager -Base 'package-prof1'
        $work = Get-ProfileWork
        Save-ProfileCell $work $p 'estimated' '12' | Should -BeTrue
        Save-ProfileCell $work $p 'maximum' '40' | Should -BeTrue
        Save-ProfileCell $work $p 'reboot' 'Force restart' | Should -BeTrue
        Save-ApplicationGridProfile -Work $work
        $definition = Get-ApplicationDefinition -ApplicationId 'catalog:package-prof1' -DataRoot $work.DataRoot
        $profiles = @(Get-Profiles -ApplicationId 'catalog:package-prof1' -DataRoot $work.DataRoot | Where-Object { -not $_.IsDefault })
        $profiles.Count | Should -Be 1
        $stored = Get-Profile -ApplicationId 'catalog:package-prof1' -ProfileId $profiles[0].ProfileId -DataRoot $work.DataRoot
        $stored.MigratedFrom | Should -Be 'package-prof1'
        $stored.Name | Should -Be 'Migrated'
        $definition.ActiveProfileId | Should -Be $stored.ProfileId
        [int]$stored.Timing.EstimatedMinutes | Should -Be 12
        [int]$stored.Timing.MaximumMinutes | Should -Be 40
        $stored.Install.RebootPolicy | Should -Be 'Force restart'
    }
    It 'reads the value back into a new working copy, and removes the key for an empty value or Inherit' {
        $p = New-TestPackager -Base 'package-prof1'
        $work = Get-ProfileWork
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'estimated') -Work $work -Packager $p).Value | Should -Be '12'
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'reboot') -Work $work -Packager $p).Value | Should -Be 'Force restart'
        [void](Save-ProfileCell $work $p 'reboot' 'Inherit')
        [void](Save-ProfileCell $work $p 'estimated' '')
        Save-ApplicationGridProfile -Work $work
        $again = Get-ProfileWork
        $id = @(Get-Profiles -ApplicationId 'catalog:package-prof1' -DataRoot $again.DataRoot | Where-Object { -not $_.IsDefault })[0].ProfileId
        $stored = Get-Profile -ApplicationId 'catalog:package-prof1' -ProfileId $id -DataRoot $again.DataRoot
        $stored.Install.PSObject.Properties['RebootPolicy'] | Should -BeNullOrEmpty
        $stored.Timing.PSObject.Properties['EstimatedMinutes'] | Should -BeNullOrEmpty
        [int]$stored.Timing.MaximumMinutes | Should -Be 40 -Because 'a field that was not cleared stays'
    }
    It 'refuses a runtime that is not a positive whole number, or an estimate above the maximum' {
        $p = New-TestPackager -Base 'package-prof2'
        $work = Get-ProfileWork
        [void](Save-ProfileCell $work $p 'maximum' '30')
        foreach ($bad in 'abc', '0', '-2') { [void](Save-ProfileCell $work $p 'estimated' $bad) }
        [void](Save-ProfileCell $work $p 'estimated' '31')
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'estimated') -Work $work -Packager $p).Value | Should -Be ''
        [void](Save-ProfileCell $work $p 'estimated' '30')
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'estimated') -Work $work -Packager $p).Value | Should -Be '30'
    }
    It 'creates no profile when every value is empty' {
        $p = New-TestPackager -Base 'package-prof3'
        $work = Get-ProfileWork
        [void](Save-ProfileCell $work $p 'reboot' 'Inherit')
        Save-ApplicationGridProfile -Work $work
        @(Get-Profiles -ApplicationId 'catalog:package-prof3' -DataRoot $work.DataRoot | Where-Object { -not $_.IsDefault }).Count | Should -Be 0
    }
    It 'greys the cells and names the profile when another profile is active, and writes nothing' {
        $id = 'catalog:package-prof4'
        $root = (Get-ProfileWork).DataRoot
        $other = New-WorkbenchProfileObject -ApplicationId $id -Name 'Pilot'
        $saved = Save-Profile -Profile $other -DataRoot $root -SetActive
        $p = New-TestPackager -Base 'package-prof4' -Application 'Prof Four'
        $work = Get-ProfileWork
        foreach ($setting in 'estimated', 'maximum', 'reboot') {
            (Get-ApplicationGridCell -Setting (Get-GridSetting $setting) -Work $work -Packager $p).Reason | Should -Be 'The profile "Pilot" is active for Prof Four, and a run reads the active profile only. Change it in the Application Workbench.'
        }
        Save-ProfileCell $work $p 'reboot' 'No reboot' | Should -BeFalse
        Save-ApplicationGridProfile -Work $work
        @(Get-Profiles -ApplicationId $id -DataRoot $root | Where-Object { -not $_.IsDefault }).Count | Should -Be 1
        (Get-ApplicationDefinition -ApplicationId $id -DataRoot $root).ActiveProfileId | Should -Be $saved.ProfileId
    }
    It 'edits the migrated profile while it is the active profile' {
        $p = New-TestPackager -Base 'package-prof1'
        $work = Get-ProfileWork
        (Get-ApplicationGridCell -Setting (Get-GridSetting 'maximum') -Work $work -Packager $p).Reason | Should -BeNullOrEmpty
    }
    It 'is what the stage reads: the effective settings and the install block of the profile' {
        $p = New-TestPackager -Base 'package-prof5'
        $work = Get-ProfileWork
        [void](Save-ProfileCell $work $p 'estimated' '7')
        [void](Save-ProfileCell $work $p 'maximum' '55')
        [void](Save-ProfileCell $work $p 'reboot' 'Soft reboot')
        Save-ApplicationGridProfile -Work $work
        $definition = Get-ApplicationDefinition -ApplicationId 'catalog:package-prof5' -DataRoot $work.DataRoot
        $profile = Get-Profile -ApplicationId 'catalog:package-prof5' -ProfileId $definition.ActiveProfileId -DataRoot $work.DataRoot
        $settings = Resolve-EffectiveSettings -GlobalDefaults ([pscustomobject]@{ EstimatedRuntimeMins = 15; MaximumRuntimeMins = 30 }) -BaseManifest $null -Profile $profile
        $settings['EstimatedMinutes'].Value | Should -Be 7
        $settings['EstimatedMinutes'].Source | Should -Be 'Profile'
        $settings['MaximumMinutes'].Value | Should -Be 55
        $profile.Install.RebootPolicy | Should -Be 'Soft reboot'
    }
}

Describe 'The page' {
    It 'is the one grid, under Applications, with the three defaults above it' {
        $text = $script:LauncherText
        $text | Should -Match "Id = 'allapps'; Title = 'All Applications'; Parent = 'applications'"
        $text | Should -Not -Match "Id = 'supersedence'; Title = 'Supersedence'" -Because 'no second grid competes with the merged page'
        $text | Should -Not -Match 'New-OptionsSupersedenceEditor'
        ([regex]::Matches($text, "Kind = 'Custom'; FullWidth = \`$true")).Count | Should -Be 1
        foreach ($id in 'sup.cmRule', 'sup.inRule', 'sup.mode', 'sup.apps') { $text | Should -Match ("Id = '" + [regex]::Escape($id) + "'") }
    }
    It 'builds the grid, the chooser and the bulk bar from the registry' {
        $text = $script:LauncherText
        $text | Should -Match '<!--COLUMNS-->'
        $text | Should -Match "'Apply', 'Apply'"
        $text | Should -Match "'None', 'None'"
        $text | Should -Match "-Label 'Columns'"
        foreach ($label in 'Check all', 'Check none') { $text | Should -Match ("'" + $label + "'") }
        $text | Should -Match 'sup\.apps\.search'
        $text | Should -Match "'Checked', 'Filtered', 'All'"
    }
    It 'commits the grid after the other pages' {
        $script:LauncherText | Should -Match "Sort-Object \{ \`$_ -eq 'allapps' \}"
    }
}
