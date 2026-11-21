#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The grid of all applications in Options (Applications, All Applications): what each column offers, what each bulk action
    does, what the grid writes, and that the packaging code reads and uses the value that the grid wrote.

.DESCRIPTION
    Windows PowerShell 5.1 and PowerShell 7. The cases open no window and read no installed product. They build the rows
    from the catalog headers, run the pure functions of the grid, write the preferences the way Save-Preferences does, read
    them back with Read-Preferences, and follow each value into the code that packages:
      Install scope      -> APP_PACKAGER_INSTALL_MODE of the child process
      Package source type -> APP_PACKAGER_VARIANTS of the child process
      Version in title   -> the resolved application title
      Supersedence       -> Resolve-Supersedence, and the title lock
      Version rule       -> the operator of the version clause in the detection
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerWorkbench.psd1') -Force -DisableNameChecking -ErrorAction SilentlyContinue

    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $script:LauncherText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-apppackager.ps1') -Raw
    foreach ($name in 'Get-ApplicationGridLabel', 'Initialize-ApplicationGridType', 'Get-ApplicationGridItemList', 'Get-ApplicationGridCapability', 'New-ApplicationGridRow',
        'Set-ApplicationGridBulkValue', 'Update-ApplicationGridPreference', 'Get-ApplicationGridNotice',
        'Resolve-FirstRunCompleted', 'Read-Preferences', 'Get-WsusClassificationNameList',
        'ConvertTo-InstallModeValue', 'ConvertTo-VariantsJson', 'Get-InstallModesMapForContext', 'Get-VariantsMapForContext', 'Get-TitleModesMapForContext',
        'Set-PackagerEnvironment', 'Get-SupersedenceLockedTitleMode') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "start-apppackager.ps1 no longer defines $name." }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:PrefPath = Join-Path $TestDrive 'preferences.json'
    function Get-PreferencesPath { $script:PrefPath }

    $script:Label = Get-ApplicationGridLabel

    # A catalog row from the header lines of one packager file, the way Get-Packagers reads them.
    function Get-CatalogRow {
        foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1' -File | Sort-Object Name)) {
            $head = @(Get-Content -LiteralPath $file.FullName -TotalCount 60)
            $variants = @(); $modes = @(); $app = $file.BaseName; $vendor = 'Vendor'
            foreach ($line in $head) {
                if ($line -match '^\s*(?:#\s*)?SupportsVariants\s*:\s*(.+?)\s*$') { $variants = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() }) }
                if ($line -match '^\s*(?:#\s*)?SupportsInstallModes\s*:\s*(.+?)\s*$') { $modes = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() }) }
                if ($line -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$') { $app = $Matches[1] }
                if ($line -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$') { $vendor = $Matches[1] }
            }
            [pscustomobject]@{ Script = $file.Name; Application = $app; Vendor = $vendor; SupportsVariants = $variants; SupportsInstallModes = $modes }
        }
    }
    function New-TestPackager {
        param([string]$Base = 'package-demo', [string[]]$Variants = @(), [string[]]$Modes = @(), [string]$Application = 'Demo App')
        [pscustomobject]@{ Script = ($Base + '.ps1'); Application = $Application; Vendor = 'Demo Vendor'; SupportsVariants = $Variants; SupportsInstallModes = $Modes }
    }
    function New-TestPrefs {
        # The preferences the way Read-Preferences returns them, written and read through a file.
        param([string]$Json = '{}')
        Set-Content -LiteralPath $script:PrefPath -Value $Json -Encoding UTF8
        return (Read-Preferences)
    }
    function Save-TestPrefs {
        param($Prefs)
        $Prefs | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:PrefPath -Encoding UTF8
        return (Read-Preferences)
    }
}

Describe 'What each column offers' {
    It 'opens the package source type only for a packager that declares an architecture split' {
        $cap = Get-ApplicationGridCapability -Packager (New-TestPackager -Variants @('Architecture'))
        $cap.Architecture.Enabled | Should -BeTrue
        $cap.Architecture.Items | Should -Be @('Packager default', 'x64 and ARM64')
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Variants @('Language'))).Architecture.Enabled | Should -BeFalse
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Variants @('Network'))).Architecture.Enabled | Should -BeFalse
        (Get-ApplicationGridCapability -Packager (New-TestPackager)).Architecture.Enabled | Should -BeFalse
    }

    It 'shows the own architecture of a sibling packager and does not open it' {
        foreach ($case in @(@('package-temurin-jdk8-x64', 'x64'), @('package-temurin-jdk8-x86', 'x86'), @('package-citrixworkspace-ltsr-arm64', 'ARM64'))) {
            $cap = Get-ApplicationGridCapability -Packager (New-TestPackager -Base $case[0])
            $cap.Architecture.Enabled | Should -BeFalse
            $cap.Architecture.Items | Should -Be @($case[1])
        }
    }

    It 'opens the install scope only for a packager that declares both modes' {
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Modes @('CurrentUser', 'AllUsers'))).InstallScope.Items | Should -Be @('Packager default', 'System', 'User')
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Modes @('CurrentUser', 'AllUsers'))).InstallScope.Enabled | Should -BeTrue
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Modes @('AllUsers'))).InstallScope.Enabled | Should -BeFalse
        (Get-ApplicationGridCapability -Packager (New-TestPackager -Modes @('CurrentUser'))).InstallScope.Enabled | Should -BeFalse
        (Get-ApplicationGridCapability -Packager (New-TestPackager)).InstallScope.Enabled | Should -BeFalse
    }

    It 'follows the header tags of the real catalog' {
        $catalog = @(Get-CatalogRow)
        $catalog.Count | Should -BeGreaterThan 300
        $rows = @($catalog | ForEach-Object { New-ApplicationGridRow -Packager $_ })
        $rows.Count | Should -Be $catalog.Count
        @($rows | Where-Object { $_.ArchEnabled }).Count | Should -Be @($catalog | Where-Object { $_.SupportsVariants -contains 'Architecture' }).Count
        @($rows | Where-Object { $_.ScopeEnabled }).Count | Should -Be @($catalog | Where-Object { $_.SupportsInstallModes -contains 'AllUsers' -and $_.SupportsInstallModes -contains 'CurrentUser' }).Count
        @($rows | Where-Object { $_.ArchEnabled }).Count | Should -BeGreaterThan 0
        @($rows | Where-Object { $_.ScopeEnabled }).Count | Should -BeGreaterThan 0
        @($rows | Where-Object { -not $_.ArchEnabled -and $_.Arch -ne 'Packager default' }).Count | Should -BeGreaterThan 0 -Because 'sibling packagers show their own architecture'
    }

    It 'lists one row for each application, with its identifier' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-chrome' -Application 'Google Chrome')
        $row.Base | Should -Be 'package-chrome'
        $row.Id | Should -Be 'catalog:package-chrome'
        $row.Application | Should -Be 'Google Chrome'
        $row.Vendor | Should -Be 'Demo Vendor'
    }

    It 'ships Follow the default and Packager default, supersedence Off and the version rule from the default' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers'))
        $row.Arch | Should -Be 'Packager default'
        $row.Scope | Should -Be 'Packager default'
        $row.Supersedence | Should -Be 'Follow the default'
        $row.Title | Should -Be 'Follow the default'
        $row.Rule | Should -Be 'Follow the default'
        $row.TitleLocked | Should -BeFalse
    }

    It 'gives every combo box of the grid one shared list of items' {
        $a = New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a')
        $b = New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b')
        [object]::ReferenceEquals($a.TitleItems, $b.TitleItems) | Should -BeTrue
        [object]::ReferenceEquals($a.SupersedenceItems, $b.SupersedenceItems) | Should -BeTrue
    }

    It 'ignores a null that a combo box writes while it changes its items' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers'))
        $row.Arch = 'x64 and ARM64'; $row.Scope = 'User'; $row.Supersedence = 'Replace'; $row.Rule = 'Exact'
        # PowerShell turns $null into an empty string for a string property; a binding writes a real null.
        $null1 = [NullString]::Value
        $row.Arch = $null1; $row.Scope = $null1; $row.Supersedence = $null1; $row.Rule = $null1; $row.TitleShown = $null1
        $row.Arch | Should -Be 'x64 and ARM64'
        $row.Scope | Should -Be 'User'
        $row.Supersedence | Should -Be 'Replace'
        $row.Rule | Should -Be 'Exact'
    }
}

Describe 'The title lock' {
    It 'shows Include version and cannot change while supersedence is on for the application' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager)
        $row.Title = 'No version'
        $row.Supersedence = 'Replace'
        $row.TitleLocked | Should -BeTrue
        $row.TitleEnabled | Should -BeFalse
        $row.TitleShown | Should -Be 'Include version'
        $row.TitleShown = 'Follow the default'
        $row.Title | Should -Be 'No version' -Because 'the lock never writes the stored choice'
        $row.TitleTip | Should -Match 'separate applications'
    }

    It 'shows the stored choice again when supersedence is off' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager)
        $row.Title = 'No version'
        $row.Supersedence = 'Replace and uninstall'
        $row.Supersedence = 'Off'
        $row.TitleLocked | Should -BeFalse
        $row.TitleShown | Should -Be 'No version'
        $row.TitleShown = 'Include version'
        $row.Title | Should -Be 'Include version'
    }

    It 'follows the default for an application that follows the default' {
        $row = New-ApplicationGridRow -Packager (New-TestPackager) -DefaultSupersedence 'Replace'
        $row.TitleLocked | Should -BeTrue
        $row.DefaultSupersedence = 'Off'
        $row.TitleLocked | Should -BeFalse
        $row.DefaultSupersedence = 'Replace'
        $row.Supersedence = 'Off'
        $row.TitleLocked | Should -BeFalse -Because 'an own value of Off outranks a default that is on'
    }

    It 'tells the user how many applications are locked' {
        $rows = @(1..3 | ForEach-Object { New-ApplicationGridRow -Packager (New-TestPackager -Base "package-$_") })
        Get-ApplicationGridNotice -Row $rows | Should -BeNullOrEmpty
        $rows[0].Supersedence = 'Replace'; $rows[1].Supersedence = 'Replace and uninstall'
        Get-ApplicationGridNotice -Row $rows | Should -Match '^Supersedence is on for 2 application\(s\)\.'
        Get-ApplicationGridNotice -Row $rows | Should -Match 'separate applications'
    }
}

Describe 'The bulk actions' {
    BeforeEach {
        $script:Rows = @(
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a' -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')))
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b' -Modes @('CurrentUser', 'AllUsers')))
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-c'))
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-d-x64'))
        )
    }

    It 'changes only the checked applications for the scope Checked' {
        $script:Rows[0].Checked = $true; $script:Rows[2].Checked = $true
        $r = Set-ApplicationGridBulkValue -Column Supersedence -Scope Checked -Value 'Replace' -Row $script:Rows
        $r.Changed | Should -Be 2
        @($script:Rows | ForEach-Object { $_.Supersedence }) | Should -Be @('Replace', 'Follow the default', 'Replace', 'Follow the default')
        $r.Default | Should -BeNullOrEmpty -Because 'Checked leaves the default alone'
    }

    It 'sets the default and removes the own values for the scope All of supersedence' {
        $script:Rows[1].Supersedence = 'Off'
        $r = Set-ApplicationGridBulkValue -Column Supersedence -Scope All -Value 'Replace' -Row $script:Rows
        $r.Default | Should -Be 'Replace'
        @($script:Rows | ForEach-Object { $_.Supersedence } | Select-Object -Unique) | Should -Be @('Follow the default')
        $r.Changed | Should -Be 1
    }

    It 'resets supersedence to Off and the rule to Minimum for the scope None' {
        $script:Rows[0].Supersedence = 'Replace'; $script:Rows[1].Rule = 'Exact'
        (Set-ApplicationGridBulkValue -Column Supersedence -Scope None -Row $script:Rows).Default | Should -Be 'Off'
        $r = Set-ApplicationGridBulkValue -Column VersionRule -Scope None -Row $script:Rows
        $r.Default | Should -Be 'Minimum'
        @($script:Rows | ForEach-Object { $_.Supersedence } | Select-Object -Unique) | Should -Be @('Follow the default')
        @($script:Rows | ForEach-Object { $_.Rule } | Select-Object -Unique) | Should -Be @('Follow the default')
    }

    It 'sets the version rule default for All and an own value for Checked' {
        $script:Rows[1].Checked = $true
        (Set-ApplicationGridBulkValue -Column VersionRule -Scope Checked -Value 'Exact' -Row $script:Rows).Changed | Should -Be 1
        $script:Rows[1].Rule | Should -Be 'Exact'
        $r = Set-ApplicationGridBulkValue -Column VersionRule -Scope All -Value 'Exact' -Row $script:Rows
        $r.Default | Should -Be 'Exact'
        $script:Rows[1].Rule | Should -Be 'Follow the default'
    }

    It 'skips an application that cannot take the install scope or the package source type' {
        $r = Set-ApplicationGridBulkValue -Column InstallScope -Scope All -Value 'User' -Row $script:Rows
        @($script:Rows | ForEach-Object { $_.Scope }) | Should -Be @('User', 'User', 'Packager default', 'Packager default')
        $r.Changed | Should -Be 2
        $r = Set-ApplicationGridBulkValue -Column Architecture -Scope All -Value 'x64 and ARM64' -Row $script:Rows
        @($script:Rows | ForEach-Object { $_.Arch }) | Should -Be @('x64 and ARM64', 'Packager default', 'Packager default', 'x64')
        $r.Changed | Should -Be 1
        (Set-ApplicationGridBulkValue -Column Architecture -Scope None -Row $script:Rows).Changed | Should -Be 1
        $script:Rows[0].Arch | Should -Be 'Packager default'
    }

    It 'sets the title choice of the checked applications even while it is locked' {
        $script:Rows[0].Supersedence = 'Replace'
        $script:Rows[0].Checked = $true
        (Set-ApplicationGridBulkValue -Column Title -Scope Checked -Value 'No version' -Row $script:Rows).Changed | Should -Be 1
        $script:Rows[0].Title | Should -Be 'No version'
        $script:Rows[0].TitleShown | Should -Be 'Include version'
        (Set-ApplicationGridBulkValue -Column Title -Scope None -Row $script:Rows).Changed | Should -Be 1
        $script:Rows[0].Title | Should -Be 'Follow the default'
    }

    It 'refuses a value that is not a value of the column' {
        { Set-ApplicationGridBulkValue -Column InstallScope -Scope All -Value 'Replace' -Row $script:Rows } | Should -Throw
        { Set-ApplicationGridBulkValue -Column VersionRule -Scope All -Value 'Follow the default' -Row $script:Rows } | Should -Throw
    }

    It 'accepts an empty list of rows' {
        (Set-ApplicationGridBulkValue -Column Title -Scope All -Value 'No version' -Row @()).Changed | Should -Be 0
    }
}

Describe 'What the grid writes' {
    It 'stores each column under the key that the packaging code reads' {
        $prefs = New-TestPrefs
        $rows = @(New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-demo' -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')) -Conditions $prefs.DeploymentConditions.Apps -Overrides $prefs.Supersedence.TitleOverrides)
        $rows[0].Arch = 'x64 and ARM64'; $rows[0].Scope = 'User'; $rows[0].Title = 'No version'; $rows[0].Supersedence = 'Replace and uninstall'; $rows[0].Rule = 'Exact'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $entry = $prefs.DeploymentConditions.Apps.'package-demo'
        $entry.Split | Should -Be 'Architecture'
        $entry.InstallMode | Should -Be 'CurrentUser'
        $entry.TitleMode | Should -Be 'NoVersion'
        $own = $prefs.Supersedence.TitleOverrides.'catalog:package-demo'
        $own.Mode | Should -Be 'Replace and uninstall'
        $own.VersionRule | Should -Be 'Exact'
    }

    It 'maps System to AllUsers and the default to no value' {
        $prefs = New-TestPrefs
        $rows = @(New-ApplicationGridRow -Packager (New-TestPackager -Modes @('CurrentUser', 'AllUsers')))
        $rows[0].Scope = 'System'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $prefs.DeploymentConditions.Apps.'package-demo'.InstallMode | Should -Be 'AllUsers'
        $rows[0].Scope = 'Packager default'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $prefs.DeploymentConditions.Apps.PSObject.Properties['package-demo'] | Should -BeNullOrEmpty -Because 'an entry with no value is not stored'
    }

    It 'removes the own value of an application that follows the default again' {
        $prefs = New-TestPrefs '{ "Supersedence": { "TitleOverrides": { "catalog:package-demo": { "Mode": "Replace", "VersionRule": "Exact" } } } }'
        $rows = @(New-ApplicationGridRow -Packager (New-TestPackager) -Overrides $prefs.Supersedence.TitleOverrides)
        $rows[0].Supersedence | Should -Be 'Replace'
        $rows[0].Rule | Should -Be 'Exact'
        $rows[0].Supersedence = 'Follow the default'; $rows[0].Rule = 'Follow the default'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $prefs.Supersedence.TitleOverrides.PSObject.Properties['catalog:package-demo'] | Should -BeNullOrEmpty
    }

    It 'keeps what the grid does not own' {
        $json = '{ "DeploymentConditions": { "Apps": { "package-demo": { "Architecture": "ARM64", "Languages": ["de-DE"], "Network": "VpnOnly", "Split": "Language", "TitleMode": "", "InstallMode": "" },' +
        ' "package-gone": { "Architecture": "x64" } } }, "Supersedence": { "TitleOverrides": { "catalog:package-demo": { "Mode": "Replace" }, "custom:other": { "Mode": "Off" } } } }'
        $prefs = New-TestPrefs $json
        $rows = @(New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Language', 'Architecture')) -Conditions $prefs.DeploymentConditions.Apps -Overrides $prefs.Supersedence.TitleOverrides)
        $rows[0].Arch | Should -Be 'Packager default' -Because 'a Language split is not a package source type'
        $rows[0].Title = 'Include version'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $entry = $prefs.DeploymentConditions.Apps.'package-demo'
        $entry.Architecture | Should -Be 'ARM64'
        $entry.Network | Should -Be 'VpnOnly'
        @($entry.Languages) | Should -Be @('de-DE')
        $entry.Split | Should -Be 'Language' -Because 'the source type was not touched'
        $entry.TitleMode | Should -Be 'IncludeVersion'
        $prefs.DeploymentConditions.Apps.'package-gone'.Architecture | Should -Be 'x64'
        $prefs.Supersedence.TitleOverrides.'custom:other'.Mode | Should -Be 'Off'
    }

    It 'turns a split back to none when the source type returns to the default' {
        $prefs = New-TestPrefs '{ "DeploymentConditions": { "Apps": { "package-demo": { "Split": "Architecture", "Network": "VpnOnly" } } } }'
        $rows = @(New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture')) -Conditions $prefs.DeploymentConditions.Apps)
        $rows[0].Arch | Should -Be 'x64 and ARM64'
        $rows[0].Arch = 'Packager default'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $prefs.DeploymentConditions.Apps.'package-demo'.Split | Should -Be 'None'
        $prefs.DeploymentConditions.Apps.'package-demo'.Network | Should -Be 'VpnOnly'
    }

    It 'reads back the same grid after it was saved' {
        $prefs = New-TestPrefs
        $rows = @(
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a' -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')))
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b'))
        )
        $rows[0].Arch = 'x64 and ARM64'; $rows[0].Scope = 'System'; $rows[0].Rule = 'Exact'
        $rows[1].Title = 'No version'; $rows[1].Supersedence = 'Off'
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $saved = Save-TestPrefs -Prefs $prefs
        $again = @(
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a' -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')) -Conditions $saved.DeploymentConditions.Apps -Overrides $saved.Supersedence.TitleOverrides)
            (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b') -Conditions $saved.DeploymentConditions.Apps -Overrides $saved.Supersedence.TitleOverrides)
        )
        $again[0].Arch | Should -Be 'x64 and ARM64'
        $again[0].Scope | Should -Be 'System'
        $again[0].Rule | Should -Be 'Exact'
        $again[1].Title | Should -Be 'No version'
        $again[1].Supersedence | Should -Be 'Off'
    }

    It 'reads the keys that the earlier grid and the command line wrote' {
        $json = '{ "DeploymentConditions": { "Apps": { "package-demo": { "Architecture": "Any", "Split": "Architecture", "InstallMode": "CurrentUser", "TitleMode": "IncludeVersion" } } },' +
        ' "Supersedence": { "Mode": "Replace", "TitleOverrides": { "catalog:package-demo": { "VersionRule": "Exact", "Mode": "Off" } } } }'
        $prefs = New-TestPrefs $json
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Variants @('Architecture') -Modes @('CurrentUser', 'AllUsers')) -Conditions $prefs.DeploymentConditions.Apps -Overrides $prefs.Supersedence.TitleOverrides -DefaultSupersedence $prefs.Supersedence.Mode
        $row.Arch | Should -Be 'x64 and ARM64'
        $row.Scope | Should -Be 'User'
        $row.Title | Should -Be 'Include version'
        $row.Supersedence | Should -Be 'Off'
        $row.Rule | Should -Be 'Exact'
    }
}

Describe 'The packaging code uses the value that the grid wrote' {
    BeforeAll {
    function Invoke-Grid {
        # Sets columns of one application, saves the preferences and reads them back, as the Options window and the next start do.
        param([hashtable]$Set, [string]$Json = '{}', [string]$Base = 'package-demo', [string[]]$Variants = @('Architecture'), [string[]]$Modes = @('CurrentUser', 'AllUsers'))
        $prefs = New-TestPrefs $Json
        $row = New-ApplicationGridRow -Packager (New-TestPackager -Base $Base -Variants $Variants -Modes $Modes) -Conditions $prefs.DeploymentConditions.Apps -Overrides $prefs.Supersedence.TitleOverrides -DefaultSupersedence $prefs.Supersedence.Mode
        foreach ($key in $Set.Keys) { $row.$key = $Set[$key] }
        Update-ApplicationGridPreference -Prefs $prefs -Row @($row)
        $script:Prefs = Save-TestPrefs -Prefs $prefs
        return $script:Prefs
    }
    }
    BeforeEach {
        $script:SavedMode = $env:APP_PACKAGER_INSTALL_MODE
        $script:SavedVariants = $env:APP_PACKAGER_VARIANTS
        $env:APP_PACKAGER_INSTALL_MODE = $null
        $env:APP_PACKAGER_VARIANTS = $null
    }
    AfterEach {
        $env:APP_PACKAGER_INSTALL_MODE = $script:SavedMode
        $env:APP_PACKAGER_VARIANTS = $script:SavedVariants
    }

    It 'passes Install scope User to the child process as APP_PACKAGER_INSTALL_MODE=CurrentUser' {
        [void](Invoke-Grid -Set @{ Scope = 'User' })
        $map = Get-InstallModesMapForContext
        $map['package-demo'] | Should -Be 'CurrentUser'
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -InstallMode ([string]$map['package-demo'])
        $start.EnvironmentVariables['APP_PACKAGER_INSTALL_MODE'] | Should -Be 'CurrentUser'
        $env:APP_PACKAGER_INSTALL_MODE = $start.EnvironmentVariables['APP_PACKAGER_INSTALL_MODE']
        Get-RequestedInstallMode | Should -Be 'CurrentUser'
    }

    It 'passes Install scope System as AllUsers, and the default as nothing' {
        [void](Invoke-Grid -Set @{ Scope = 'System' })
        (Get-InstallModesMapForContext)['package-demo'] | Should -Be 'AllUsers'
        [void](Invoke-Grid -Set @{ Scope = 'Packager default' })
        (Get-InstallModesMapForContext).ContainsKey('package-demo') | Should -BeFalse
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -InstallMode ''
        $start.EnvironmentVariables.ContainsKey('APP_PACKAGER_INSTALL_MODE') | Should -BeFalse
        Get-RequestedInstallMode | Should -BeNullOrEmpty
    }

    It 'passes the package source type x64 and ARM64 as an Architecture split' {
        [void](Invoke-Grid -Set @{ Arch = 'x64 and ARM64' })
        $json = (Get-VariantsMapForContext)['package-demo']
        $start = New-Object System.Diagnostics.ProcessStartInfo
        Set-PackagerEnvironment -StartInfo $start -VariantsJson $json
        $env:APP_PACKAGER_VARIANTS = $start.EnvironmentVariables['APP_PACKAGER_VARIANTS']
        (Get-RequestedPackagerVariants).Split | Should -Be 'Architecture'
        [void](Invoke-Grid -Set @{ Arch = 'Packager default' } -Json (Get-Content -LiteralPath $script:PrefPath -Raw))
        (Get-VariantsMapForContext).ContainsKey('package-demo') | Should -BeFalse
    }

    It 'resolves the application title from Version in title' {
        [void](Invoke-Grid -Set @{ Title = 'No version' })
        $mode = (Get-TitleModesMapForContext)['package-demo']
        $mode | Should -Be 'NoVersion'
        Get-PackagedApplicationName -AppName 'Demo App 2.4.1' -Version '2.4.1' -Mode $mode | Should -Be 'Demo App'
        [void](Invoke-Grid -Set @{ Title = 'Include version' })
        $mode = (Get-TitleModesMapForContext)['package-demo']
        $mode | Should -Be 'IncludeVersion'
        Get-PackagedApplicationName -AppName 'Demo App' -Version '2.4.1' -Mode $mode | Should -Be 'Demo App - 2.4.1'
        [void](Invoke-Grid -Set @{ Title = 'Follow the default' })
        (Get-TitleModesMapForContext).ContainsKey('package-demo') | Should -BeFalse
    }

    It 'locks the title to the version while supersedence is on, whatever the stored choice is' {
        Mock Get-SupersedenceSettings { $script:Prefs.Supersedence }
        Mock Get-SupersedenceCommandLine { $null }
        [void](Invoke-Grid -Set @{ Title = 'No version'; Supersedence = 'Replace' })
        $stored = (Get-TitleModesMapForContext)['package-demo']
        $stored | Should -Be 'NoVersion' -Because 'the choice of the operator stays stored'
        $path = 'C:\Packagers\package-demo.ps1'
        Get-SupersedenceLockedTitleMode -PackagerPath $path -TitleMode $stored | Should -Be 'IncludeVersion'
        [void](Invoke-Grid -Set @{ Title = 'No version'; Supersedence = 'Off' })
        Get-SupersedenceLockedTitleMode -PackagerPath $path -TitleMode 'NoVersion' | Should -Be 'NoVersion'
        [void](Invoke-Grid -Set @{ Title = 'No version' } -Json '{ "Supersedence": { "Mode": "Replace" } }')
        Get-SupersedenceLockedTitleMode -PackagerPath $path -TitleMode 'NoVersion' | Should -Be 'IncludeVersion' -Because 'the application follows a default that is on'
    }

    It 'resolves the supersedence of the application, and the default for the others' {
        [void](Invoke-Grid -Set @{ Supersedence = 'Replace and uninstall'; Rule = 'Exact' })
        $resolved = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        $resolved.Mode.Value | Should -Be 'ReplaceAndUninstall'
        $resolved.Mode.Source | Should -Be 'TitleOverride'
        $other = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-other'
        $other.Mode.Value | Should -Be 'Off'
        $other.ConfigMgrRule.Value | Should -Be 'Minimum'
    }

    It 'writes the operator of the detection from Version rule' {
        [void](Invoke-Grid -Set @{ Rule = 'Exact' })
        $resolved = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        $resolved.ConfigMgrRule.Value | Should -Be 'Exact'
        $resolved.IntuneRule.Value | Should -Be 'Exact'
        $resolved.ConfigMgrRule.Source | Should -Be 'TitleOverride'
        $clause = [pscustomobject]@{ Type = 'File'; FilePath = '%ProgramFiles%\Demo'; FileName = 'demo.exe'; ExpectedValue = '2.4.1'; Operator = 'GreaterEquals'; PropertyType = 'Version' }
        $allow = { param($c) [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null } }
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion '2.4.1' -Rule $resolved.ConfigMgrRule.Value -Guard $allow).Detection.Operator | Should -Be 'IsEquals'
        [void](Invoke-Grid -Set @{ Rule = 'Minimum' })
        $minimum = Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo'
        (Set-DetectionVersionRule -Detection $clause -PackagedVersion '2.4.1' -Rule $minimum.ConfigMgrRule.Value -Guard $allow).Detection.Operator | Should -Be 'GreaterEquals'
        [void](Invoke-Grid -Set @{ Rule = 'Follow the default' })
        (Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-demo').ConfigMgrRule.Source | Should -Be 'SiteDefault'
    }

    It 'resolves the default that the scope All of the bulk action wrote' {
        $prefs = New-TestPrefs
        $rows = @((New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-a')), (New-ApplicationGridRow -Packager (New-TestPackager -Base 'package-b')))
        $result = Set-ApplicationGridBulkValue -Column VersionRule -Scope All -Value 'Exact' -Row $rows
        $prefs.Supersedence.ConfigMgrVersionRule = $result.Default
        $prefs.Supersedence.IntuneVersionRule = $result.Default
        Update-ApplicationGridPreference -Prefs $prefs -Row $rows
        $script:Prefs = Save-TestPrefs -Prefs $prefs
        (Resolve-Supersedence -Settings $script:Prefs.Supersedence -ApplicationId 'catalog:package-b').ConfigMgrRule.Value | Should -Be 'Exact'
    }
}

Describe 'The page' {
    It 'is the one grid, under Applications, with the three defaults above it' {
        $tokens = $null; $errors = $null
        $text = $script:LauncherText
        $text | Should -Match "Id = 'allapps'; Title = 'All Applications'; Parent = 'applications'"
        $text | Should -Not -Match "Id = 'supersedence'; Title = 'Supersedence'" -Because 'no second grid competes with the merged page'
        $text | Should -Not -Match 'New-OptionsSupersedenceEditor'
        ([regex]::Matches($text, "Kind = 'Custom'; FullWidth = \`$true")).Count | Should -Be 1
        foreach ($id in 'sup.cmRule', 'sup.inRule', 'sup.mode', 'sup.apps') { $text | Should -Match ("Id = '" + [regex]::Escape($id) + "'") }
    }

    It 'has the five columns, the bulk buttons and a tip on every header' {
        $text = $script:LauncherText
        foreach ($header in '', 'Application', 'Vendor', 'Source Type', 'Install Scope', 'Supersedence', 'Version in Title', 'Version Rule') {
            $text | Should -Match ('Header="' + [regex]::Escape($header) + '"')
        }
        foreach ($button in 'Checked applications', 'All applications', 'No applications', 'Check all', 'Check none') { $text | Should -Match ("'" + $button + "'") }
        foreach ($header in 'Source Type', 'Install Scope', 'Supersedence', 'Version in Title', 'Version Rule') { $text | Should -Match ("'" + $header + "'\s+= '") }
    }

    It 'keeps the label words of the row class and of the functions equal' {
        $label = Get-ApplicationGridLabel
        Initialize-ApplicationGridType
        [AppPackagerGridRow]::FollowLabel | Should -Be $label.Follow
        [AppPackagerGridRow]::OffLabel | Should -Be $label.Off
        [AppPackagerGridRow]::IncludeLabel | Should -Be $label.Include
    }
}
