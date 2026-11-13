#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every preference key has a place in the Options window, or a documented reason that it has none.

.DESCRIPTION
    The cases read the defaults of Read-Preferences, Read-CwaSwitches and Read-TvHostConfig, and the field
    definitions of the Options window. They fail when a key has no field and no entry in
    Get-OptionsNonUiPreference, when a field points at a key that does not exist, and when a setting of the
    old Options panels lost its home. An older preference file must load unchanged.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:LauncherPath = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:LauncherPath, [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    # The functions load into this scope, so the loop does the dot-sourcing.
    foreach ($name in 'Resolve-FirstRunCompleted', 'Read-Preferences', 'Get-WsusClassificationNameList', 'Read-CwaSwitches', 'Get-CwaSwitchesPath', 'Read-TvHostConfig', 'Get-TvHostConfigPath', 'Get-OptionsNonUiPreference', 'Get-OptionsSectionDefinition', 'Save-Preferences') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "start-apppackager.ps1 no longer defines $name." }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:PrefPath = Join-Path $TestDrive 'none.json'
    function Get-PreferencesPath { $script:PrefPath }
    function Get-CwaSwitchesPath { Join-Path $TestDrive 'none-cwa.json' }
    function Get-TvHostConfigPath { Join-Path $TestDrive 'none-tv.json' }

    # A key whose value is a map with keys that the user adds is one leaf.
    $script:MapKey = @('LocalSourceFolders', 'DeploymentConditions.Apps', 'CommandOverrides.Apps', 'AppFlow.CadenceOverrides', 'AppFlow.Destinations')
    function Get-LeafPath {
        param($Object, [string]$Prefix = '')
        foreach ($p in $Object.PSObject.Properties) {
            $path = $(if ($Prefix) { $Prefix + '.' + $p.Name } else { $p.Name })
            $v = $p.Value
            if ($v -is [pscustomobject] -and $script:MapKey -notcontains $path -and @($v.PSObject.Properties).Count -gt 0) { Get-LeafPath -Object $v -Prefix $path }
            else { $path }
        }
    }

    $script:Defaults = Read-Preferences
    $script:PrefLeaf = @(Get-LeafPath -Object $script:Defaults)
    $script:CwaLeaf = @(Get-LeafPath -Object (Read-CwaSwitches) | ForEach-Object { 'cwa:' + $_ })
    $script:TvLeaf = @(Get-LeafPath -Object (Read-TvHostConfig) | ForEach-Object { 'tv:' + $_ })
    $script:Definition = @(Get-OptionsSectionDefinition)
    $script:FieldPref = @($script:Definition | ForEach-Object { $_.Fields } | Where-Object { $_.Pref } | ForEach-Object { [string]$_.Pref })
    $script:PanelPref = @($script:Definition | Where-Object { $_.Prefs } | ForEach-Object { $_.Prefs } | ForEach-Object { [string]$_ })
    $script:NonUi = Get-OptionsNonUiPreference

    function Test-Homed {
        param([string]$Path)
        foreach ($homePath in @($script:FieldPref) + @($script:PanelPref) + @($script:NonUi.Keys)) {
            if ($Path -eq $homePath -or $Path.StartsWith($homePath + '.')) { return $true }
        }
        return $false
    }
}

Describe 'Every preference key has a home' {
    It 'has no preference key without a field, a panel or a documented non-UI reason' {
        $missing = @($script:PrefLeaf | Where-Object { -not (Test-Homed -Path $_) })
        $missing | Should -BeNullOrEmpty -Because ('these keys have no home: ' + ($missing -join ', '))
    }

    It 'has no Citrix Workspace App switch without a field or a documented reason' {
        $missing = @($script:CwaLeaf | Where-Object { -not (Test-Homed -Path $_) })
        $missing | Should -BeNullOrEmpty -Because ('these switches have no home: ' + ($missing -join ', '))
    }

    It 'has no TeamViewer Host key without a field' {
        $missing = @($script:TvLeaf | Where-Object { -not (Test-Homed -Path $_) })
        $missing | Should -BeNullOrEmpty -Because ('these keys have no home: ' + ($missing -join ', '))
    }

    It 'has no field that points at a key that does not exist' {
        $known = @($script:PrefLeaf) + @($script:CwaLeaf) + @($script:TvLeaf)
        $unknown = @($script:FieldPref | Where-Object { $p = $_; -not ($known | Where-Object { $_ -eq $p -or $_.StartsWith($p + '.') }) })
        $unknown | Should -BeNullOrEmpty -Because ('these fields point at no key: ' + ($unknown -join ', '))
    }

    It 'has no panel that points at a key that does not exist' {
        $unknown = @($script:PanelPref | Where-Object { $p = $_; -not ($script:PrefLeaf | Where-Object { $_ -eq $p -or $_.StartsWith($p + '.') }) })
        $unknown | Should -BeNullOrEmpty -Because ('these panels point at no key: ' + ($unknown -join ', '))
    }

    It 'has no stale entry in the non-UI list, and no entry that a field also edits' {
        $known = @($script:PrefLeaf) + @($script:CwaLeaf)
        foreach ($key in @($script:NonUi.Keys)) {
            @($known | Where-Object { $_ -eq $key -or $_.StartsWith($key + '.') }).Count | Should -BeGreaterThan 0 -Because "$key is not a key"
            ([string]$script:NonUi[$key]).Length | Should -BeGreaterThan 20 -Because "$key needs a reason"
            @($script:FieldPref | Where-Object { $_ -eq $key }).Count | Should -Be 0 -Because "$key has a field"
        }
    }

    It 'gives each field one preference key' {
        $dup = @($script:FieldPref | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        $dup | Should -BeNullOrEmpty -Because ('keys with two fields: ' + ($dup -join ', '))
    }

    It 'documents every non-UI key in docs\reference\settings.md' {
        $doc = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'docs\reference\settings.md') -Raw
        foreach ($key in @($script:NonUi.Keys)) { $doc | Should -Match ([regex]::Escape(($key -replace '^cwa:', ''))) -Because "settings.md must name $key" }
    }
}

Describe 'The old Options window and the new one write the same keys' {
    BeforeAll {
        # The keys that the old Options panels (ConfigMgr Preferences, Packager Preferences, One Click Settings,
        # Product Filter, Script Signing, WSUS Publishing) and the Setup wizard wrote.
        $script:OldCommitPath = @(
            'Systems.ConfigMgr', 'Systems.Intune', 'Systems.Wsus', 'SiteCode', 'ProviderMachineName', 'DiscoverySqlServer', 'DiscoverySqlDatabase', 'DiscoveryDismissed',
            'FileShareRoot', 'ContentLayout', 'ContentFallback', 'SlowNetworkDeploymentMode', 'DownloadRoot', 'EstimatedRuntimeMins', 'MaximumRuntimeMins',
            'ContentDistribution.AutoDistribute', 'ContentDistribution.DPGroupName', 'ContentDistribution.DeployToTestCollection', 'ContentDistribution.TestCollectionName', 'ContentDistribution.CreateTestCollectionIfMissing',
            'IncludeVersionInTitle', 'RemoveVersionFromTitle', 'Intune.TenantId', 'Intune.ClientId', 'Intune.ClientSecretProtected', 'Vulnerability.NvdApiKeyProtected', 'Vulnerability.LookupAfterCheckCm',
            'AppFlow.Tracked', 'AppFlow.Action', 'AppFlow.CadenceOverrides', 'AppFlow.ForceOnLaunch', 'AppFlow.Destinations', 'AppFlow.DefaultDestinations', 'AppFlow.OnExisting', 'Intune.DeploymentTarget', 'Intune.PublishToIntune',
            'HiddenApplications',
            'CompanyName', 'M365Channel', 'M365DeployMode', 'M365ExcludeApps',
            'SSMSInstallOptions.UIMode', 'SSMSInstallOptions.DownloadThenInstall', 'SSMSInstallOptions.NoUpdateInstaller', 'SSMSInstallOptions.IncludeRecommended', 'SSMSInstallOptions.IncludeOptional', 'SSMSInstallOptions.RemoveOos', 'SSMSInstallOptions.ForceClose', 'SSMSInstallOptions.InstallPath',
            'DBeaverInstallOptions.InstallScope', 'DBeaverInstallOptions.DisableAI', 'AdobeReaderInstallOptions.Edition', 'AdobeReaderInstallOptions.Languages', 'BeyondCompareKeyFile', 'LocalSourceFolders',
            'cwa:Store.Name', 'cwa:Store.Url', 'cwa:Installation.CleanInstall', 'cwa:Installation.IncludeSSON', 'cwa:Installation.EnableSSON', 'cwa:Installation.AppProtection', 'cwa:Installation.SessionPreLaunch', 'cwa:Installation.SelfServiceMode',
            'cwa:Plugins.MSTeamsPlugin', 'cwa:Plugins.ZoomPlugin', 'cwa:Plugins.WebExPlugin', 'cwa:Plugins.UberAgent', 'cwa:Plugins.UberAgentSkipUpgrade', 'cwa:Plugins.EPAClient', 'cwa:Plugins.SessionRecording',
            'cwa:UpdateAndTelemetry.AutoUpdateCheck', 'cwa:UpdateAndTelemetry.EnableCEIP', 'cwa:UpdateAndTelemetry.EnableTracing', 'cwa:StorePolicy.AllowAddStore', 'cwa:StorePolicy.AllowSavePwd',
            'cwa:Components.Customize', 'cwa:Components.ReceiverInside', 'cwa:Components.ICA_Client', 'cwa:Components.AM', 'cwa:Components.SelfService', 'cwa:Components.DesktopViewer', 'cwa:Components.WebHelper', 'cwa:Components.BCR_Client', 'cwa:Components.USB', 'cwa:Components.SSON',
            'tv:ApiToken', 'tv:CustomConfigId', 'tv:AssignmentOptions', 'tv:RemoveDesktopShortcut',
            'ScriptSigning.SignDetection', 'ScriptSigning.SignRequirements', 'ScriptSigning.SignDeployment', 'ScriptSigning.RequireDetection', 'ScriptSigning.RequireRequirements', 'ScriptSigning.RequireDeployment',
            'ScriptSigning.CertificateThumbprint', 'ScriptSigning.StoreLocation', 'ScriptSigning.TimestampServer', 'ScriptSigning.TimestampRequired',
            'Wsus.ServerName', 'Wsus.PortNumber', 'Wsus.UseSsl', 'Wsus.Classification', 'Wsus.ApprovalGroup', 'Wsus.DeclineSuperseded',
            'FirstRunCompleted'
        )
    }

    It 'keeps a home for <_>' -ForEach @(
        'Systems.ConfigMgr', 'SiteCode', 'DiscoverySqlServer', 'DiscoverySqlDatabase', 'DownloadRoot', 'M365ExcludeApps', 'SSMSInstallOptions.InstallPath', 'AdobeReaderInstallOptions.Languages',
        'cwa:Store.Url', 'cwa:Components.USB', 'tv:ApiToken', 'ScriptSigning.TimestampRequired', 'Wsus.PortNumber', 'AppFlow.Tracked', 'HiddenApplications', 'LocalSourceFolders', 'FirstRunCompleted'
    ) {
        (Test-Homed -Path $_) | Should -BeTrue
    }

    It 'gives every key that the old window wrote a home in the new window' {
        $lost = @($script:OldCommitPath | Where-Object { -not (Test-Homed -Path $_) })
        $lost | Should -BeNullOrEmpty -Because ('these settings lost their home: ' + ($lost -join ', '))
    }

    It 'keeps the old key names, and the old keys exist in the defaults' {
        $known = @($script:PrefLeaf) + @($script:CwaLeaf) + @($script:TvLeaf)
        $renamed = @($script:OldCommitPath | Where-Object { $p = $_; -not ($known | Where-Object { $_ -eq $p -or $_.StartsWith($p + '.') }) })
        $renamed | Should -BeNullOrEmpty -Because ('these keys were renamed or removed: ' + ($renamed -join ', '))
    }

    It 'lists the keys that are new, so that a change of the schema is a decision' {
        $new = @($script:PrefLeaf | Where-Object { $p = $_; -not ($script:OldCommitPath | Where-Object { $_ -eq $p -or $p.StartsWith($_ + '.') }) -and $p -notmatch '^(DetectedTools|DeploymentConditions|CommandOverrides)' })
        $expected = @(
            'DiscoveryEnabled', 'DiscoverySqlInstance', 'DiscoverySqlPort', 'WorkbenchDataRoot', 'Vulnerability.CacheTtlMinutes',
            'M365Options.Languages', 'M365Options.OfficeMgmtCom', 'M365Options.UpdatesEnabled', 'M365Options.ForceAppShutdown', 'M365Options.RemoveMsi', 'M365Options.MigrateArch', 'M365Options.SharedComputerLicensing', 'M365Options.DeviceBasedLicensing',
            'M365CustomConfig.Enabled', 'M365CustomConfig.FileName', 'M365CustomConfig.SizeBytes',
            'SSMSInstallOptions.Workloads', 'SSMSInstallOptions.Languages', 'SSMSInstallOptions.NoCache',
            'AdobeReaderInstallOptions.AcceptEula', 'AdobeReaderInstallOptions.DisableDesktopShortcut', 'AdobeReaderInstallOptions.DisableChromeExtension', 'AdobeReaderInstallOptions.DisableBrowserIntegration', 'AdobeReaderInstallOptions.UpdateMode',
            'BeyondCompareOptions.DisableRemoteProfiles', 'BeyondCompareOptions.DisableSavedPasswords',
            'ScriptSigning.HashAlgorithm'
        )
        $unexpected = @($new | Where-Object { $expected -notcontains $_ })
        $unexpected | Should -BeNullOrEmpty -Because ('new keys that the test does not list: ' + ($unexpected -join ', '))
    }
}

Describe 'An older preference file loads unchanged' {
    BeforeEach {
        $script:Older = @'
{
  "SiteCode": "ABC", "ProviderMachineName": "cm01.contoso.com", "FileShareRoot": "\\\\cm01\\sources$", "ContentLayout": "Flat", "DownloadRoot": "D:\\ap",
  "EstimatedRuntimeMins": 20, "MaximumRuntimeMins": 45, "ContentFallback": "Deny", "SlowNetworkDeploymentMode": "DoNothing", "CompanyName": "Contoso",
  "M365Channel": "Current", "M365DeployMode": "Online", "M365ExcludeApps": ["Access", "Bing", "Word"],
  "SSMSInstallOptions": { "UIMode": "Passive", "DownloadThenInstall": false, "NoUpdateInstaller": true, "IncludeRecommended": true, "IncludeOptional": true, "RemoveOos": false, "ForceClose": true, "InstallPath": "D:\\SSMS" },
  "DBeaverInstallOptions": { "InstallScope": "User", "DisableAI": true },
  "AdobeReaderInstallOptions": { "Edition": "MUI", "Languages": ["de_DE"] },
  "BeyondCompareKeyFile": "D:\\keys\\BC5Key.txt",
  "HiddenApplications": ["package-7zip.ps1"], "DiscoverySqlServer": "sql01\\INST,1444", "DiscoverySqlDatabase": "CM_ABC",
  "Systems": { "ConfigMgr": true, "Intune": false, "Wsus": true },
  "Wsus": { "ServerName": "wsus01.contoso.com", "PortNumber": 9000, "UseSsl": false, "Classification": "SecurityUpdates", "ApprovalGroup": "Pilot", "DeclineSuperseded": true },
  "Vulnerability": { "CacheTtlMinutes": 90, "LookupAfterCheckCm": true }
}
'@
        Set-Content -LiteralPath $script:PrefPath -Value $script:Older -Encoding UTF8
    }

    It 'reads every value of the older file' {
        $p = Read-Preferences
        $p.SiteCode | Should -Be 'ABC'
        $p.DownloadRoot | Should -Be 'D:\ap'
        $p.EstimatedRuntimeMins | Should -Be 20
        $p.ContentFallback | Should -Be 'Deny'
        $p.M365Channel | Should -Be 'Current'
        $p.SSMSInstallOptions.InstallPath | Should -Be 'D:\SSMS'
        $p.SSMSInstallOptions.NoUpdateInstaller | Should -BeTrue
        $p.AdobeReaderInstallOptions.Edition | Should -Be 'MUI'
        $p.Wsus.PortNumber | Should -Be 9000
        $p.Vulnerability.CacheTtlMinutes | Should -Be 90
    }

    It 'gives the keys that the older file lacks their defaults' {
        $p = Read-Preferences
        $p.DiscoverySqlInstance | Should -Be ''
        $p.DiscoverySqlPort | Should -Be 0
        $p.M365Options.OfficeMgmtCom | Should -BeTrue
        $p.M365Options.MigrateArch | Should -BeFalse
        $p.M365Options.SharedComputerLicensing | Should -BeFalse
        $p.M365CustomConfig.Enabled | Should -BeFalse
        @($p.SSMSInstallOptions.Workloads).Count | Should -Be 0
        $p.AdobeReaderInstallOptions.AcceptEula | Should -BeFalse
        $p.BeyondCompareOptions.DisableRemoteProfiles | Should -BeFalse
    }

    It 'keeps a server\instance,port value in the old SQL server field as written' {
        (Read-Preferences).DiscoverySqlServer | Should -Be 'sql01\INST,1444'
    }

    It 'drops the retired Bing ExcludeApp ID and keeps the other apps' {
        @((Read-Preferences).M365ExcludeApps) | Should -Be @('Access', 'Word')
    }

    It 'reads a port of the SQL server that is out of range as not set' {
        $edited = $script:Older | ConvertFrom-Json
        $edited | Add-Member -NotePropertyName DiscoverySqlPort -NotePropertyValue 70000
        $edited | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:PrefPath -Encoding UTF8
        (Read-Preferences).DiscoverySqlPort | Should -Be 0
    }
}

Describe 'Save-Preferences mirrors only known keys into packager-preferences.json' {
    It 'writes <Key> from a preference' -ForEach @(
        @{ Key = 'CompanyName' }, @{ Key = 'M365ExcludeApps' }, @{ Key = 'SSMSInstallOptions' }, @{ Key = 'DBeaverInstallOptions' }, @{ Key = 'AdobeReaderInstallOptions' },
        @{ Key = 'BeyondCompareKeyFile' }, @{ Key = 'BeyondCompareOptions' }, @{ Key = 'M365Options' }, @{ Key = 'LocalSourceFolders' }
    ) {
        $text = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Save-Preferences' }, $false)).Extent.Text
        $text | Should -Match ('\$pkgPrefs\["{0}"\]' -f $Key)
        @($script:PrefLeaf | Where-Object { $_ -eq $Key -or $_.StartsWith($Key + '.') }).Count | Should -BeGreaterThan 0
    }

    It 'writes the imported Microsoft 365 file only while it is selected and present' {
        $text = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Save-Preferences' }, $false)).Extent.Text
        $text | Should -Match 'M365UseCustomConfig'
        $text | Should -Match 'Test-Path -LiteralPath \$customPath -PathType Leaf'
        $text | Should -Match 'm365-configuration\.user\.xml'
    }

    It 'mirrors the preferences into the packager file that the packagers read' {
        $root = Join-Path $TestDrive 'app'
        New-Item -ItemType Directory -Path (Join-Path $root 'Packagers') -Force | Out-Null
        Set-Content -LiteralPath $script:PrefPath -Value '{}' -Encoding UTF8
        $prefs = Read-Preferences
        $prefs.M365CustomConfig.Enabled = $true
        Set-Content -LiteralPath (Join-Path $root 'Packagers\m365-configuration.user.xml') -Value '<Configuration />' -Encoding UTF8
        $saveText = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Save-Preferences' }, $false)).Extent.Text.Replace('$PSScriptRoot', ("'" + $root + "'"))
        function Set-DeploymentTypeContentEnvironment { param($Prefs) }
        . ([scriptblock]::Create($saveText))
        Save-Preferences -Prefs $prefs
        $mirror = Get-Content -LiteralPath (Join-Path $root 'Packagers\packager-preferences.json') -Raw | ConvertFrom-Json
        $mirror.M365UseCustomConfig | Should -BeTrue
        $mirror.M365CustomConfigPath | Should -Be (Join-Path $root 'Packagers\m365-configuration.user.xml')
        $mirror.M365Options.OfficeMgmtCom | Should -BeTrue
        $mirror.BeyondCompareOptions.DisableRemoteProfiles | Should -BeFalse
        $prefs.M365CustomConfig.Enabled = $false
        Save-Preferences -Prefs $prefs
        (Get-Content -LiteralPath (Join-Path $root 'Packagers\packager-preferences.json') -Raw | ConvertFrom-Json).M365UseCustomConfig | Should -BeFalse
    }
}
