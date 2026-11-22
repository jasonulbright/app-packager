#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The per-site data layout of the PowerShell app: the site folder, the one-time copy of the single-site files,
    the site switch, and two sites that never mix.

.DESCRIPTION
    The data root is a temp folder (APP_PACKAGER_DATA_ROOT) and LOCALAPPDATA points at a temp folder, so no test
    touches the profile. The site codes are fictional (AB1, CD2) and the hosts use example.com. No window and no ConfigMgr.
    The launcher functions are lifted out of start-apppackager.ps1 by AST, as the Options probes do.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking -Global
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerWsus.psd1') -Force -DisableNameChecking -Global -ErrorAction SilentlyContinue

    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    foreach ($name in 'Read-Preferences', 'Write-PreferenceFile', 'Save-Preferences', 'Enable-AppSite', 'Initialize-AppSiteLayout', 'Switch-AppSite',
        'Get-DeploymentTargetNames', 'Get-KnownDeploymentTarget', 'Resolve-FirstRunCompleted', 'Set-DeploymentTypeContentEnvironment', 'Get-PackagerPreferenceMirror') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "start-apppackager.ps1 no longer defines $name." }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function Test-PipelineBusy { $false }
    function Get-PreferencesPath { $script:PrefsFile }

    $script:Field = $null
    $fieldAst = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] -and @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Id' -and $_.Item2.Extent.Text -eq "'cm.siteCode'" }).Count -eq 1 }, $true)
    if (-not $fieldAst) { throw 'The cm.siteCode field is gone from start-apppackager.ps1.' }
    $script:Field = & ([scriptblock]::Create($fieldAst.Extent.Text))

    function script:New-Fixture {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) ('site-' + [guid]::NewGuid().ToString('N'))
        $fx = [pscustomobject]@{
            Root = $root; Data = (Join-Path $root 'data'); Local = (Join-Path $root 'local'); App = (Join-Path $root 'suite\app-packager')
        }
        New-Item -ItemType Directory -Path (Join-Path $fx.Local 'AppPackager'), (Join-Path $fx.App 'Packagers') -Force | Out-Null
        $script:Saved = @{}
        foreach ($n in 'APP_PACKAGER_DATA_ROOT', 'APP_PACKAGER_SITE', 'SUITE_CM_SITECODE', 'LOCALAPPDATA') { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n) }
        $env:APP_PACKAGER_DATA_ROOT = $fx.Data
        $env:APP_PACKAGER_SITE = $null
        $env:SUITE_CM_SITECODE = $null
        $env:LOCALAPPDATA = $fx.Local
        $fx
    }
    function script:Remove-Fixture {
        param($Fixture)
        foreach ($n in $script:Saved.Keys) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
        Remove-Item -LiteralPath $Fixture.Root -Recurse -Force -ErrorAction SilentlyContinue
    }
    function script:Get-Sha {
        param([string]$Path)
        (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    }
    # An old single-site install: the preferences file with site AB1, a history with ConfigMgr versions, a result file.
    function script:Initialize-LegacyFile {
        param($Fixture)
        $prefs = [ordered]@{
            SiteCode = 'AB1'; ProviderMachineName = 'cm01.example.com'; FileShareRoot = '\\files.example.com\content'
            DiscoverySqlServer = 'sql01.example.com'; DiscoverySqlDatabase = 'CM_AB1'; CompanyName = 'Example Corp'
            ContentDistribution = [ordered]@{ AutoDistribute = $true; DPGroupName = 'All DPs AB1'; TestCollectionName = 'Test AB1' }
            Wsus = [ordered]@{ ServerName = 'wsus01.example.com'; PortNumber = 8530 }
            GitHub = [ordered]@{ TokenProtected = 'dpapi-not-real' }
            DistributionOverrides = [ordered]@{ Apps = [ordered]@{ 'package-example' = [ordered]@{ DPGroup = 'Edge DPs' } } }
        }
        ($prefs | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $Fixture.App 'AppPackager.preferences.json') -Encoding UTF8
        $history = [ordered]@{
            'package-example' = [ordered]@{ LastChecked = '2026-11-01T10:00:00Z'; LastKnownVersion = '2.0'; LastCmChecked = '2026-11-01T10:05:00Z'; LastCmVersion = '1.5'; LastCmSite = 'AB1' }
        }
        ($history | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $Fixture.Local 'AppPackager\app-history.json') -Encoding UTF8
        'Application,ConfigMgrVersion' | Set-Content -LiteralPath (Join-Path $Fixture.Local 'AppPackager\CheckConfigMgrResult.csv') -Encoding UTF8
        '[]' | Set-Content -LiteralPath (Join-Path $Fixture.Local 'AppPackager\CheckConfigMgrResult.json') -Encoding UTF8
    }
}

Describe 'The site of the process' {
    BeforeEach { $script:Fx = New-Fixture }
    AfterEach { Remove-Fixture -Fixture $script:Fx }

    It 'has no site and no site folder without APP_PACKAGER_SITE' {
        Get-AppSiteName | Should -BeNullOrEmpty
        Get-AppSiteFolder | Should -BeNullOrEmpty
        Get-CheckConfigMgrResultFolder | Should -Be (Join-Path $script:Fx.Local 'AppPackager')
    }

    It 'takes a valid site name only' {
        $env:APP_PACKAGER_SITE = 'ab1'
        Get-AppSiteName | Should -Be 'AB1'
        $env:APP_PACKAGER_SITE = 'AB1~2'
        Get-AppSiteName | Should -Be 'AB1~2'
        $env:APP_PACKAGER_SITE = 'SMS'
        Get-AppSiteName | Should -BeNullOrEmpty
        $env:APP_PACKAGER_SITE = '..\x'
        Get-AppSiteName | Should -BeNullOrEmpty
    }

    It 'puts the site folder in the data root, outside the install folder' {
        $env:APP_PACKAGER_SITE = 'AB1'
        $folder = Get-AppSiteFolder
        $folder | Should -BeLike (Join-Path (Join-Path $script:Fx.Data 'Users') '*\sites\AB1')
        $folder | Should -Not -BeLike ($script:Fx.App + '*')
        Get-AppSiteFolder -Area 'results' | Should -Be (Join-Path $folder 'results')
        Get-CheckConfigMgrResultFolder | Should -Be (Join-Path $folder 'results')
    }

    It 'resolves the start site: the explicit value, the environment, active-site.json, then the saved code' {
        Resolve-AppStartupSite -SavedSiteCode 'AB1' | Should -Be 'AB1'
        Set-AppActiveSite -SiteName 'CD2'
        Resolve-AppStartupSite -SavedSiteCode 'AB1' | Should -Be 'CD2'
        $env:SUITE_CM_SITECODE = 'EF3'
        Resolve-AppStartupSite -SavedSiteCode 'AB1' | Should -Be 'EF3'
        Resolve-AppStartupSite -Explicit 'gh4' -SavedSiteCode 'AB1' | Should -Be 'GH4'
        Resolve-AppStartupSite -Explicit 'toolong' -SavedSiteCode 'zz' | Should -Be 'EF3'
    }

    It 'resolves no site when nothing names a valid one' {
        Resolve-AppStartupSite -SavedSiteCode 'MC' | Should -BeNullOrEmpty
        Resolve-AppStartupSite -SavedSiteCode '' | Should -BeNullOrEmpty
    }
}

Describe 'The copy of the single-site files' {
    BeforeEach { $script:Fx = New-Fixture; Initialize-LegacyFile -Fixture $script:Fx }
    AfterEach { Remove-Fixture -Fixture $script:Fx }

    It 'copies the site files into the site folder, verifies each hash and keeps every old file' {
        $old = @{}
        foreach ($f in Get-ChildItem -LiteralPath $script:Fx.Root -Recurse -File) { $old[$f.FullName] = Get-Sha $f.FullName }
        $legacyPrefs = Get-Content -LiteralPath (Join-Path $script:Fx.App 'AppPackager.preferences.json') -Raw | ConvertFrom-Json

        $r = Initialize-AppSiteData -SiteName 'AB1' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local
        $r.Ok | Should -BeTrue
        $r.Report.Outcome | Should -Be 'Migrated'

        $site = Join-Path (Get-SuiteDataRoot) 'sites\AB1'
        (Get-Sha (Join-Path $site 'results\CheckConfigMgrResult.csv')) | Should -Be $old[(Join-Path $script:Fx.Local 'AppPackager\CheckConfigMgrResult.csv')]
        (Get-Sha (Join-Path $site 'results\CheckConfigMgrResult.json')) | Should -Be $old[(Join-Path $script:Fx.Local 'AppPackager\CheckConfigMgrResult.json')]
        $state = Get-Content -LiteralPath (Join-Path $site 'state\app-state.json') -Raw | ConvertFrom-Json
        $state.'package-example'.LastCmVersion | Should -Be '1.5'
        foreach ($path in $old.Keys) { Get-Sha $path | Should -Be $old[$path] }
    }

    It 'writes the site keys, the keys the copy left out, and no secret' {
        $legacyPrefs = Get-Content -LiteralPath (Join-Path $script:Fx.App 'AppPackager.preferences.json') -Raw | ConvertFrom-Json
        [void](Initialize-AppSiteData -SiteName 'AB1' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local)
        $text = Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\AB1\settings.json') -Raw
        $s = $text | ConvertFrom-Json
        $s.SiteCode | Should -Be 'AB1'
        $s.ProviderMachineName | Should -Be 'cm01.example.com'
        $s.FileShareRoot | Should -Be '\\files.example.com\content'
        $s.DiscoverySqlServer | Should -Be 'sql01.example.com'
        $s.ContentDistribution.DPGroupName | Should -Be 'All DPs AB1'
        $s.Wsus.ServerName | Should -Be 'wsus01.example.com'
        $s.DistributionOverrides.Apps.'package-example'.DPGroup | Should -Be 'Edge DPs'
        $text | Should -Not -Match 'dpapi-not-real'
        $text | Should -Not -Match 'CompanyName|GitHub'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\AB1\site.json') -Raw | ConvertFrom-Json).providers | Should -Contain 'cm01.example.com'
    }

    It 'changes nothing on a second run' {
        $legacyPrefs = Get-Content -LiteralPath (Join-Path $script:Fx.App 'AppPackager.preferences.json') -Raw | ConvertFrom-Json
        [void](Initialize-AppSiteData -SiteName 'AB1' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local)
        $site = Join-Path (Get-SuiteDataRoot) 'sites\AB1'
        $before = @{}
        foreach ($f in Get-ChildItem -LiteralPath $site -Recurse -File | Where-Object { $_.Name -notin 'site.json', 'migration.json' }) { $before[$f.FullName] = Get-Sha $f.FullName }
        $r = Initialize-AppSiteData -SiteName 'AB1' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local
        $r.Report.Outcome | Should -Be 'NothingChanged'
        foreach ($p in $before.Keys) { Get-Sha $p | Should -Be $before[$p] -Because $p }
    }

    It 'copies the files for the site code of the old preferences when another site starts' {
        $legacyPrefs = Get-Content -LiteralPath (Join-Path $script:Fx.App 'AppPackager.preferences.json') -Raw | ConvertFrom-Json
        $r = Initialize-AppSiteData -SiteName 'CD2' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local
        $r.Ok | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\AB1\results\CheckConfigMgrResult.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\results') | Should -BeFalse
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\site.json') -Raw | ConvertFrom-Json).siteCode | Should -Be 'CD2'
    }

    It 'refuses an invalid site name and writes nothing' {
        $r = Initialize-AppSiteData -SiteName 'SMS' -AppRoot $script:Fx.App
        $r.Ok | Should -BeFalse
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
    }
}

Describe 'Two sites in one profile' {
    AfterEach { Remove-Fixture -Fixture $script:Fx }

    BeforeEach {
        $script:Fx = New-Fixture
        Initialize-LegacyFile -Fixture $script:Fx
        $legacyPrefs = Get-Content -LiteralPath (Join-Path $script:Fx.App 'AppPackager.preferences.json') -Raw | ConvertFrom-Json
        [void](Initialize-AppSiteData -SiteName 'AB1' -AppRoot $script:Fx.App -MigrateLegacy -LegacySiteCode 'AB1' -LegacyPrefs $legacyPrefs -LocalRoot $script:Fx.Local)
        [void](Initialize-AppSiteData -SiteName 'CD2' -AppRoot $script:Fx.App)
    }

    It 'shows the values of one site and the defaults, not the other site, in the preferences data' {
        $env:APP_PACKAGER_SITE = 'AB1'
        $a = Merge-AppSiteIntoPreferenceData -Data ([pscustomobject]@{ SiteCode = 'XXX'; CompanyName = 'Example Corp' })
        $a.SiteCode | Should -Be 'AB1'
        $a.ProviderMachineName | Should -Be 'cm01.example.com'
        $a.ContentDistribution.DPGroupName | Should -Be 'All DPs AB1'
        $a.CompanyName | Should -Be 'Example Corp'

        $env:APP_PACKAGER_SITE = 'CD2'
        $b = Merge-AppSiteIntoPreferenceData -Data ([pscustomobject]@{ SiteCode = 'AB1'; ProviderMachineName = 'cm01.example.com'; FileShareRoot = 'x'; Wsus = [pscustomobject]@{ ServerName = 'w' }; CompanyName = 'Example Corp' })
        $b.SiteCode | Should -Be 'CD2'
        $b.PSObject.Properties['ProviderMachineName'] | Should -BeNullOrEmpty
        $b.PSObject.Properties['FileShareRoot'] | Should -BeNullOrEmpty
        $b.PSObject.Properties['Wsus'] | Should -BeNullOrEmpty
        $b.CompanyName | Should -Be 'Example Corp'
    }

    It 'saves the site keys into the active site only and keeps the keys that another app wrote' {
        $ab1 = Join-Path (Get-SuiteDataRoot) 'sites\AB1\settings.json'
        $cd2 = Join-Path (Get-SuiteDataRoot) 'sites\CD2\settings.json'
        $extra = Get-Content -LiteralPath $ab1 -Raw | ConvertFrom-Json
        $extra | Add-Member -NotePropertyName 'ProviderMachine' -NotePropertyValue 'apd.example.com' -Force
        ($extra | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $ab1 -Encoding UTF8

        $env:APP_PACKAGER_SITE = 'CD2'
        $prefs = [pscustomobject]@{ SiteCode = 'CD2'; ProviderMachineName = 'cm02.example.com'; FileShareRoot = '\\files2.example.com\content'; CompanyName = 'Example Corp'; GitHub = [pscustomobject]@{ TokenProtected = 'x' } }
        $beforeAb1 = Get-Sha $ab1
        Save-AppSiteSettings -Prefs $prefs | Should -BeTrue
        Get-Sha $ab1 | Should -Be $beforeAb1
        $s = Get-Content -LiteralPath $cd2 -Raw | ConvertFrom-Json
        $s.ProviderMachineName | Should -Be 'cm02.example.com'
        $s.PSObject.Properties['CompanyName'] | Should -BeNullOrEmpty

        $env:APP_PACKAGER_SITE = 'AB1'
        $prefs.ProviderMachineName = 'cm01b.example.com'
        [void](Save-AppSiteSettings -Prefs $prefs)
        $a = Get-Content -LiteralPath $ab1 -Raw | ConvertFrom-Json
        $a.ProviderMachine | Should -Be 'apd.example.com'
        $a.ProviderMachineName | Should -Be 'cm01b.example.com'
        (Get-Content -LiteralPath $cd2 -Raw | ConvertFrom-Json).ProviderMachineName | Should -Be 'cm02.example.com'
    }

    It 'keeps the ConfigMgr state of a packager per site and the rest of the history global' {
        $env:APP_PACKAGER_SITE = 'AB1'
        (Read-PackagerHistory)['package-example'].LastCmVersion | Should -Be '1.5'

        $env:APP_PACKAGER_SITE = 'CD2'
        $h = Read-PackagerHistory
        $h['package-example'].PSObject.Properties['LastCmVersion'] | Should -BeNullOrEmpty
        $h['package-example'].LastKnownVersion | Should -Be '2.0'
        Update-PackagerHistory -PackagerName 'package-example' -Event CmChecked -Version '1.9' -Data @{ SiteCode = 'CD2' }
        (Read-PackagerHistory)['package-example'].LastCmVersion | Should -Be '1.9'

        $env:APP_PACKAGER_SITE = 'AB1'
        (Read-PackagerHistory)['package-example'].LastCmVersion | Should -Be '1.5'
        $state = Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\state\app-state.json') -Raw | ConvertFrom-Json
        $state.'package-example'.LastCmVersion | Should -Be '1.9'
        $global = Get-Content -LiteralPath (Join-Path $script:Fx.Local 'AppPackager\app-history.json') -Raw | ConvertFrom-Json
        $global.'package-example'.LastCmVersion | Should -Be '1.5'
    }

    It 'writes the Check ConfigMgr result into the folder of the active site' {
        $rows = @([pscustomobject]@{ Application = 'Example'; Vendor = 'Vendor'; CMName = 'Example'; ArpName = ''; ArpArch = ''; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available' })
        $env:APP_PACKAGER_SITE = 'CD2'
        $legacy = Join-Path $script:Fx.Local 'AppPackager\CheckConfigMgrResult.csv'
        $ab1 = Join-Path (Get-SuiteDataRoot) 'sites\AB1\results\CheckConfigMgrResult.csv'
        $legacyBefore = Get-Sha $legacy
        $ab1Before = Get-Sha $ab1
        Write-CheckConfigMgrResultDefault -Result @(ConvertTo-CheckConfigMgrResultRow -Rows $rows)
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\results\CheckConfigMgrResult.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\results\CheckConfigMgrResult.json') | Should -BeTrue
        Get-Sha $legacy | Should -Be $legacyBefore
        Get-Sha $ab1 | Should -Be $ab1Before
    }

    It 'lists the known sites by folder name' {
        @(Get-AppSiteList | ForEach-Object { $_.Name }) | Should -Be @('AB1', 'CD2')
        (Get-AppSiteList)[0].Provider | Should -Be 'cm01.example.com'
    }

    It 'writes active-site.json and leaves it when the site did not change' {
        Set-AppActiveSite -SiteName 'CD2'
        $file = Join-Path (Get-SuiteDataRoot) 'active-site.json'
        $j = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
        $j.schemaVersion | Should -Be 1
        $j.siteCode | Should -Be 'CD2'
        $before = Get-Sha $file
        Set-AppActiveSite -SiteName 'CD2'
        Get-Sha $file | Should -Be $before
        { Set-AppActiveSite -SiteName 'SMS' } | Should -Throw
    }
}

Describe 'Start, save and switch in the launcher' {
    BeforeEach {
        $script:Fx = New-Fixture
        $script:PrefsFile = Join-Path $script:Fx.App 'AppPackager.preferences.json'
        Mock -CommandName Write-Log -MockWith { }
        $script:SiteRequest = $null
    }
    AfterEach { Remove-Fixture -Fixture $script:Fx }

    It 'starts in the site of the saved preferences, copies the files once and records the copy' {
        Initialize-LegacyFile -Fixture $script:Fx
        $localBefore = Get-Sha (Join-Path $script:Fx.Local 'AppPackager\app-history.json')
        Mock -CommandName Enable-AppSite -MockWith { [pscustomobject]@{ Ok = $true; MigratedFrom = 'AB1'; Text = '' } }
        Initialize-AppSiteLayout -ExplicitSite ''
        $env:APP_PACKAGER_SITE | Should -Be 'AB1'
        $script:PendingSiteMigratedFrom | Should -Be 'AB1'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'active-site.json') -Raw | ConvertFrom-Json).siteCode | Should -Be 'AB1'
        Get-Sha (Join-Path $script:Fx.Local 'AppPackager\app-history.json') | Should -Be $localBefore
    }

    It 'keeps the single-site layout when the copy fails' {
        Initialize-LegacyFile -Fixture $script:Fx
        Mock -CommandName Enable-AppSite -MockWith { [pscustomobject]@{ Ok = $false; MigratedFrom = ''; Text = 'The copy did not match its source.' } }
        Mock -CommandName Write-Warning -MockWith { }
        Initialize-AppSiteLayout -ExplicitSite ''
        Get-AppSiteName | Should -BeNullOrEmpty
        Should -Invoke Write-Warning -Times 1
    }

    It 'starts with no site on a fresh install, even when another tool left an active site' {
        Set-AppActiveSite -SiteName 'CD2'
        Initialize-AppSiteLayout -ExplicitSite ''
        Get-AppSiteName | Should -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
    }

    It 'does not start a site before the Setup wizard is done' {
        $prefs = Read-Preferences
        $prefs.SiteCode = 'AB1'
        Save-Preferences -Prefs $prefs
        Get-AppSiteName | Should -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
    }

    It 'starts the first site when the Setup wizard is done' {
        $prefs = Read-Preferences
        $prefs.SiteCode = 'AB1'
        $prefs.ProviderMachineName = 'cm01.example.com'
        $prefs.FirstRunCompleted = $true
        Save-Preferences -Prefs $prefs
        Get-AppSiteName | Should -Be 'AB1'
        $prefs.SiteDataMigratedFrom | Should -Be 'AB1'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\AB1\settings.json') -Raw | ConvertFrom-Json).ProviderMachineName | Should -Be 'cm01.example.com'
        (Read-Preferences).ProviderMachineName | Should -Be 'cm01.example.com'
    }

    It 'saves the changes to the active site, then switches to a new site with empty site values' {
        Initialize-LegacyFile -Fixture $script:Fx
        $env:APP_PACKAGER_SITE = $null
        Initialize-AppSiteLayout -ExplicitSite ''
        $prefs = Read-Preferences
        $prefs.SiteDataMigratedFrom = $script:PendingSiteMigratedFrom
        $prefs.ProviderMachineName | Should -Be 'cm01.example.com'

        $prefs.ProviderMachineName = 'cm01-edited.example.com'
        $prefs.SiteCode = 'CD2'
        $script:SiteRequest = 'CD2'
        Save-Preferences -Prefs $prefs

        Get-AppSiteName | Should -Be 'CD2'
        $prefs.SiteCode | Should -Be 'CD2'
        $prefs.ProviderMachineName | Should -BeNullOrEmpty
        $prefs.ContentDistribution.DPGroupName | Should -BeNullOrEmpty
        $prefs.CompanyName | Should -Be 'Example Corp'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\AB1\settings.json') -Raw | ConvertFrom-Json).ProviderMachineName | Should -Be 'cm01-edited.example.com'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'active-site.json') -Raw | ConvertFrom-Json).siteCode | Should -Be 'CD2'
        (Get-Content -LiteralPath $script:PrefsFile -Raw | ConvertFrom-Json).SiteCode | Should -Be 'CD2'
        $script:LastCheckConfigMgrResult | Should -BeNullOrEmpty

        $prefs.ProviderMachineName = 'cm02.example.com'
        $prefs.SiteCode = 'AB1'
        $script:SiteRequest = 'AB1'
        Save-Preferences -Prefs $prefs
        Get-AppSiteName | Should -Be 'AB1'
        $prefs.ProviderMachineName | Should -Be 'cm01-edited.example.com'
        $prefs.ContentDistribution.DPGroupName | Should -Be 'All DPs AB1'
        (Get-Content -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites\CD2\settings.json') -Raw | ConvertFrom-Json).ProviderMachineName | Should -Be 'cm02.example.com'
    }

    It 'does not switch while a run is active' {
        Initialize-LegacyFile -Fixture $script:Fx
        Initialize-AppSiteLayout -ExplicitSite ''
        $prefs = Read-Preferences
        $prefs.SiteDataMigratedFrom = $script:PendingSiteMigratedFrom
        Mock -CommandName Test-PipelineBusy -MockWith { $true }
        $script:SiteRequest = 'CD2'
        { Save-Preferences -Prefs $prefs } | Should -Throw '*run is active*'
        Get-AppSiteName | Should -Be 'AB1'
    }

    It 'keeps the old files as they were through the start and the first save' {
        Initialize-LegacyFile -Fixture $script:Fx
        $old = @{}
        foreach ($f in Get-ChildItem -LiteralPath $script:Fx.Local -Recurse -File) { $old[$f.FullName] = Get-Sha $f.FullName }
        Initialize-AppSiteLayout -ExplicitSite ''
        $prefs = Read-Preferences
        $prefs.SiteDataMigratedFrom = $script:PendingSiteMigratedFrom
        Save-Preferences -Prefs $prefs
        Update-PackagerHistory -PackagerName 'package-example' -Event CmChecked -Version '1.7' -Data @{ SiteCode = 'AB1' }
        foreach ($path in @($old.Keys | Where-Object { $_ -notlike '*app-history.json' })) { Get-Sha $path | Should -Be $old[$path] }
        (Get-Content -LiteralPath (Join-Path $script:Fx.Local 'AppPackager\app-history.json') -Raw | ConvertFrom-Json).'package-example'.LastCmVersion | Should -Be '1.5'
    }
}

Describe 'The site code field of Options' {
    It 'is one editable list in the ConfigMgr section' {
        $script:Field.Kind | Should -Be 'ComboText'
        $script:Field.Id | Should -Be 'cm.siteCode'
        $script:Field.Title | Should -Be 'Site code'
        $script:Field.Pref | Should -Be 'SiteCode'
        $script:Field.Tip | Should -Not -BeNullOrEmpty
    }

    It 'accepts a site code and a reused code, and refuses a reserved or malformed one' {
        & $script:Field.Validate '' | Should -BeNullOrEmpty
        & $script:Field.Validate 'AB1' | Should -BeNullOrEmpty
        & $script:Field.Validate 'ab1~2' | Should -BeNullOrEmpty
        & $script:Field.Validate 'SMS' | Should -Not -BeNullOrEmpty
        & $script:Field.Validate 'AB' | Should -Not -BeNullOrEmpty
        & $script:Field.Validate 'AB1~1' | Should -Not -BeNullOrEmpty
        & $script:Field.Validate 'AB 1' | Should -Not -BeNullOrEmpty
    }

    It 'stores the three-character code and keeps the site name for the switch' {
        $script:SiteRequest = $null
        & $script:Field.FromText 'ab1~2' | Should -Be 'AB1'
        $script:SiteRequest | Should -Be 'AB1~2'
    }

    It 'shows the folder name of the active site' {
        $fx = New-Fixture
        try {
            & $script:Field.ToText 'AB1' | Should -Be 'AB1'
            $env:APP_PACKAGER_SITE = 'AB1~2'
            & $script:Field.ToText 'AB1' | Should -Be 'AB1~2'
            & $script:Field.ToText 'CD2' | Should -Be 'CD2'
        }
        finally { Remove-Fixture -Fixture $fx }
    }

    It 'lists the known sites' {
        $fx = New-Fixture
        try {
            [void](Initialize-AppSiteData -SiteName 'AB1' -AppRoot $fx.App)
            [void](Initialize-AppSiteData -SiteName 'CD2' -AppRoot $fx.App)
            @(& $script:Field.ItemsFrom) | Should -Be @('AB1', 'CD2')
        }
        finally { Remove-Fixture -Fixture $fx }
    }
}
