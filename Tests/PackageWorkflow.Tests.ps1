BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in @('Invoke-PackagerPackageWithConflictPrompt', 'Get-TitleModesMapForContext')) {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function Invoke-PackagerPackage { param($Preflight, $OnExisting, $DeploymentTarget, $TitleMode, $SiteCode, $ProviderMachineName) }
    function Get-MecmCurrentVersionByCMName { param($SiteCode, $ProviderMachineName, $CMName, [switch]$ExactMatch) }
}

Describe 'Application title policy' {
    It 'resolves <Name> as <Expected>' -TestCases @(
        @{ Name = 'Google Chrome 130.0.1'; Version = '130.0.1'; Mode = 'NoVersion'; Expected = 'Google Chrome' }
        @{ Name = 'Google Chrome 131.0.2'; Version = '131.0.2'; Mode = 'NoVersion'; Expected = 'Google Chrome' }
        @{ Name = 'Mozilla Firefox (x64 en-US)'; Version = '130.0'; Mode = 'IncludeVersion'; Expected = 'Mozilla Firefox (x64 en-US) - 130.0' }
        @{ Name = 'Microsoft Edge - 130.0'; Version = '130.0'; Mode = 'IncludeVersion'; Expected = 'Microsoft Edge - 130.0' }
        @{ Name = 'M365 Apps - 16.0.1 (x64) (Current)'; Version = '16.0.1'; Mode = 'NoVersion'; Expected = 'M365 Apps (x64) (Current)' }
        @{ Name = 'SQL Server 2022 - 16.0.1'; Version = '16.0.1'; Mode = 'NoVersion'; Expected = 'SQL Server 2022' }
        @{ Name = 'App - 1.0'; Version = '1.0'; Mode = 'Default'; Expected = 'App - 1.0' }
        @{ Name = 'App'; Version = '1.0'; Mode = ''; Expected = 'App' }
    ) {
        param($Name, $Version, $Mode, $Expected)
        Get-PackagedApplicationName -AppName $Name -Version $Version -Mode $Mode | Should -Be $Expected
    }

    It 'round-trips per-app preferences into the background context without changing defaults' {
        $script:Prefs = @{ DeploymentConditions = @{ Apps = ([pscustomobject]@{
            'package-chrome' = [pscustomobject]@{ TitleMode = 'NoVersion' }
            'package-firefox' = [pscustomobject]@{ TitleMode = 'IncludeVersion' }
            'package-edge' = [pscustomobject]@{ Architecture = 'Any' }
        } | ConvertTo-Json -Depth 5 | ConvertFrom-Json) } }
        $map = Get-TitleModesMapForContext
        $map['package-chrome'] | Should -Be 'NoVersion'
        $map['package-firefox'] | Should -Be 'IncludeVersion'
        $map.ContainsKey('package-edge') | Should -BeFalse
    }

    It 'ends an isolated probe at manifest read without running subsequent package writes' {
        $manifestPath = Join-Path $TestDrive 'stage-manifest.json'
        @{ SchemaVersion = 2; AppName = "Contoso's Browser - 2.0"; SoftwareVersion = '2.0' } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding ASCII
        $modulePath = (Resolve-Path "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1").Path
        $probe = @"
Import-Module '$modulePath' -Force
`$env:APP_PACKAGER_TITLE_MODE = 'NoVersion'
`$env:APP_PACKAGER_PACKAGE_PREFLIGHT = '1'
`$capturedManifest = Read-StageManifest -Path '$manifestPath'
throw 'Continued past preflight'
"@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe))
        $output = & powershell.exe -NoProfile -OutputFormat Text -EncodedCommand $encoded
        $LASTEXITCODE | Should -Be 0
        $line = @($output | Where-Object { $_ -like '[[]APP_PACKAGER_PREFLIGHT]*' })
        $line.Count | Should -Be 1
        $identity = $line[0].Substring('[APP_PACKAGER_PREFLIGHT] '.Length) | ConvertFrom-Json
        $identity.AppName | Should -Be "Contoso's Browser"
        (Get-Content $manifestPath -Raw | ConvertFrom-Json).AppName | Should -Be "Contoso's Browser - 2.0"
    }
}

Describe 'Package conflict preflight' {
    BeforeEach {
        $script:state = @{ ConflictDecisionForAll = 'Skip'; CancelRequested = $false; LogQueue = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]') }
        $script:packageTestArgs = @{ SiteCode = 'MCM'; ProviderMachineName = 'cm.example'; DeploymentTarget = 'MECM'; TitleMode = 'NoVersion' }
        Mock Invoke-PackagerPackage {
            if ($Preflight) { return [pscustomobject]@{ ExitCode = 0; StdOut = '[APP_PACKAGER_PREFLIGHT] {"AppName":"Browser","Version":"2.0"}' } }
            return [pscustomobject]@{ ExitCode = 0; StdOut = 'Packaged' }
        }
        Mock Get-MecmCurrentVersionByCMName { [pscustomobject]@{ Found = $true; SoftwareVersion = '1.0' } }
    }

    It 'skips a different-version collision without running the real package phase' {
        $r = Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs
        $r.PackageOutcome | Should -Be 'Skipped'
        Should -Invoke Invoke-PackagerPackage -Times 0 -Exactly -ParameterFilter { -not $Preflight }
        Should -Invoke Get-MecmCurrentVersionByCMName -Times 1 -Exactly -ParameterFilter { $ExactMatch -and $CMName -eq 'Browser' }
    }

    It 'overwrites a different-version collision only after approval' {
        $script:state.ConflictDecisionForAll = 'Overwrite'
        Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs | Out-Null
        Should -Invoke Invoke-PackagerPackage -Times 1 -Exactly -ParameterFilter { -not $Preflight -and $OnExisting -eq 'Overwrite' -and $TitleMode -eq 'NoVersion' }
    }

    It 'cancels the remaining run without packaging' {
        $script:state.ConflictDecisionForAll = 'Cancel'
        $r = Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs
        $r.PackageOutcome | Should -Be 'Canceled'
        $script:state.CancelRequested | Should -BeTrue
        Should -Invoke Invoke-PackagerPackage -Times 0 -Exactly -ParameterFilter { -not $Preflight }
    }

    It 'creates a new application and fails closed if a collision appears after preflight' {
        Mock Get-MecmCurrentVersionByCMName { [pscustomobject]@{ Found = $false } }
        Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs | Out-Null
        Should -Invoke Invoke-PackagerPackage -Times 1 -Exactly -ParameterFilter { -not $Preflight -and $OnExisting -eq 'Fail' }
    }

    It 'fails closed if no authoritative manifest identity is returned' {
        Mock Invoke-PackagerPackage { [pscustomobject]@{ ExitCode = 0; StdOut = 'no identity' } }
        { Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs } | Should -Throw '*refusing to package*'
        Should -Invoke Get-MecmCurrentVersionByCMName -Times 0 -Exactly
    }

    It 'does not interpret a site-query failure as an absent application' {
        Mock Get-MecmCurrentVersionByCMName { throw 'Site unavailable' }
        { Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs } | Should -Throw '*Site unavailable*'
        Should -Invoke Invoke-PackagerPackage -Times 0 -Exactly -ParameterFilter { -not $Preflight }
    }

    It 'leaves Intune-only packaging on its existing path' {
        $script:packageTestArgs.DeploymentTarget = 'IntuneOnly'
        Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs | Out-Null
        Should -Invoke Get-MecmCurrentVersionByCMName -Times 0 -Exactly
        Should -Invoke Invoke-PackagerPackage -Times 1 -Exactly -ParameterFilter { -not $Preflight }
    }

    It 'honors cancellation during preflight even when the title is new' {
        Mock Get-MecmCurrentVersionByCMName { $script:state.CancelRequested = $true; [pscustomobject]@{ Found = $false } }
        $r = Invoke-PackagerPackageWithConflictPrompt -State $script:state -AppLabel Browser -PackageArgs $script:packageTestArgs
        $r.PackageOutcome | Should -Be 'Canceled'
        Should -Invoke Invoke-PackagerPackage -Times 0 -Exactly -ParameterFilter { -not $Preflight }
    }
}

# ---------------------------------------------------------------------------
# Level-3 shared mechanisms: profile precedence, migration, stage isolation,
# build selection, run overrides, and CLI/GUI plan equivalence.
# ---------------------------------------------------------------------------

Describe 'Workbench profile precedence' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerWorkbench.psd1" -Force
        $script:GlobalDefaults = [pscustomobject]@{ Description = 'global text'; EstimatedRuntimeMins = 15; MaximumRuntimeMins = 30; TitleMode = 'IncludeVersion' }
        $script:BaseManifest = [pscustomobject]@{
            AppName = 'Contoso Reader'; Publisher = 'Contoso'; Description = 'packager text'
            SoftwareVersion = '3.1'; InstallCommandLine = 'install.bat'
            InstallationBehaviorType = 'InstallForSystem'
            Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso' }
        }
    }

    It 'takes the packager value over the global default' {
        $effective = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $script:BaseManifest -Profile $null
        $effective['Description'].Value | Should -Be 'packager text'
        $effective['Description'].Source | Should -Be 'Packager'
    }

    It 'takes the profile value over the packager value' {
        $profile = [pscustomobject]@{ Application = [pscustomobject]@{ DisplayName = 'Reader (Managed)' } }
        $effective = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $script:BaseManifest -Profile $profile
        $effective['DisplayName'].Value | Should -Be 'Reader (Managed)'
        $effective['DisplayName'].Source | Should -Be 'Profile'
    }

    It 'distinguishes an explicit null (none) from an absent key (inherit)' {
        $explicitNone = [pscustomobject]@{ Application = [pscustomobject]@{ Description = $null } }
        $inherit = [pscustomobject]@{ Application = [pscustomobject]@{ Publisher = 'Contoso Ltd' } }

        $none = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $script:BaseManifest -Profile $explicitNone
        $none['Description'].Value | Should -BeNullOrEmpty
        $none['Description'].Source | Should -Be 'Profile'

        $inherited = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $script:BaseManifest -Profile $inherit
        $inherited['Description'].Value | Should -Be 'packager text'
        $inherited['Description'].Source | Should -Be 'Packager'
    }

    It 'treats an explicit empty string as an override, not as inherit' {
        $profile = [pscustomobject]@{ Application = [pscustomobject]@{ Description = '' } }
        $effective = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $script:BaseManifest -Profile $profile
        $effective['Description'].Value | Should -Be ''
        $effective['Description'].Source | Should -Be 'Profile'
    }

    It 'lets a variant override outrank the profile and a run override outrank the variant' {
        $profile = [pscustomobject]@{
            Install = [pscustomobject]@{ Command = 'profile.bat' }
            Variants = [pscustomobject]@{ Overrides = [pscustomobject]@{ x64 = [pscustomobject]@{ InstallCommand = 'variant.bat' } } }
        }
        $variant = Resolve-EffectiveSettings -BaseManifest $script:BaseManifest -Profile $profile -Variant 'x64'
        $variant['InstallCommand'].Value | Should -Be 'variant.bat'
        $variant['InstallCommand'].Source | Should -Be 'Variant'

        $run = Resolve-EffectiveSettings -BaseManifest $script:BaseManifest -Profile $profile -Variant 'x64' -RunOverrides ([pscustomobject]@{ InstallCommand = 'run.bat' })
        $run['InstallCommand'].Value | Should -Be 'run.bat'
        $run['InstallCommand'].Source | Should -Be 'Run'
    }

    It 'reports a field only a completed stage can resolve as NeedsStaging' {
        $effective = Resolve-EffectiveSettings -GlobalDefaults $script:GlobalDefaults -BaseManifest $null -Profile $null
        $effective['Icon'].Source | Should -Be 'NeedsStaging'
        $effective['TitleMode'].Source | Should -Be 'Global'
    }
}

Describe 'Legacy preference migration' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerWorkbench.psd1" -Force
    }
    BeforeEach {
        $script:DataRoot = Join-Path $TestDrive ('wb-migrate-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        # Round-tripped through JSON so the fixture has the PSCustomObject
        # shape the preferences file produces on load.
        $script:LegacyPrefs = @{
            DeploymentConditions = @{
                Apps = @{
                    'package-7zip' = @{ Architecture = 'x64'; Languages = @('en-US', 'de-DE'); Network = 'VpnOnly'; Split = 'Architecture'; InstallMode = 'CurrentUser'; TitleMode = 'NoVersion' }
                    'package-git'  = @{ Architecture = 'Any' }
                }
            }
            CommandOverrides = @{
                Apps = @{
                    'package-7zip'      = @{ Install = 'setup.exe /S /custom'; Uninstall = 'uninst.exe /S' }
                    'package-notepadpp' = @{ Install = 'npp.exe /S' }
                }
            }
            InstallModes = @{ 'package-vlc' = 'AllUsers' }
            TitleModes   = @{ 'package-vlc' = 'NoVersion' }
        } | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    }

    It 'creates one active Migrated profile per legacy application' {
        $result = Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot
        $result.MigratedCount | Should -Be 3
        foreach ($key in 'package-7zip', 'package-git', 'package-notepadpp') {
            $definition = Get-ApplicationDefinition -ApplicationId "catalog:$key" -DataRoot $script:DataRoot
            $definition.ActiveProfileId | Should -Not -Be 'default'
            (Get-Profile -ApplicationId "catalog:$key" -ProfileId $definition.ActiveProfileId -DataRoot $script:DataRoot).Name | Should -Be 'Migrated'
        }
    }

    It 'maps conditions, split, install mode, title mode and commands onto the profile' {
        [void](Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot)
        $definition = Get-ApplicationDefinition -ApplicationId 'catalog:package-7zip' -DataRoot $script:DataRoot
        $profile = Get-Profile -ApplicationId 'catalog:package-7zip' -ProfileId $definition.ActiveProfileId -DataRoot $script:DataRoot

        $profile.MigratedFrom | Should -Be 'package-7zip'
        @($profile.Requirements.Operations | ForEach-Object { $_.RuleId }) | Should -Be @('cpu-arch', 'os-language', 'vpn-connected')
        ($profile.Requirements.Operations | Where-Object { $_.RuleId -eq 'vpn-connected' }).Rule.Value | Should -BeTrue
        $profile.Variants.Split.Split | Should -Be 'Architecture'
        $profile.InstallMode | Should -Be 'CurrentUser'
        $profile.Application.TitleMode | Should -Be 'NoVersion'
        $profile.Install.Command | Should -Be 'setup.exe /S /custom'
        $profile.Uninstall.Command | Should -Be 'uninst.exe /S'
    }

    It 'produces effective settings that carry the migrated overrides' {
        [void](Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot)
        $definition = Get-ApplicationDefinition -ApplicationId 'catalog:package-7zip' -DataRoot $script:DataRoot
        $profile = Get-Profile -ApplicationId 'catalog:package-7zip' -ProfileId $definition.ActiveProfileId -DataRoot $script:DataRoot

        $base = [pscustomobject]@{ AppName = '7-Zip'; InstallCommandLine = 'install.bat'; InstallMode = 'AllUsers' }
        $effective = Resolve-EffectiveSettings -BaseManifest $base -Profile $profile
        $effective['InstallCommand'].Value | Should -Be 'setup.exe /S /custom'
        $effective['InstallCommand'].Source | Should -Be 'Profile'
        $effective['InstallMode'].Value | Should -Be 'CurrentUser'
        $effective['TitleMode'].Value | Should -Be 'NoVersion'
    }

    It 'skips an application whose legacy content has not changed' {
        [void](Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot)
        $second = Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot
        $second.MigratedCount | Should -Be 0
        $second.SkippedCount | Should -Be 3
    }

    It 'leaves the legacy preference keys in place for the environment bridge' {
        [void](Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot)
        $script:LegacyPrefs.DeploymentConditions.Apps.'package-7zip'.TitleMode | Should -Be 'NoVersion'
        $script:LegacyPrefs.CommandOverrides.Apps.'package-7zip'.Install | Should -Be 'setup.exe /S /custom'
    }

    It 'migrates nothing for an application that carries no legacy per-app entry' {
        # InstallModes and TitleModes are not top-level maps in the shipped
        # preferences shape; both live inside DeploymentConditions.Apps.
        [void](Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPrefs -DataRoot $script:DataRoot)
        $profiles = @(Get-Profiles -ApplicationId 'catalog:package-vlc' -DataRoot $script:DataRoot)
        $profiles.Count | Should -Be 1
        $profiles[0].IsDefault | Should -BeTrue
    }
}

Describe 'Per-profile stage isolation and build selection' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
        $script:IsoRoot = Join-Path $TestDrive 'isolation'
        $script:DataRoot = Join-Path $script:IsoRoot 'data'
        $script:AppId = 'catalog:package-gfixture'
        $script:PreviousSnapshot = $env:APP_PACKAGER_RUN_SNAPSHOT
        # Write-BuildRecord resolves its own data root, which is why the GUI
        # and the CLI both put APP_PACKAGER_WORKBENCH_ROOT on the child.
        $script:PreviousRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
        $env:APP_PACKAGER_WORKBENCH_ROOT = $script:DataRoot

        function New-IsolationProfile {
            param([string]$Name, [string]$DisplayName)
            $profile = New-WorkbenchProfileObject -ApplicationId $script:AppId -Name $Name
            $profile['Application'] = @{ DisplayName = $DisplayName }
            return (Save-Profile -Profile ([pscustomobject]$profile) -DataRoot $script:DataRoot)
        }

        function Invoke-IsolationStage {
            param([string]$ProfileId, [string]$StageRoot, [hashtable]$RunOverrides)
            New-Item -ItemType Directory -Path $StageRoot -Force | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $StageRoot 'setup.exe'), "payload for $ProfileId")
            $snapshot = New-RunSnapshot -ApplicationId $script:AppId -ProfileId $ProfileId -Target 'ContentOnly' `
                -DataRoot $script:DataRoot -DownloadRoot $StageRoot -RunOverrides $RunOverrides `
                -PackagerScriptPath (Join-Path $script:IsoRoot 'package-gfixture.ps1')
            $env:APP_PACKAGER_RUN_SNAPSHOT = $snapshot.Path
            try {
                $manifest = @{
                    AppName         = 'G Fixture'
                    Publisher       = 'Contoso'
                    SoftwareVersion = '1.0.0'
                    InstallerFile   = 'setup.exe'
                    InstallerType   = 'EXE'
                    Detection       = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\GFixture'; ValueName = 'DisplayVersion'; ExpectedValue = '1.0.0'; Operator = 'GreaterEquals'; Is64Bit = $true }
                }
                Write-StageManifest -Path (Join-Path $StageRoot 'stage-manifest.json') -ManifestData $manifest -PackagerScriptPath (Join-Path $script:IsoRoot 'package-gfixture.ps1') 6>$null | Out-Null
                return $manifest
            }
            finally { $env:APP_PACKAGER_RUN_SNAPSHOT = $script:PreviousSnapshot }
        }

        $script:ProfileAlpha = New-IsolationProfile -Name 'Alpha' -DisplayName 'G Fixture (Alpha)'
        $script:ProfileBeta = New-IsolationProfile -Name 'Beta' -DisplayName 'G Fixture (Beta)'
        $script:StageAlpha = Join-Path $script:IsoRoot 'stage\profiles\alpha'
        $script:StageBeta = Join-Path $script:IsoRoot 'stage\profiles\beta'
        $script:ManifestAlpha = Invoke-IsolationStage -ProfileId $script:ProfileAlpha.ProfileId -StageRoot $script:StageAlpha -RunOverrides @{ EstimatedMinutes = 7; MaximumMinutes = 21 }
        $script:ManifestBeta = Invoke-IsolationStage -ProfileId $script:ProfileBeta.ProfileId -StageRoot $script:StageBeta -RunOverrides @{}
    }
    AfterAll {
        $env:APP_PACKAGER_RUN_SNAPSHOT = $script:PreviousSnapshot
        $env:APP_PACKAGER_WORKBENCH_ROOT = $script:PreviousRoot
    }

    It 'gives each profile its own BuildId' {
        $script:ManifestAlpha.BuildId | Should -Not -BeNullOrEmpty
        $script:ManifestBeta.BuildId | Should -Not -BeNullOrEmpty
        $script:ManifestAlpha.BuildId | Should -Not -Be $script:ManifestBeta.BuildId
    }

    It 'stamps schema 4 and a plan digest on both builds' {
        foreach ($manifest in $script:ManifestAlpha, $script:ManifestBeta) {
            $manifest.SchemaVersion | Should -Be 4
            $manifest.PlanDigest | Should -Match '^[0-9A-F]{64}$'
        }
        $script:ManifestAlpha.PlanDigest | Should -Not -Be $script:ManifestBeta.PlanDigest
    }

    It 'shares no staged file between the two profile roots' {
        $alpha = @(Get-ChildItem -LiteralPath $script:StageAlpha -Recurse -File | ForEach-Object { $_.FullName })
        $beta = @(Get-ChildItem -LiteralPath $script:StageBeta -Recurse -File | ForEach-Object { $_.FullName })
        $alpha.Count | Should -BeGreaterThan 0
        @($alpha | Where-Object { $beta -contains $_ }).Count | Should -Be 0
        (Get-Content -LiteralPath (Join-Path $script:StageAlpha 'setup.exe') -Raw) | Should -Not -Be (Get-Content -LiteralPath (Join-Path $script:StageBeta 'setup.exe') -Raw)
    }

    It 'seals one build record per profile' {
        foreach ($entry in @(@{ Id = $script:ProfileAlpha.ProfileId; Build = $script:ManifestAlpha.BuildId }, @{ Id = $script:ProfileBeta.ProfileId; Build = $script:ManifestBeta.BuildId })) {
            $records = @(Get-BuildRecords -ApplicationId $script:AppId -ProfileId $entry.Id -DataRoot $script:DataRoot)
            $records.Count | Should -Be 1
            $records[0].BuildId | Should -Be $entry.Build
        }
    }

    It 'resolves the manifest for an exact BuildId across both roots' {
        $searchRoot = Join-Path $script:IsoRoot 'stage'
        $resolved = Resolve-StageManifestForBuild -BuildId $script:ManifestAlpha.BuildId -SearchRoot $searchRoot
        $resolved.StageRoot | Should -Be $script:StageAlpha
        $resolved.Manifest.BuildId | Should -Be $script:ManifestAlpha.BuildId
    }

    It 'refuses a stale build instead of packaging the newest manifest' {
        $searchRoot = Join-Path $script:IsoRoot 'stage'
        { Resolve-StageManifestForBuild -BuildId '20200101-000000-deadbeef' -SearchRoot $searchRoot } | Should -Throw '*stale build*'
    }

    It 'records run overrides in build.json and never writes them back to the profile' {
        $record = Get-LatestBuildRecord -ApplicationId $script:AppId -ProfileId $script:ProfileAlpha.ProfileId -DataRoot $script:DataRoot
        [int]$record.RunOverrides.EstimatedMinutes | Should -Be 7
        [int]$record.RunOverrides.MaximumMinutes | Should -Be 21

        $stored = Get-Profile -ApplicationId $script:AppId -ProfileId $script:ProfileAlpha.ProfileId -DataRoot $script:DataRoot
        $stored.Timing.EstimatedMinutes | Should -BeNullOrEmpty
        $stored.Timing.MaximumMinutes | Should -BeNullOrEmpty
        $stored.Revision | Should -Be 1
    }

    It 'refuses a run override that is not a positive whole number' {
        { New-RunSnapshot -ApplicationId $script:AppId -ProfileId $script:ProfileAlpha.ProfileId -DataRoot $script:DataRoot -RunOverrides @{ EstimatedMinutes = 0 } } |
            Should -Throw '*positive whole number*'
    }

    It 'refuses a run override whose estimate exceeds the maximum' {
        { New-RunSnapshot -ApplicationId $script:AppId -ProfileId $script:ProfileAlpha.ProfileId -DataRoot $script:DataRoot -RunOverrides @{ EstimatedMinutes = 30; MaximumMinutes = 10 } } |
            Should -Throw '*cannot exceed*'
    }

    It 'keeps one shared download cache for a profile-isolated download root' {
        $shared = Get-SharedDownloadCacheRoot -DownloadRoot (Join-Path (Join-Path $script:IsoRoot 'dl\profiles') 'p1') -NoCreate
        $shared | Should -Be (Join-Path (Join-Path $script:IsoRoot 'dl') '_cache')
    }
}

Describe 'One Click freshness' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerWorkbench.psd1" -Force
        # The GUI cannot be dot-sourced, so the freshness function is lifted
        # out of it by AST the way the conflict-prompt tests above are. The
        # rest of the One Click path is covered by Tests/Invoke-WorkbenchSmoke.ps1.
        $t = $null; $e = $null
        $guiAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
        $fn = $guiAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-WorkbenchBuildIsCurrent' }, $false)
        $script:FreshnessAvailable = [bool]$fn
        if ($fn) { . ([scriptblock]::Create($fn.Extent.Text)) }

        $script:DataRoot = Join-Path $TestDrive 'freshness'
        $script:AppId = 'catalog:package-gfresh'
        $script:PreviousRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
        $env:APP_PACKAGER_WORKBENCH_ROOT = $script:DataRoot
        $stage = Join-Path $script:DataRoot 'stage'
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        $script:Record = Write-BuildRecord -StageRoot $stage -DataRoot $script:DataRoot -Manifest @{
            ApplicationId = $script:AppId; ProfileId = 'p1'; BuildId = (New-BuildId)
            ProfileRevision = 4; SoftwareVersion = '2.5.0'; AppName = 'G Fresh'
            PlanDigest = 'ABC'; ScriptSigning = @{ PolicyDigest = 'POLICY1' }
        }
    }
    AfterAll { $env:APP_PACKAGER_WORKBENCH_ROOT = $script:PreviousRoot }

    It 'seals the vendor version, profile revision and policy digest that the skip decision needs' {
        $record = Get-LatestBuildRecord -ApplicationId $script:AppId -ProfileId 'p1' -DataRoot $script:DataRoot
        $record.SoftwareVersion | Should -Be '2.5.0'
        [int]$record.ProfileRevision | Should -Be 4
        $record.PolicyDigest | Should -Be 'POLICY1'
    }

    It 'rebuilds when the profile revision moved' {
        $script:FreshnessAvailable | Should -BeTrue
        Test-WorkbenchBuildIsCurrent -ApplicationId $script:AppId -ProfileId 'p1' -Version '2.5.0' -ProfileRevision 5 -PolicyDigest 'POLICY1' | Should -BeFalse
    }

    It 'rebuilds when the signing policy digest moved' {
        Test-WorkbenchBuildIsCurrent -ApplicationId $script:AppId -ProfileId 'p1' -Version '2.5.0' -ProfileRevision 4 -PolicyDigest 'POLICY2' | Should -BeFalse
    }

    It 'rebuilds when no build has ever been sealed' {
        Test-WorkbenchBuildIsCurrent -ApplicationId 'catalog:package-never' -ProfileId 'p1' -Version '1.0' -ProfileRevision 0 -PolicyDigest '' | Should -BeFalse
    }

    It 'skips a build whose version, revision and policy all still match' {
        # The sealed record and the freshness check read the same version field.
        Test-WorkbenchBuildIsCurrent -ApplicationId $script:AppId -ProfileId 'p1' -Version '2.5.0' -ProfileRevision 4 -PolicyDigest 'POLICY1' | Should -BeTrue
    }
}

Describe 'CLI stage of a fixture packager' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerWorkbench.psd1" -Force
        $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
        $script:CliRoot = Join-Path $TestDrive 'cli'
        # A path with a space: the launcher quoting has to survive it.
        $script:FixtureRoot = Join-Path $script:CliRoot 'fixture repo'
        $script:FixturePackagers = Join-Path $script:FixtureRoot 'Packagers'
        New-Item -ItemType Directory -Path $script:FixturePackagers -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:FixtureRoot 'Lib') -Force | Out-Null
        foreach ($name in 'AppPackagerCommon', 'AppPackagerWorkbench', 'AppPackagerSigning') {
            foreach ($extension in 'psm1', 'psd1') {
                Copy-Item -LiteralPath (Join-Path $script:RepoRoot "Packagers\$name.$extension") -Destination $script:FixturePackagers -Force
            }
        }
        foreach ($folder in 'SuiteCommon', 'InstallerAnalysisCommon') {
            $source = Join-Path $script:RepoRoot "Lib\$folder"
            if (Test-Path -LiteralPath $source) {
                Copy-Item -LiteralPath $source -Destination (Join-Path $script:FixtureRoot 'Lib') -Recurse -Force
            }
        }

        $script:DownloadRoot = Join-Path $script:CliRoot 'dl'
        # The installer sits in the shared cache, which a profile-isolated
        # download root resolves back to; that is the contract the fixture
        # reads it through.
        $script:SourceFolder = Join-Path $script:DownloadRoot '_cache'
        New-Item -ItemType Directory -Path $script:SourceFolder -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $script:SourceFolder 'gfixture-setup.exe'), 'not a real installer')

        # Derived from Samples/package-template-exe.ps1 with the vendor lookup
        # and the download replaced by a pre-placed file, so the sweep needs no
        # network and no elevation.
        $fixture = @'
<#
Vendor: Contoso
App: G Fixture
CMName: G Fixture
VendorUrl: https://contoso.invalid
CPE: cpe:2.3:a:contoso:gfixture:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://contoso.invalid/notes
DownloadPageUrl: https://contoso.invalid/download
UpdateCadenceDays: 30

.SYNOPSIS
    Offline fixture packager for the regression suite.
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)

Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

$Version = "4.2.1"
$InstallerFileName = "gfixture-setup.exe"
$BaseDownloadRoot = Join-Path $DownloadRoot "GFixture"

if ($GetLatestVersionOnly) {
    Write-Output $Version
    exit 0
}

function Invoke-StageApp {
    $contentPath = Join-Path $BaseDownloadRoot $Version
    Initialize-Folder -Path $contentPath

    $source = Join-Path (Get-SharedDownloadCacheRoot -DownloadRoot $DownloadRoot) $InstallerFileName
    if (-not (Test-Path -LiteralPath $source)) { throw "Pre-placed installer not found: $source" }
    Copy-Item -LiteralPath $source -Destination (Join-Path $contentPath $InstallerFileName) -Force

    $installPs1 = '$proc = Start-Process -FilePath (Join-Path $PSScriptRoot ''gfixture-setup.exe'') -ArgumentList ''/S'' -Wait -PassThru -NoNewWindow' + "`r`n" + 'exit $proc.ExitCode'
    $uninstallPs1 = 'exit 0'

    Write-ContentWrappers -OutputPath $contentPath -InstallPs1Content $installPs1 -UninstallPs1Content $uninstallPs1

    Write-StageManifest -Path (Join-Path $contentPath "stage-manifest.json") -ManifestData @{
        AppName          = "G Fixture"
        Publisher        = "Contoso"
        SoftwareVersion  = $Version
        InstallerFile    = $InstallerFileName
        InstallerType    = "EXE"
        InstallArgs      = "/S"
        Architecture     = "x64"
        InstallationBehaviorType = "InstallForSystem"
        LogonRequirementType     = "WhetherOrNotUserLoggedOn"
        Detection        = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GFixture"
            ValueName           = "DisplayVersion"
            ExpectedValue       = $Version
            Operator            = "GreaterEquals"
            Is64Bit             = $true
        }
    }
    Write-Log "Stage complete               : $contentPath"
    return $contentPath
}

try {
    if ($StageOnly) { [void](Invoke-StageApp); exit 0 }
    if ($PackageOnly) { Write-Log "Package phase is not exercised by this fixture."; exit 0 }
    [void](Invoke-StageApp)
    exit 0
}
catch {
    Write-Log ("Fixture failed: " + $_.Exception.Message) -Level ERROR
    exit 1
}
'@
        [System.IO.File]::WriteAllText((Join-Path $script:FixturePackagers 'package-gfixture.ps1'), $fixture, (New-Object System.Text.ASCIIEncoding))

        $script:DataRoot = Join-Path $script:CliRoot 'data'
        $script:AppId = 'catalog:package-gfixture'
        $script:PreviousRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
        $env:APP_PACKAGER_WORKBENCH_ROOT = $script:DataRoot

        $profile = New-WorkbenchProfileObject -ApplicationId $script:AppId -Name 'Managed'
        $profile['Timing'] = @{ EstimatedMinutes = 6; MaximumMinutes = 24 }
        # Populated so this fixture exercises the stage path; the bare-profile
        # case below covers sections left unset.
        $profile['Install'] = @{ Command = 'install.bat' }
        $profile['Uninstall'] = @{ Command = 'uninstall.bat' }
        $profile['Application'] = @{ TitleMode = 'IncludeVersion' }
        $script:Profile = Save-Profile -Profile ([pscustomobject]$profile) -DataRoot $script:DataRoot -SetActive

        $cli = Join-Path $script:RepoRoot 'Invoke-AppPackagerBuild.ps1'
        $stdout = Join-Path $script:CliRoot 'cli.out'
        $stderr = Join-Path $script:CliRoot 'cli.err'
        $process = Start-Process -FilePath (Get-Command powershell.exe).Source -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $stdout -RedirectStandardError $stderr `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $cli + '"'),
                '-Application', 'catalog:package-gfixture',
                '-Profile', $script:Profile.ProfileId,
                '-Stage', '-Target', 'ContentOnly',
                '-PackagersRoot', ('"' + $script:FixturePackagers + '"'),
                '-DownloadRoot', ('"' + $script:DownloadRoot + '"'))
        $script:CliExit = $process.ExitCode
        $script:CliOut = (Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue)
        $script:CliErr = (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)

        $script:StagedManifestPath = Join-Path (Join-Path (Join-Path (Join-Path $script:DownloadRoot 'profiles') $script:Profile.ProfileId) 'GFixture\4.2.1') 'stage-manifest.json'
        $script:StagedManifest = if (Test-Path -LiteralPath $script:StagedManifestPath) {
            Get-Content -LiteralPath $script:StagedManifestPath -Raw | ConvertFrom-Json
        } else { $null }
    }
    AfterAll { $env:APP_PACKAGER_WORKBENCH_ROOT = $script:PreviousRoot }

    It 'stages the fixture without touching the network' {
        $script:CliExit | Should -Be 0 -Because ($script:CliErr + $script:CliOut)
    }

    It 'writes a schema-4 manifest carrying the build identity and plan digest' {
        $script:StagedManifest | Should -Not -BeNullOrEmpty
        $script:StagedManifest.SchemaVersion | Should -Be 4
        $script:StagedManifest.BuildId | Should -Match '^\d{8}-\d{6}-[0-9a-f]{8}$'
        $script:StagedManifest.ApplicationId | Should -Be $script:AppId
        $script:StagedManifest.ProfileId | Should -Be $script:Profile.ProfileId
        $script:StagedManifest.PlanDigest | Should -Match '^[0-9A-F]{64}$'
        @($script:StagedManifest.FileHashes).Count | Should -BeGreaterThan 0
    }

    It 'applies the profile timing to the staged plan' {
        [int]$script:StagedManifest.Timing.EstimatedMinutes | Should -Be 6
        [int]$script:StagedManifest.Timing.MaximumMinutes | Should -Be 24
    }

    It 'seals a build record for the staged BuildId' {
        $record = Get-LatestBuildRecord -ApplicationId $script:AppId -ProfileId $script:Profile.ProfileId -DataRoot $script:DataRoot
        $record | Should -Not -BeNullOrEmpty
        $record.BuildId | Should -Be $script:StagedManifest.BuildId
        $record.PlanDigest | Should -Be $script:StagedManifest.PlanDigest
    }

    It 'leaves no execution-policy relaxation in the unsigned wrappers it did not sign' {
        $bat = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $script:StagedManifestPath) 'install.bat') -Raw
        # Signing is off for this run, so the historical unsigned string stands.
        $bat | Should -Match '(?i)-ExecutionPolicy Bypass'
    }

    It 'produces the same resolved plan from a GUI-equivalent snapshot' {
        $cliSnapshot = Get-RunSnapshot -Path (Join-Path (Join-Path $script:DataRoot 'runs') ($script:StagedManifest.BuildId + '.json'))
        $guiSnapshot = New-RunSnapshot -ApplicationId $script:AppId -ProfileId $script:Profile.ProfileId -Target 'ContentOnly' `
            -DataRoot $script:DataRoot `
            -DownloadRoot (Join-Path (Join-Path $script:DownloadRoot 'profiles') $script:Profile.ProfileId) `
            -PackagerScriptPath (Join-Path $script:FixturePackagers 'package-gfixture.ps1')

        $volatile = @('BuildId', 'CreatedAt', 'Path')
        # Canonical form: key order is not part of the plan, and hashtable
        # enumeration order differs between the two hosts.
        $canonical = {
            param($Value, [int]$Depth = 0)
            if ($null -eq $Value -or $Depth -gt 12) { return $Value }
            if ($Value -is [string] -or $Value -is [System.ValueType]) { return $Value }
            if ($Value -is [System.Collections.IDictionary]) {
                $ordered = [ordered]@{}
                foreach ($key in @($Value.Keys | Sort-Object)) { $ordered[[string]$key] = (& $canonical $Value[$key] ($Depth + 1)) }
                return $ordered
            }
            if ($Value -is [pscustomobject]) {
                $ordered = [ordered]@{}
                foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) { $ordered[$property.Name] = (& $canonical $property.Value ($Depth + 1)) }
                return $ordered
            }
            if ($Value -is [System.Collections.IEnumerable]) {
                return @(foreach ($item in $Value) { & $canonical $item ($Depth + 1) })
            }
            return $Value
        }
        $normalize = {
            param($Snapshot)
            $copy = [ordered]@{}
            foreach ($property in @($Snapshot.PSObject.Properties | Sort-Object Name)) {
                if ($volatile -contains $property.Name) { continue }
                $copy[$property.Name] = $property.Value
            }
            ((& $canonical $copy 0) | ConvertTo-Json -Depth 12)
        }
        (& $normalize $guiSnapshot) | Should -Be (& $normalize $cliSnapshot)
    }

    It 'stages a profile that sets no command and no title mode' {
        # New-WorkbenchProfileObject seeds Application, Install and Uninstall
        # as empty maps; the CLI has to probe their fields under strict mode.
        $bare = New-WorkbenchProfileObject -ApplicationId $script:AppId -Name 'Bare'
        $saved = Save-Profile -Profile ([pscustomobject]$bare) -DataRoot $script:DataRoot
        $stdout = Join-Path $script:CliRoot 'bare.out'
        $stderr = Join-Path $script:CliRoot 'bare.err'
        $process = Start-Process -FilePath (Get-Command powershell.exe).Source -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $stdout -RedirectStandardError $stderr `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + (Join-Path $script:RepoRoot 'Invoke-AppPackagerBuild.ps1') + '"'),
                '-Application', 'catalog:package-gfixture',
                '-Profile', $saved.ProfileId,
                '-Stage', '-Target', 'ContentOnly',
                '-PackagersRoot', ('"' + $script:FixturePackagers + '"'),
                '-DownloadRoot', ('"' + $script:DownloadRoot + '"'))
        $process.ExitCode | Should -Be 0 -Because (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)
    }
}
