#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for the AppPackagerWorkbench module.

.DESCRIPTION
    Uses a temporary data root via APP_PACKAGER_WORKBENCH_ROOT. Touches no
    network, no MECM, no certificate store, and no real installer.

.EXAMPLE
    Invoke-Pester .\AppPackagerWorkbench.Tests.ps1
#>

BeforeDiscovery {
    Import-Module "$PSScriptRoot\AppPackagerWorkbench.psd1" -Force
}

BeforeAll {
    Import-Module "$PSScriptRoot\AppPackagerWorkbench.psd1" -Force

    $script:SavedRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
    $script:SavedSnapshot = $env:APP_PACKAGER_RUN_SNAPSHOT
    $script:SavedSigning = $env:APP_PACKAGER_SIGNING
    $script:TestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('apwb-tests-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:TestRoot -Force | Out-Null
    $env:APP_PACKAGER_WORKBENCH_ROOT = Join-Path $script:TestRoot 'data'
    Remove-Item Env:APP_PACKAGER_RUN_SNAPSHOT -ErrorAction SilentlyContinue
    Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue

    function New-TestFolder {
        param([string]$Name = ([guid]::NewGuid().ToString('N')))
        $path = Join-Path $script:TestRoot $Name
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        return $path
    }

    function New-TestFile {
        param([Parameter(Mandatory)][string]$Path, [string]$Content = 'test')
        $folder = Split-Path -Path $Path -Parent
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        [System.IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
        return $Path
    }

    function New-TestProfile {
        param([Parameter(Mandatory)][string]$ApplicationId, [string]$Name = 'Managed')
        $profile = New-WorkbenchProfileObject -ApplicationId $ApplicationId -Name $Name
        return (Save-Profile -Profile ([pscustomobject]$profile) -SetActive)
    }

    function New-TestStage {
        param([string]$InstallerFile = 'setup.exe', [string]$Version = '1.2.3')
        $stage = New-TestFolder
        New-TestFile -Path (Join-Path $stage $InstallerFile) -Content 'installer' | Out-Null
        New-TestFile -Path (Join-Path $stage 'install.ps1') -Content 'exit 0' | Out-Null
        New-TestFile -Path (Join-Path $stage 'uninstall.ps1') -Content 'exit 0' | Out-Null
        New-TestFile -Path (Join-Path $stage 'install.bat') -Content '@echo off' | Out-Null
        New-TestFile -Path (Join-Path $stage 'uninstall.bat') -Content '@echo off' | Out-Null
        return $stage
    }

    function New-TestManifest {
        param([string]$InstallerFile = 'setup.exe', [string]$Version = '1.2.3')
        return @{
            AppName         = 'Test App'
            Publisher       = 'Test Publisher'
            SoftwareVersion = $Version
            InstallerFile   = $InstallerFile
            InstallerType   = 'EXE'
            InstallArgs     = '/S'
            ProductCode     = '{11111111-2222-3333-4444-555555555555}'
            Detection       = @{ Type = 'RegistryKeyValue'; Hive = 'LocalMachine'; RegistryKeyRelative = 'SOFTWARE\Test'; ValueName = 'DisplayVersion' }
            Requirements    = @()
        }
    }
}

AfterAll {
    Remove-Item Env:APP_PACKAGER_RUN_SNAPSHOT -ErrorAction SilentlyContinue
    Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue
    if ($script:SavedRoot) { $env:APP_PACKAGER_WORKBENCH_ROOT = $script:SavedRoot }
    else { Remove-Item Env:APP_PACKAGER_WORKBENCH_ROOT -ErrorAction SilentlyContinue }
    if ($script:SavedSnapshot) { $env:APP_PACKAGER_RUN_SNAPSHOT = $script:SavedSnapshot }
    if ($script:SavedSigning) { $env:APP_PACKAGER_SIGNING = $script:SavedSigning }
    if ($script:TestRoot -and (Test-Path -LiteralPath $script:TestRoot)) {
        Remove-Item -LiteralPath $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Storage roots' {
    It 'prefers the environment variable over the preference' {
        $preferences = [pscustomobject]@{ WorkbenchDataRoot = 'C:\ignored' }
        Get-WorkbenchDataRoot -Preferences $preferences | Should -Be $env:APP_PACKAGER_WORKBENCH_ROOT
    }

    It 'uses the preference when no environment variable is set' {
        $saved = $env:APP_PACKAGER_WORKBENCH_ROOT
        Remove-Item Env:APP_PACKAGER_WORKBENCH_ROOT -ErrorAction SilentlyContinue
        try {
            $wanted = Join-Path $script:TestRoot 'pref-root'
            Get-WorkbenchDataRoot -Preferences ([pscustomobject]@{ WorkbenchDataRoot = $wanted }) | Should -Be $wanted
        }
        finally { $env:APP_PACKAGER_WORKBENCH_ROOT = $saved }
    }

    It 'resolves a profile download subroot back to the shared cache' {
        $root = Join-Path $script:TestRoot 'dl'
        $shared = Get-SharedDownloadCacheRoot -DownloadRoot $root -NoCreate
        $profileScoped = Get-SharedDownloadCacheRoot -DownloadRoot (Join-Path (Join-Path $root 'profiles') 'abc123') -NoCreate
        $profileScoped | Should -Be $shared
    }
}

Describe 'Identity' {
    It 'builds catalog and custom ids from a script path' {
        New-ApplicationId -Kind Catalog -ScriptPath 'C:\x\package-7zip.ps1' | Should -Be 'catalog:package-7zip'
        New-ApplicationId -Kind Custom -Name 'package-mine' | Should -Be 'custom:package-mine'
    }

    It 'builds a unique BYO id' {
        $first = New-ApplicationId -Kind Byo
        $second = New-ApplicationId -Kind Byo
        $first | Should -Match '^byo:[0-9a-f]{32}$'
        $first | Should -Not -Be $second
    }

    It 'round-trips an application key' {
        $key = ConvertTo-ApplicationKey -ApplicationId 'catalog:package-7zip'
        $key | Should -Be 'catalog_package-7zip'
        ConvertFrom-ApplicationKey -ApplicationKey $key | Should -Be 'catalog:package-7zip'
    }

    It 'rejects an application id with path characters' {
        { ConvertTo-ApplicationKey -ApplicationId 'catalog:..\..\evil' } | Should -Throw '*Invalid ApplicationId*'
    }
}

Describe 'Applications' {
    It 'discovers catalog and custom packagers and persisted BYO apps' {
        $catalogRoot = New-TestFolder
        $customRoot = New-TestFolder
        New-TestFile -Path (Join-Path $catalogRoot 'package-alpha.ps1') -Content '# packager' | Out-Null
        New-TestFile -Path (Join-Path $customRoot 'package-mine.ps1') -Content '# packager' | Out-Null
        $byoId = New-ApplicationId -Kind Byo
        Save-ApplicationDefinition -Definition ([pscustomobject]@{ ApplicationId = $byoId; DisplayName = 'Dropped App'; Origin = 'byo' }) | Out-Null

        $apps = Get-WorkbenchApplications -PackagersRoot $catalogRoot -CustomScriptRoot $customRoot
        @($apps | Where-Object { $_.ApplicationId -eq 'catalog:package-alpha' }).Count | Should -Be 1
        @($apps | Where-Object { $_.ApplicationId -eq 'custom:package-mine' }).Count | Should -Be 1
        @($apps | Where-Object { $_.ApplicationId -eq $byoId }).Count | Should -Be 1
    }

    It 'returns an unsaved default definition for an unknown application' {
        $definition = Get-ApplicationDefinition -ApplicationId 'catalog:package-not-saved'
        $definition.Persisted | Should -BeFalse
        $definition.ActiveProfileId | Should -Be 'default'
    }
}

Describe 'Profiles' {
    It 'raises the revision on every save' {
        $id = 'catalog:package-rev'
        $saved = New-TestProfile -ApplicationId $id
        $saved.Revision | Should -Be 1
        $again = Save-Profile -Profile $saved
        $again.Revision | Should -Be 2
    }

    It 'refuses a save built on a stale revision' {
        $id = 'catalog:package-conflict'
        $saved = New-TestProfile -ApplicationId $id
        Save-Profile -Profile $saved | Out-Null
        { Save-Profile -Profile $saved } | Should -Throw '*changed on disk*'
    }

    It 'refuses to save the virtual default profile' {
        $default = Get-Profile -ApplicationId 'catalog:package-default' -ProfileId 'default'
        { Save-Profile -Profile $default } | Should -Throw '*default profile cannot be saved*'
    }

    It 'rejects a profile schema newer than this build' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-future'
        $profile['SchemaVersion'] = 99
        { Save-Profile -Profile ([pscustomobject]$profile) } | Should -Throw '*newer than this build supports*'
    }

    It 'rejects an estimate greater than the maximum' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-timing'
        $profile['Timing'] = @{ EstimatedMinutes = 40; MaximumMinutes = 30 }
        { Save-Profile -Profile ([pscustomobject]$profile) } | Should -Throw '*cannot exceed*'
    }

    It 'rejects an invalid minute range' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-range'
        $profile['Timing'] = @{ EstimatedMinutes = 0 }
        { Save-Profile -Profile ([pscustomobject]$profile) } | Should -Throw '*positive whole number*'
    }

    It 'rejects a requirement operation without a RuleId' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-req'
        $profile['Requirements'] = @{ Operations = @(@{ Op = 'Add'; Rule = @{ ConditionId = 'cpu-arch' } }) }
        { Save-Profile -Profile ([pscustomobject]$profile) } | Should -Throw '*needs a RuleId*'
    }

    It 'rejects an absolute source file destination' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-abs'
        $profile['SourceFiles'] = @(@{ Asset = 'aaaaaaaaaaaaaaaa'; Destination = 'C:\evil.txt' })
        { Save-Profile -Profile ([pscustomobject]$profile) } | Should -Throw '*must be relative*'
    }

    It 'copies a profile under a new id at revision 1' {
        $id = 'catalog:package-copy'
        $original = New-TestProfile -ApplicationId $id -Name 'Managed'
        $copy = Copy-Profile -ApplicationId $id -ProfileId $original.ProfileId -NewName 'Lab'
        $copy.ProfileId | Should -Not -Be $original.ProfileId
        $copy.Name | Should -Be 'Lab'
        $copy.Revision | Should -Be 1
    }

    It 'lists the virtual default profile alongside saved ones' {
        $id = 'catalog:package-list'
        New-TestProfile -ApplicationId $id | Out-Null
        $profiles = Get-Profiles -ApplicationId $id
        @($profiles | Where-Object { $_.ProfileId -eq 'default' }).Count | Should -Be 1
        @($profiles).Count | Should -BeGreaterOrEqual 2
    }

    It 'distinguishes inherit from explicit none' {
        $profile = New-WorkbenchProfileObject -ApplicationId 'catalog:package-inherit'
        $profile['Install'] = @{ Command = 'setup.exe /S' }
        $inherited = Remove-ProfileField -Profile ([pscustomobject]$profile) -Field 'Install.Command' -Mode Inherit
        $explicit = Remove-ProfileField -Profile ([pscustomobject]$profile) -Field 'Install.Command' -Mode Explicit
        $inherited.Install.ContainsKey('Command') | Should -BeFalse
        $explicit.Install.ContainsKey('Command') | Should -BeTrue
        $explicit.Install['Command'] | Should -BeNullOrEmpty
    }
}

Describe 'Profile assets' {
    It 'copies an asset into managed storage with provenance and size' {
        $id = 'catalog:package-assets'
        $profile = New-TestProfile -ApplicationId $id
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'settings.json') -Content '{"a":1}'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $source
        $asset.AssetId | Should -Match '^[0-9a-f]{16}$'
        $asset.Size | Should -Be 7
        $asset.Provenance | Should -Be ([System.IO.Path]::GetFullPath($source))
        Test-Path -LiteralPath $asset.Path | Should -BeTrue
    }

    It 'removes a managed asset' {
        $id = 'catalog:package-assets-remove'
        $profile = New-TestProfile -ApplicationId $id
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'a.txt') -Content 'x'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $source
        Remove-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -AssetId $asset.AssetId | Should -BeTrue
        Remove-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -AssetId $asset.AssetId | Should -BeFalse
    }

    It 'fails a linked asset whose file is gone' {
        $id = 'catalog:package-linked'
        $profile = New-TestProfile -ApplicationId $id
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'linked.txt') -Content 'x'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $source -Linked
        Remove-Item -LiteralPath $source -Force
        { Resolve-ProfileAssetPath -ApplicationId $id -ProfileId $profile.ProfileId -Asset $asset } |
            Should -Throw '*is missing*'
    }

    It 'refuses assets on the default profile' {
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'b.txt') -Content 'x'
        { Add-ProfileAsset -ApplicationId 'catalog:package-x' -ProfileId 'default' -Path $source } |
            Should -Throw '*default profile*'
    }
}

Describe 'Resolve-EffectiveSettings' {
    It 'reports the source of each precedence level' {
        $globals = [pscustomobject]@{ EstimatedRuntimeMins = 15; MaximumRuntimeMins = 30 }
        $manifest = @{ AppName = 'Base App'; InstallCommandLine = 'install.bat' }
        $profile = [pscustomobject]@{
            Application = @{ DisplayName = 'Profile App' }
            Timing      = @{ EstimatedMinutes = 20 }
        }
        $runOverrides = @{ MaximumMinutes = 45 }

        $result = Resolve-EffectiveSettings -GlobalDefaults $globals -BaseManifest $manifest -Profile $profile -RunOverrides $runOverrides
        $result['DisplayName'].Value | Should -Be 'Profile App'
        $result['DisplayName'].Source | Should -Be 'Profile'
        $result['InstallCommand'].Source | Should -Be 'Packager'
        $result['EstimatedMinutes'].Value | Should -Be 20
        $result['EstimatedMinutes'].Source | Should -Be 'Profile'
        $result['MaximumMinutes'].Value | Should -Be 45
        $result['MaximumMinutes'].Source | Should -Be 'Run'
    }

    It 'reports NeedsStaging for an unresolved stage-time field' {
        $result = Resolve-EffectiveSettings
        $result['PinnedVersion'].Source | Should -Be 'NeedsStaging'
        $result['Detection'].Source | Should -Be 'NeedsStaging'
    }

    It 'lets a variant override outrank the profile and a target outrank the variant' {
        $profile = [pscustomobject]@{
            Timing   = @{ EstimatedMinutes = 10 }
            Variants = @{ Overrides = @{ 'x64' = @{ EstimatedMinutes = 20 } } }
        }
        $variantResult = Resolve-EffectiveSettings -Profile $profile -Variant 'x64'
        $variantResult['EstimatedMinutes'].Value | Should -Be 20
        $variantResult['EstimatedMinutes'].Source | Should -Be 'Variant'

        $targetResult = Resolve-EffectiveSettings -Profile $profile -Variant 'x64' -TargetOverrides @{ EstimatedMinutes = 25 }
        $targetResult['EstimatedMinutes'].Value | Should -Be 25
        $targetResult['EstimatedMinutes'].Source | Should -Be 'Target'
    }

    It 'treats an explicit null run override as a value, not as inherit' {
        $profile = [pscustomobject]@{ Install = @{ Command = 'setup.exe /S' }; Timing = @{ EstimatedMinutes = 10 } }
        $result = Resolve-EffectiveSettings -Profile $profile -RunOverrides @{ InstallCommand = $null; EstimatedMinutes = $null }
        $result['InstallCommand'].Source | Should -Be 'Run'
        $result['InstallCommand'].Value | Should -BeNullOrEmpty
        $result['EstimatedMinutes'].Source | Should -Be 'Run'
        $result['EstimatedMinutes'].Value | Should -BeNullOrEmpty
    }

    It 'treats an explicit profile null as a value, not as inherit' {
        $manifest = @{ Icon = 'app-icon.ico' }
        $profile = [pscustomobject]@{ Application = @{ Icon = $null } }
        $result = Resolve-EffectiveSettings -BaseManifest $manifest -Profile $profile
        $result['Icon'].Source | Should -Be 'Profile'
        $result['Icon'].Value | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-LegacyPreferenceMigration' {
    BeforeAll {
        $script:LegacyCatalogRoot = New-TestFolder
        $script:LegacyCustomRoot = New-TestFolder
        New-TestFile -Path (Join-Path $script:LegacyCatalogRoot 'package-legacy-a.ps1') -Content '# packager' | Out-Null
        New-TestFile -Path (Join-Path $script:LegacyCustomRoot 'package-legacy-b.ps1') -Content '# packager' | Out-Null
        $script:LegacyPreferences = [pscustomobject]@{
            DeploymentConditions = [pscustomobject]@{
                Apps = [pscustomobject]@{
                    'package-legacy-a' = [pscustomobject]@{
                        Architecture = 'x64'
                        Languages    = @('en-US')
                        Network      = 'VpnOnly'
                        Split        = 'Architecture'
                        InstallMode  = 'CurrentUser'
                        TitleMode    = 'NoVersion'
                    }
                }
            }
            CommandOverrides     = [pscustomobject]@{
                Apps = [pscustomobject]@{
                    'package-legacy-b' = [pscustomobject]@{ Install = 'setup.exe /q'; Uninstall = 'setup.exe /x' }
                }
            }
        }
    }

    It 'creates one active Migrated profile per legacy application' {
        $result = Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPreferences -PackagersRoot $script:LegacyCatalogRoot -CustomScriptRoot $script:LegacyCustomRoot
        $result.MigratedCount | Should -Be 2

        $profile = Get-Profile -ApplicationId 'catalog:package-legacy-a' -ProfileId (Get-ApplicationDefinition -ApplicationId 'catalog:package-legacy-a').ActiveProfileId
        $profile.Name | Should -Be 'Migrated'
        $profile.MigratedFrom | Should -Be 'package-legacy-a'
        $profile.InstallMode | Should -Be 'CurrentUser'
        $profile.Application.TitleMode | Should -Be 'NoVersion'
        @($profile.Requirements.Operations).Count | Should -Be 3
        $profile.Variants.Split.Split | Should -Be 'Architecture'
    }

    It 'gives a user-authored script a custom id, not a catalog one' {
        (Get-ApplicationDefinition -ApplicationId 'custom:package-legacy-b').Persisted | Should -BeTrue
        (Get-ApplicationDefinition -ApplicationId 'catalog:package-legacy-b').Persisted | Should -BeFalse
    }

    It 'migrates command overrides into the profile' {
        $definition = Get-ApplicationDefinition -ApplicationId 'custom:package-legacy-b'
        $profile = Get-Profile -ApplicationId 'custom:package-legacy-b' -ProfileId $definition.ActiveProfileId
        $profile.Install.Command | Should -Be 'setup.exe /q'
        $profile.Uninstall.Command | Should -Be 'setup.exe /x'
    }

    It 'is idempotent on a second run' {
        $second = Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPreferences -PackagersRoot $script:LegacyCatalogRoot -CustomScriptRoot $script:LegacyCustomRoot
        $second.MigratedCount | Should -Be 0
        $second.SkippedCount | Should -Be 2
    }

    It 'stays idempotent after the migrated profile is renamed' {
        $id = 'catalog:package-legacy-a'
        $profileId = (Get-ApplicationDefinition -ApplicationId $id).ActiveProfileId
        $stored = Get-Profile -ApplicationId $id -ProfileId $profileId
        $update = @{
            SchemaVersion  = 1; ProfileId = $profileId; ApplicationId = $id; Name = 'Renamed by hand'
            Revision       = $stored.Revision
            MigratedFrom   = $stored.MigratedFrom
            MigratedDigest = $stored.MigratedDigest
        }
        Save-Profile -Profile ([pscustomobject]$update) | Out-Null

        $third = Invoke-LegacyPreferenceMigration -Preferences $script:LegacyPreferences -PackagersRoot $script:LegacyCatalogRoot -CustomScriptRoot $script:LegacyCustomRoot
        $third.MigratedCount | Should -Be 0
        @(Get-Profiles -ApplicationId $id | Where-Object { -not $_.IsDefault }).Count | Should -Be 1
    }

    It 'does not reset an active profile the operator chose' {
        $id = 'catalog:package-legacy-c'
        New-TestFile -Path (Join-Path $script:LegacyCatalogRoot 'package-legacy-c.ps1') -Content '# packager' | Out-Null
        $chosen = New-TestProfile -ApplicationId $id -Name 'Operator choice'
        $preferences = [pscustomobject]@{
            DeploymentConditions = [pscustomobject]@{ Apps = [pscustomobject]@{ 'package-legacy-c' = [pscustomobject]@{ Architecture = 'x64' } } }
            CommandOverrides     = [pscustomobject]@{ Apps = [pscustomobject]@{} }
        }
        Invoke-LegacyPreferenceMigration -Preferences $preferences -PackagersRoot $script:LegacyCatalogRoot -CustomScriptRoot $script:LegacyCustomRoot | Out-Null
        (Get-ApplicationDefinition -ApplicationId $id).ActiveProfileId | Should -Be $chosen.ProfileId
    }

    It 'leaves the legacy preference keys in place' {
        $script:LegacyPreferences.DeploymentConditions.Apps.PSObject.Properties['package-legacy-a'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Run snapshots' {
    It 'writes an immutable snapshot carrying the profile and run overrides' {
        $id = 'catalog:package-snapshot'
        $profile = New-TestProfile -ApplicationId $id
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -Target 'MECM' -RunOverrides @{ EstimatedMinutes = 12 }
        Test-Path -LiteralPath $snapshot.Path | Should -BeTrue
        $snapshot.ProfileRevision | Should -Be $profile.Revision
        $snapshot.RunOverrides['EstimatedMinutes'] | Should -Be 12

        $reloaded = Get-RunSnapshot -Path $snapshot.Path
        $reloaded.BuildId | Should -Be $snapshot.BuildId
    }

    It 'never writes a run override back into the profile' {
        $id = 'catalog:package-nowriteback'
        $profile = New-TestProfile -ApplicationId $id
        New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -RunOverrides @{ EstimatedMinutes = 42 } | Out-Null
        $reloaded = Get-Profile -ApplicationId $id -ProfileId $profile.ProfileId
        $reloaded.Revision | Should -Be $profile.Revision
        ([string]$reloaded.Timing.EstimatedMinutes) | Should -BeNullOrEmpty
    }

    It 'rejects a run override estimate above the maximum' {
        $id = 'catalog:package-runrange'
        $profile = New-TestProfile -ApplicationId $id
        { New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -RunOverrides @{ EstimatedMinutes = 60; MaximumMinutes = 30 } } |
            Should -Throw '*cannot exceed*'
    }

    It 'rejects a snapshot schema newer than this build' {
        $path = Join-Path (New-TestFolder) 'future.json'
        [System.IO.File]::WriteAllText($path, '{"SchemaVersion":99}')
        { Get-RunSnapshot -Path $path } | Should -Throw '*newer than this build supports*'
    }

    It 'returns null when no snapshot is configured' {
        Get-RunSnapshot | Should -BeNullOrEmpty
    }
}

Describe 'Token binding' {
    It 'binds the four supported tokens' {
        $manifest = New-TestManifest
        Resolve-WorkbenchTokens -Text '{{InstallerFile}} {{Version}}' -ManifestData $manifest |
            Should -Be 'setup.exe 1.2.3'
        Resolve-WorkbenchTokens -Text '{{ContentRoot}}install.bat' -ManifestData $manifest -Context Command |
            Should -Be '%~dp0install.bat'
        Resolve-WorkbenchTokens -Text '{{ContentRoot}}' -ManifestData $manifest -Context Script |
            Should -Be '$PSScriptRoot'
    }

    It 'rejects an unknown token' {
        { Resolve-WorkbenchTokens -Text 'x {{Secret}}' -ManifestData (New-TestManifest) } |
            Should -Throw '*Unknown build token*'
    }

    It 'rejects a token with no value at build time' {
        $manifest = New-TestManifest
        $manifest['ProductCode'] = ''
        { Resolve-WorkbenchTokens -Text '{{ProductCode}}' -ManifestData $manifest } |
            Should -Throw '*no value at build time*'
    }
}

Describe 'Invoke-StageFinalization without a snapshot' {
    It 'adds only the schema-4 fields and changes nothing else' {
        $stage = New-TestStage
        $manifest = New-TestManifest
        $before = ($manifest.Clone())
        $before.Remove('Detection') | Out-Null

        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null | Out-Null

        foreach ($key in $before.Keys) {
            ([string]$manifest[$key]) | Should -Be ([string]$before[$key])
        }
        $manifest['ProfileId'] | Should -Be 'default'
        $manifest['ProfileRevision'] | Should -Be 0
        $manifest['DetectionSource'] | Should -Be 'Default'
        $manifest['InstallCommandLine'] | Should -Be 'install.bat'
        $manifest['SetupFile'] | Should -Be 'install.bat'
        # A native detector has no script to sign: NotApplicable from the
        # signing module, NotRequested from the stand-in when it is absent.
        $manifest['ScriptSigning'].Detection.Status | Should -BeIn @('NotApplicable', 'NotRequested')
    }

    It 'leaves the staged files untouched' {
        $stage = New-TestStage
        $manifest = New-TestManifest
        $before = @(Get-ChildItem -LiteralPath $stage -File | Sort-Object Name | ForEach-Object { $_.Name })
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null | Out-Null
        $after = @(Get-ChildItem -LiteralPath $stage -File | Sort-Object Name | ForEach-Object { $_.Name })
        ($after -join ',') | Should -Be ($before -join ',')
    }

    It 'throws when a signing switch is on and no certificate or module can serve it' {
        $stage = New-TestStage
        $manifest = New-TestManifest
        $env:APP_PACKAGER_SIGNING = '{"SignDeployment":true}'
        try {
            $expected = if (Get-Command Invoke-CategorySigning -ErrorAction SilentlyContinue) { '*SigningCertificateUnavailable*' } else { '*AppPackagerSigning is not loaded*' }
            { Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null } |
                Should -Throw $expected
        }
        finally { Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }
    }
}

Describe 'Invoke-StageFinalization with a profile' {
    It 'applies commands, timing, execution, detection and requirements' {
        $id = 'catalog:package-finalize'
        $profile = New-TestProfile -ApplicationId $id
        $data = $profile | ConvertTo-Json -Depth 12 | ConvertFrom-Json
        $update = @{
            SchemaVersion = 1
            ProfileId     = $profile.ProfileId
            ApplicationId = $id
            Name          = 'Managed'
            Revision      = $profile.Revision
            Install       = @{ Command = '{{ContentRoot}}install.bat /extra' }
            Detection     = @{ Mode = 'Custom'; Rule = @{ Type = 'File'; FilePath = 'C:\Program Files\Test\test.exe' } }
            Requirements  = @{ Operations = @(@{ Op = 'Add'; RuleId = 'cpu-arch'; Rule = @{ ConditionId = 'cpu-arch'; Value = 'x64' }; AppliesTo = @('*') }) }
            Timing        = @{ EstimatedMinutes = 12; MaximumMinutes = 40 }
            Execution     = @{ Context = 'InstallForSystem'; ScriptHost = 'x64' }
        }
        $saved = Save-Profile -Profile ([pscustomobject]$update)
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId -Target 'MECM'

        $stage = New-TestStage
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null

        $manifest['InstallCommandLine'] | Should -Be '%~dp0install.bat /extra'
        $manifest['DetectionSource'] | Should -Be 'Custom'
        $manifest['Detection']['Type'] | Should -Be 'File'
        $manifest['Timing']['EstimatedMinutes'] | Should -Be 12
        $manifest['Execution']['Context'] | Should -Be 'InstallForSystem'
        $manifest['InstallationBehaviorType'] | Should -Be 'InstallForSystem'
        @($manifest['Requirements']).Count | Should -Be 1
        $manifest['BuildId'] | Should -Be $snapshot.BuildId
    }

    It 'copies source files and records their hashes' {
        $id = 'catalog:package-sourcefiles'
        $profile = New-TestProfile -ApplicationId $id
        $file = New-TestFile -Path (Join-Path (New-TestFolder) 'settings.json') -Content '{"k":"v"}'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $file

        $update = @{
            SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
            Revision      = $profile.Revision
            SourceFiles   = @(@{ Asset = $asset.AssetId; Destination = 'config\settings.json'; Sha256 = $asset.Sha256; Size = $asset.Size })
        }
        $saved = Save-Profile -Profile ([pscustomobject]$update)
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId

        $stage = New-TestStage
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null

        Test-Path -LiteralPath (Join-Path $stage 'config\settings.json') | Should -BeTrue
        $entry = @($manifest['CustomAssets'] | Where-Object { $_.Category -eq 'SourceFile' })[0]
        $entry.Sha256 | Should -Be $asset.Sha256
    }

    It 'rejects a destination that escapes the content root' {
        $stage = New-TestStage
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'x.txt') -Content 'x'
        { & (Get-Module AppPackagerWorkbench) { param($r, $s, $d) Copy-WorkbenchStageFile -StageRoot $r -SourcePath $s -Destination $d } $stage $source '..\escape.txt' } |
            Should -Throw '*escapes the content root*'
    }

    It 'rejects a collision with a generated file' {
        $stage = New-TestStage
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'install.bat') -Content 'x'
        { & (Get-Module AppPackagerWorkbench) { param($r, $s, $d) Copy-WorkbenchStageFile -StageRoot $r -SourcePath $s -Destination $d } $stage $source 'INSTALL.BAT' } |
            Should -Throw '*collides with a file already in the stage root*'
    }

    It 'rejects a missing source file' {
        $stage = New-TestStage
        { & (Get-Module AppPackagerWorkbench) { param($r, $s, $d) Copy-WorkbenchStageFile -StageRoot $r -SourcePath $s -Destination $d } $stage (Join-Path $stage 'gone.txt') 'new.txt' } |
            Should -Throw '*is missing*'
    }

    It 'rejects a reparse point in the destination chain' {
        $stage = New-TestStage
        $realTarget = New-TestFolder
        $link = Join-Path $stage 'linked'
        $created = $false
        try {
            New-Item -ItemType Junction -Path $link -Target $realTarget -ErrorAction Stop | Out-Null
            $created = $true
        }
        catch { }
        if (-not $created) {
            Set-ItResult -Skipped -Because 'this host cannot create a junction'
            return
        }
        $source = New-TestFile -Path (Join-Path (New-TestFolder) 'y.txt') -Content 'y'
        { & (Get-Module AppPackagerWorkbench) { param($r, $s, $d) Copy-WorkbenchStageFile -StageRoot $r -SourcePath $s -Destination $d } $stage $source 'linked\y.txt' } |
            Should -Throw '*reparse point*'
    }

    It 'keeps two profiles from sharing output' {
        $id = 'catalog:package-twoprofiles'
        $first = New-TestProfile -ApplicationId $id -Name 'One'
        $second = Copy-Profile -ApplicationId $id -ProfileId $first.ProfileId -NewName 'Two'
        $firstSnapshot = New-RunSnapshot -ApplicationId $id -ProfileId $first.ProfileId
        $secondSnapshot = New-RunSnapshot -ApplicationId $id -ProfileId $second.ProfileId

        $firstStage = New-TestStage
        $secondStage = New-TestStage
        $firstManifest = New-TestManifest
        $secondManifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $firstStage -ManifestData $firstManifest -PackagerScriptPath $null -SnapshotPath $firstSnapshot.Path | Out-Null
        Invoke-StageFinalization -StageRoot $secondStage -ManifestData $secondManifest -PackagerScriptPath $null -SnapshotPath $secondSnapshot.Path | Out-Null

        $firstManifest['ProfileId'] | Should -Not -Be $secondManifest['ProfileId']
        $firstManifest['BuildId'] | Should -Not -Be $secondManifest['BuildId']

        Write-BuildRecord -StageRoot $firstStage -Manifest $firstManifest | Out-Null
        Write-BuildRecord -StageRoot $secondStage -Manifest $secondManifest | Out-Null
        $firstRecords = @(Get-BuildRecords -ApplicationId $id -ProfileId $first.ProfileId)
        $secondRecords = @(Get-BuildRecords -ApplicationId $id -ProfileId $second.ProfileId)
        $firstRecords.Count | Should -Be 1
        $secondRecords.Count | Should -Be 1
        $firstRecords[0].BuildId | Should -Not -Be $secondRecords[0].BuildId
    }

    It 'replaces the icon from a profile asset' {
        $id = 'catalog:package-icon'
        $profile = New-TestProfile -ApplicationId $id
        $icon = New-TestFile -Path (Join-Path (New-TestFolder) 'brand.png') -Content 'PNGDATA'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $icon
        $update = @{
            SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
            Revision      = $profile.Revision
            Application   = @{ Icon = @{ Asset = $asset.AssetId } }
        }
        $saved = Save-Profile -Profile ([pscustomobject]$update)
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId

        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'app-icon.ico') -Content 'old' | Out-Null
        $manifest = New-TestManifest
        $manifest['Icon'] = 'app-icon.ico'
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null

        $manifest['Icon'] | Should -Be 'app-icon.png'
        Test-Path -LiteralPath (Join-Path $stage 'app-icon.ico') | Should -BeFalse
    }
}

Describe 'Extend orchestrator' {
    It 'propagates a failing step exit code through powershell.exe' {
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'install-generated.ps1') -Content 'exit 7' | Out-Null
        $orchestrator = New-ExtendOrchestratorContent -Generated 'install-generated.ps1'
        [System.IO.File]::WriteAllText((Join-Path $stage 'install.ps1'), $orchestrator, [System.Text.Encoding]::ASCII)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $stage 'install.ps1') | Out-Null
        $LASTEXITCODE | Should -Be 7
    }

    It 'preserves 3010 from a hook when everything else succeeds' {
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'install-generated.ps1') -Content 'exit 0' | Out-Null
        New-TestFile -Path (Join-Path $stage 'install-after.ps1') -Content 'exit 3010' | Out-Null
        $orchestrator = New-ExtendOrchestratorContent -Generated 'install-generated.ps1' -After 'install-after.ps1'
        [System.IO.File]::WriteAllText((Join-Path $stage 'install.ps1'), $orchestrator, [System.Text.Encoding]::ASCII)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $stage 'install.ps1') | Out-Null
        $LASTEXITCODE | Should -Be 3010
    }

    It 'runs every step in its own process so an early exit cannot skip the rest' {
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'install-before.ps1') -Content 'exit 0' | Out-Null
        New-TestFile -Path (Join-Path $stage 'install-generated.ps1') -Content 'exit 0' | Out-Null
        $marker = Join-Path $stage 'after-ran.txt'
        New-TestFile -Path (Join-Path $stage 'install-after.ps1') -Content ("Set-Content -LiteralPath '$marker' -Value 'ran'; exit 0") | Out-Null
        $orchestrator = New-ExtendOrchestratorContent -Before 'install-before.ps1' -Generated 'install-generated.ps1' -After 'install-after.ps1'
        [System.IO.File]::WriteAllText((Join-Path $stage 'install.ps1'), $orchestrator, [System.Text.Encoding]::ASCII)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $stage 'install.ps1') | Out-Null
        $LASTEXITCODE | Should -Be 0
        Test-Path -LiteralPath $marker | Should -BeTrue
    }

    It 'runs the after step on failure and still returns the failing code' {
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'install-generated.ps1') -Content 'exit 5' | Out-Null
        $marker = Join-Path $stage 'cleanup.txt'
        New-TestFile -Path (Join-Path $stage 'install-after.ps1') -Content ("Set-Content -LiteralPath '$marker' -Value 'ran'; exit 0") | Out-Null
        $orchestrator = New-ExtendOrchestratorContent -Generated 'install-generated.ps1' -After 'install-after.ps1' -AfterRunsOnFailure $true
        [System.IO.File]::WriteAllText((Join-Path $stage 'install.ps1'), $orchestrator, [System.Text.Encoding]::ASCII)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $stage 'install.ps1') | Out-Null
        $LASTEXITCODE | Should -Be 5
        Test-Path -LiteralPath $marker | Should -BeTrue
    }

    It 'omits the execution policy argument in signed deployment mode' {
        $signed = New-ExtendOrchestratorContent -Generated 'install-generated.ps1' -SignedDeployment $true
        $signed | Should -Not -Match 'ExecutionPolicy'
        $unsigned = New-ExtendOrchestratorContent -Generated 'install-generated.ps1' -SignedDeployment $false
        $unsigned | Should -Match 'ExecutionPolicy'
    }

    It 'composes the orchestrator from the profile during finalization' {
        $id = 'catalog:package-extend'
        $profile = New-TestProfile -ApplicationId $id
        $hook = New-TestFile -Path (Join-Path (New-TestFolder) 'after.ps1') -Content 'exit 0'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $hook
        $update = @{
            SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
            Revision      = $profile.Revision
            Install       = @{ Mode = 'Extend'; After = $asset.AssetId }
        }
        $saved = Save-Profile -Profile ([pscustomobject]$update)
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId

        $stage = New-TestStage
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null

        Test-Path -LiteralPath (Join-Path $stage 'install-generated.ps1') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $stage 'install-after.ps1') | Should -BeTrue
        (Get-Content -LiteralPath (Join-Path $stage 'install.ps1') -Raw) | Should -Match 'Invoke-Step'
    }
}

Describe 'Build records' {
    It 'writes and reads a sealed build record' {
        $id = 'catalog:package-build'
        $profile = New-TestProfile -ApplicationId $id
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId
        $stage = New-TestStage
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null
        $manifest['PlanDigest'] = 'abc123'

        $record = Write-BuildRecord -StageRoot $stage -Manifest $manifest
        $record.BuildId | Should -Be $snapshot.BuildId
        (Get-LatestBuildRecord -ApplicationId $id).BuildId | Should -Be $snapshot.BuildId
    }

    It 'writes no record for a default-profile build' {
        $stage = New-TestStage
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null | Out-Null
        Write-BuildRecord -StageRoot $stage -Manifest $manifest | Should -BeNullOrEmpty
    }

    It 'resolves a stage manifest by exact BuildId' {
        $searchRoot = New-TestFolder
        $stage = Join-Path $searchRoot '1.0.0'
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        New-TestFile -Path (Join-Path $stage 'stage-manifest.json') -Content '{"SchemaVersion":4,"BuildId":"20260101-010101-aaaaaaaa"}' | Out-Null

        $resolved = Resolve-StageManifestForBuild -BuildId '20260101-010101-aaaaaaaa' -SearchRoot $searchRoot
        $resolved.StageRoot | Should -Be $stage
    }

    It 'refuses a build id no staged manifest carries' {
        $searchRoot = New-TestFolder
        { Resolve-StageManifestForBuild -BuildId 'nope' -SearchRoot $searchRoot } | Should -Throw '*stale build*'
    }
}

Describe 'BYO applications' {
    It 'persists a dropped installer and its first source revision' {
        $installer = New-TestFile -Path (Join-Path (New-TestFolder) 'internal-tool.exe') -Content 'binary-v1'
        $result = Save-ByoApplication -InstallerPath $installer -DisplayName 'Internal Tool' -Publisher 'Contoso' -SoftwareVersion '1.0'
        $result.Revision | Should -Be 1
        $result.Source.Sha256 | Should -Not -BeNullOrEmpty

        $definition = Get-ApplicationDefinition -ApplicationId $result.ApplicationId
        $definition.Persisted | Should -BeTrue
        $definition.UpdatePolicy | Should -Be 'Manual'
    }

    It 'flags an incompatible architecture change on a replacement installer' {
        $first = New-TestFile -Path (Join-Path (New-TestFolder) 'tool.exe') -Content 'v1'
        $created = Save-ByoApplication -InstallerPath $first -DisplayName 'Tool' -SoftwareVersion '1.0' `
            -Analysis ([pscustomobject]@{ InstallerType = 'EXE'; Architecture = 'x86'; AppName = 'Tool' })
        $second = New-TestFile -Path (Join-Path (New-TestFolder) 'tool.exe') -Content 'v2'
        $updated = Update-ByoSource -ApplicationId $created.ApplicationId -InstallerPath $second -SoftwareVersion '2.0' `
            -Analysis ([pscustomobject]@{ InstallerType = 'MSI'; Architecture = 'x64'; AppName = 'Tool' })

        $updated.Revision | Should -Be 2
        @($updated.IncompatibleOverrides) | Should -Contain 'Architecture'
        @($updated.IncompatibleOverrides) | Should -Contain 'InstallerType'
    }

    It 'refuses a source revision on a non-BYO application' {
        $installer = New-TestFile -Path (Join-Path (New-TestFolder) 'x.exe') -Content 'x'
        { Update-ByoSource -ApplicationId 'catalog:package-7zip' -InstallerPath $installer } |
            Should -Throw '*BYO applications only*'
    }
}

Describe 'Drafts' {
    It 'saves, reads and removes a draft' {
        $id = 'catalog:package-draft'
        Save-Draft -ApplicationId $id -ProfileId 'default' -Draft @{ Name = 'work in progress' } | Out-Null
        (Get-Draft -ApplicationId $id -ProfileId 'default').Name | Should -Be 'work in progress'
        Remove-Draft -ApplicationId $id -ProfileId 'default' | Should -BeTrue
        Get-Draft -ApplicationId $id -ProfileId 'default' | Should -BeNullOrEmpty
    }
}

Describe 'Bundles' {
    It 'exports and re-imports an application with hash verification' {
        $id = 'catalog:package-bundle'
        New-TestProfile -ApplicationId $id | Out-Null
        Save-ApplicationDefinition -Definition ([pscustomobject]@{ ApplicationId = $id; DisplayName = 'Bundled' }) | Out-Null

        $bundle = Join-Path (New-TestFolder) 'bundle.zip'
        Export-WorkbenchBundle -ApplicationId $id -Path $bundle | Out-Null
        Test-Path -LiteralPath $bundle | Should -BeTrue

        $targetId = 'catalog:package-bundle-copy'
        $imported = Import-WorkbenchBundle -Path $bundle -ApplicationId $targetId
        $imported.ApplicationId | Should -Be $targetId
        (Get-ApplicationDefinition -ApplicationId $targetId).Persisted | Should -BeTrue
    }

    It 'refuses to overwrite an existing application without -Force' {
        $id = 'catalog:package-bundle-existing'
        Save-ApplicationDefinition -Definition ([pscustomobject]@{ ApplicationId = $id; DisplayName = 'Existing' }) | Out-Null
        $bundle = Join-Path (New-TestFolder) 'existing.zip'
        Export-WorkbenchBundle -ApplicationId $id -Path $bundle | Out-Null
        { Import-WorkbenchBundle -Path $bundle } | Should -Throw '*already exists*'
    }

    It 'rejects a bundle schema newer than this build' {
        $staging = New-TestFolder
        New-Item -ItemType Directory -Path (Join-Path $staging 'application') -Force | Out-Null
        New-TestFile -Path (Join-Path $staging 'bundle.json') -Content '{"SchemaVersion":99,"ApplicationId":"catalog:package-x","Entries":[]}' | Out-Null
        $bundle = Join-Path (New-TestFolder) 'future.zip'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $bundle)
        { Import-WorkbenchBundle -Path $bundle } | Should -Throw '*newer than this build supports*'
    }
}

Describe 'Stage tree pruning' {
    It 'removes a source file the profile no longer carries on the next build' {
        $id = 'catalog:package-prune'
        $profile = New-TestProfile -ApplicationId $id
        $file = New-TestFile -Path (Join-Path (New-TestFolder) 'settings.json') -Content '{"k":"v"}'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $file

        $withFile = Save-Profile -Profile ([pscustomobject]@{
                SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
                Revision      = $profile.Revision
                SourceFiles   = @(@{ Asset = $asset.AssetId; Destination = 'config\settings.json' })
            })

        $stage = New-TestStage
        $firstSnapshot = New-RunSnapshot -ApplicationId $id -ProfileId $withFile.ProfileId
        $firstManifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $firstManifest -PackagerScriptPath $null -SnapshotPath $firstSnapshot.Path | Out-Null
        Write-BuildRecord -StageRoot $stage -Manifest $firstManifest | Out-Null
        Test-Path -LiteralPath (Join-Path $stage 'config\settings.json') | Should -BeTrue

        $withoutFile = Save-Profile -Profile ([pscustomobject]@{
                SchemaVersion = 1; ProfileId = $withFile.ProfileId; ApplicationId = $id; Name = 'Managed'
                Revision      = $withFile.Revision
                SourceFiles   = @()
            })
        $secondSnapshot = New-RunSnapshot -ApplicationId $id -ProfileId $withoutFile.ProfileId
        $secondManifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $secondManifest -PackagerScriptPath $null -SnapshotPath $secondSnapshot.Path | Out-Null

        Test-Path -LiteralPath (Join-Path $stage 'config\settings.json') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $stage 'config') | Should -BeFalse
        @($secondManifest['CustomAssets']).Count | Should -Be 0
    }

    It 'keeps a source file the profile still carries' {
        $id = 'catalog:package-prune-keep'
        $profile = New-TestProfile -ApplicationId $id
        $file = New-TestFile -Path (Join-Path (New-TestFolder) 'keep.json') -Content '{"k":1}'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $file
        $saved = Save-Profile -Profile ([pscustomobject]@{
                SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
                Revision      = $profile.Revision
                SourceFiles   = @(@{ Asset = $asset.AssetId; Destination = 'config\keep.json' })
            })

        $stage = New-TestStage
        foreach ($pass in 1, 2) {
            $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId
            $manifest = New-TestManifest
            Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null
            Write-BuildRecord -StageRoot $stage -Manifest $manifest | Out-Null
        }
        Test-Path -LiteralPath (Join-Path $stage 'config\keep.json') | Should -BeTrue
    }

    It 'removes a stale generated wrapper when the profile leaves Extend mode' {
        $id = 'catalog:package-prune-extend'
        $profile = New-TestProfile -ApplicationId $id
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'install-generated.ps1') -Content 'exit 0' | Out-Null
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId
        $manifest = New-TestManifest
        Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null
        Test-Path -LiteralPath (Join-Path $stage 'install-generated.ps1') | Should -BeFalse
    }
}

Describe 'Re-staging a profile source file' {
    It 'replaces its own identical copy when no build record was written' {
        $id = 'catalog:package-restage'
        $profile = New-TestProfile -ApplicationId $id
        $file = New-TestFile -Path (Join-Path (New-TestFolder) 'settings.txt') -Content 'value=1'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $file
        $saved = Save-Profile -Profile ([pscustomobject]@{
                SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
                Revision      = $profile.Revision
                SourceFiles   = @(@{ Asset = $asset.AssetId; Destination = 'config\settings.txt' })
            })

        $stage = New-TestStage
        foreach ($pass in 1, 2) {
            $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId
            $manifest = New-TestManifest
            Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path | Out-Null
        }
        (Get-Content -LiteralPath (Join-Path $stage 'config\settings.txt') -Raw).Trim() | Should -Be 'value=1'
    }

    It 'still refuses a differing file already at the destination' {
        $id = 'catalog:package-restage-collide'
        $profile = New-TestProfile -ApplicationId $id
        $file = New-TestFile -Path (Join-Path (New-TestFolder) 'settings.txt') -Content 'value=1'
        $asset = Add-ProfileAsset -ApplicationId $id -ProfileId $profile.ProfileId -Path $file
        $saved = Save-Profile -Profile ([pscustomobject]@{
                SchemaVersion = 1; ProfileId = $profile.ProfileId; ApplicationId = $id; Name = 'Managed'
                Revision      = $profile.Revision
                SourceFiles   = @(@{ Asset = $asset.AssetId; Destination = 'config\settings.txt' })
            })

        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'config\settings.txt') -Content 'generated by the packager' | Out-Null
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $saved.ProfileId
        $manifest = New-TestManifest
        { Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null -SnapshotPath $snapshot.Path } |
            Should -Throw '*collides with a file already in the stage root*'
    }
}

Describe 'Signing gate before the run starts' {
    AfterEach {
        Remove-Item Function:\Resolve-SigningCertificate -ErrorAction SilentlyContinue
    }

    It 'resolves the certificate while creating the snapshot' {
        $script:ResolveCalls = 0
        function global:Resolve-SigningCertificate {
            param($Policy)
            $script:ResolveCalls++
            return [pscustomobject]@{ Thumbprint = 'AA' }
        }
        $id = 'catalog:package-signgate-ok'
        $profile = New-TestProfile -ApplicationId $id
        New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -SigningPolicy @{ SignDeployment = $true } | Out-Null
        $script:ResolveCalls | Should -Be 1
    }

    It 'fails before the snapshot file is written when the certificate is unusable' {
        function global:Resolve-SigningCertificate {
            param($Policy)
            throw 'SigningCertificateUnavailable: no certificate thumbprint is selected.'
        }
        $id = 'catalog:package-signgate-fail'
        $profile = New-TestProfile -ApplicationId $id
        $before = @(Get-ChildItem -LiteralPath (Join-Path $env:APP_PACKAGER_WORKBENCH_ROOT 'runs') -File -ErrorAction SilentlyContinue).Count
        { New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -SigningPolicy @{ RequireDeployment = $true } } |
            Should -Throw '*SigningCertificateUnavailable*'
        @(Get-ChildItem -LiteralPath (Join-Path $env:APP_PACKAGER_WORKBENCH_ROOT 'runs') -File -ErrorAction SilentlyContinue).Count | Should -Be $before
    }

    It 'refuses a require switch when the signing module is absent' {
        $id = 'catalog:package-signgate-absent'
        $profile = New-TestProfile -ApplicationId $id
        { New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -SigningPolicy @{ RequireDetection = $true } } |
            Should -Throw '*AppPackagerSigning is not loaded*'
    }

    It 'leaves an all-off policy alone' {
        $id = 'catalog:package-signgate-off'
        $profile = New-TestProfile -ApplicationId $id
        $snapshot = New-RunSnapshot -ApplicationId $id -ProfileId $profile.ProfileId -SigningPolicy @{ SignDeployment = $false }
        Test-Path -LiteralPath $snapshot.Path | Should -BeTrue
    }
}

Describe 'Invoke-AppPackagerBuild CLI' {
    BeforeAll {
        $script:CliRoot = New-TestFolder
        Copy-Item -LiteralPath "$PSScriptRoot\AppPackagerWorkbench.psd1" -Destination $script:CliRoot -Force
        Copy-Item -LiteralPath "$PSScriptRoot\AppPackagerWorkbench.psm1" -Destination $script:CliRoot -Force
        Copy-Item -LiteralPath "$PSScriptRoot\AppPackagerSigning.psd1" -Destination $script:CliRoot -Force
        Copy-Item -LiteralPath "$PSScriptRoot\AppPackagerSigning.psm1" -Destination $script:CliRoot -Force
        $fixture = @(
            '[CmdletBinding()]'
            'param('
            '    [switch]$StageOnly, [switch]$PackageOnly, [switch]$GetLatestVersionOnly,'
            '    [string]$SiteCode, [string]$Comment, [string]$FileServerPath, [string]$DownloadRoot,'
            '    [string]$LogPath, [string]$ContentLayout, [int]$EstimatedRuntimeMins, [int]$MaximumRuntimeMins'
            ')'
            'if (-not $env:APP_PACKAGER_RUN_SNAPSHOT) { exit 91 }'
            'if (-not $env:APP_PACKAGER_WORKBENCH_ROOT) { exit 92 }'
            'exit 0'
        ) -join "`r`n"
        New-TestFile -Path (Join-Path $script:CliRoot 'package-clifixture.ps1') -Content $fixture | Out-Null
        $script:CliScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-AppPackagerBuild.ps1'
        $script:CliDownloadRoot = New-TestFolder

        function Invoke-Cli {
            # Streams go through files: a host running with a stop preference
            # would otherwise turn the child's first stderr line into a
            # terminating error before the exit code and output are read.
            param([string[]]$Arguments)
            $stdout = Join-Path $script:CliRoot ('cli-' + [guid]::NewGuid().ToString('N') + '.out')
            $stderr = [System.IO.Path]::ChangeExtension($stdout, '.err')
            $quoted = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $script:CliScript + '"')) + @($Arguments | ForEach-Object { if ($_ -match '[\s"]') { '"' + $_ + '"' } else { $_ } })
            $process = Start-Process -FilePath (Get-Command powershell.exe).Source -ArgumentList $quoted -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdout -RedirectStandardError $stderr
            $text = (Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue) + [Environment]::NewLine + (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)
            return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $text }
        }
    }

    It 'runs the virtual default profile' {
        $result = Invoke-Cli -Arguments @('-Application', 'package-clifixture.ps1', '-Profile', 'default', '-Stage',
            '-PackagersRoot', $script:CliRoot, '-DownloadRoot', $script:CliDownloadRoot)
        $result.Output | Should -Not -Match 'cannot be found on this object'
        $result.ExitCode | Should -Be 0
    }

    It 'runs Package without Stage' {
        $result = Invoke-Cli -Arguments @('-Application', 'package-clifixture.ps1', '-Profile', 'default', '-Package',
            '-Target', 'IntuneOnly', '-PackagersRoot', $script:CliRoot, '-DownloadRoot', $script:CliDownloadRoot)
        $result.Output | Should -Not -Match "property 'Count' cannot be found"
        $result.ExitCode | Should -Be 0
    }

    It 'takes the ad-hoc path for a BYO application instead of looking for a packager script' {
        $byoId = New-ApplicationId -Kind Byo
        Save-ApplicationDefinition -Definition ([pscustomobject]@{ ApplicationId = $byoId; DisplayName = 'No Source'; Origin = 'byo' }) | Out-Null
        $result = Invoke-Cli -Arguments @('-Application', $byoId, '-Stage', '-PackagersRoot', $PSScriptRoot, '-DownloadRoot', $script:CliDownloadRoot)
        $result.Output | Should -Not -Match 'No packager script'
        $result.Output | Should -Match 'has no stored source revision'
    }

    It 'reports a BYO source revision whose installer is gone' {
        $installer = New-TestFile -Path (Join-Path (New-TestFolder) 'byo-tool.exe') -Content 'binary'
        $created = Save-ByoApplication -InstallerPath $installer -DisplayName 'BYO Tool' -SoftwareVersion '1.0' `
            -Analysis ([pscustomobject]@{ InstallerType = 'EXE'; Architecture = 'x64'; AppName = 'BYO Tool' })
        Remove-Item -LiteralPath ([string]$created.Source.Path) -Force

        $result = Invoke-Cli -Arguments @('-Application', $created.ApplicationId, '-Stage', '-PackagersRoot', $PSScriptRoot, '-DownloadRoot', $script:CliDownloadRoot)
        $result.Output | Should -Not -Match 'No packager script'
        $result.Output | Should -Match 'is missing its installer'
    }
}

Describe 'Signing hand-off' {
    BeforeEach {
        $script:SigningCalls = @{ Stage = $null }
        function global:Test-DeploymentLauncherChain {
            param($StageRoot, $Manifest, $Policy)
            return [pscustomobject]@{
                Findings   = @([pscustomobject]@{ Message = 'bypass in launcher'; File = 'install.bat'; LineNumber = 2 })
                BypassFree = $false
                Enforced   = $false
            }
        }
    }
    AfterEach {
        Remove-Item Function:\Invoke-CategorySigning -ErrorAction SilentlyContinue
        Remove-Item Function:\Test-DeploymentLauncherChain -ErrorAction SilentlyContinue
    }

    It 'stores the findings array, not the chain summary' {
        function global:Invoke-CategorySigning {
            param($StageRoot, $ManifestData, $Policy)
            return [pscustomobject]@{
                PolicyDigest = 'abc'
                Detection    = [pscustomobject]@{ Status = 'NotApplicable' }
                Requirements = [pscustomobject]@{ Status = 'NotApplicable'; Items = @() }
                Deployment   = [pscustomobject]@{ Status = 'SignedAndVerified'; Files = @() }
            }
        }
        $stage = New-TestStage
        $manifest = New-TestManifest
        $env:APP_PACKAGER_SIGNING = '{"SignDeployment":true}'
        try {
            Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null | Out-Null
        }
        finally { Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }

        $findings = @($manifest['ScriptSigning']['LauncherFindings'])
        $findings.Count | Should -Be 1
        $findings[0]['Message'] | Should -Be 'bypass in launcher'
        $manifest['ScriptSigning']['LaunchersBypassFree'] | Should -BeFalse
    }

    It 'refuses the build when deployment signing fails and SignDeployment is on' {
        function global:Invoke-CategorySigning {
            param($StageRoot, $ManifestData, $Policy)
            return [pscustomobject]@{
                PolicyDigest = 'abc'
                Detection    = [pscustomobject]@{ Status = 'NotApplicable' }
                Requirements = [pscustomobject]@{ Status = 'NotApplicable'; Items = @() }
                Deployment   = [pscustomobject]@{ Status = 'Failed'; Reason = 'certificate key inaccessible'; Files = @() }
            }
        }
        $stage = New-TestStage
        $manifest = New-TestManifest
        $env:APP_PACKAGER_SIGNING = '{"SignDeployment":true,"RequireDeployment":false}'
        try {
            { Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null } |
                Should -Throw '*certificate key inaccessible*'
        }
        finally { Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }
    }

    It 'clears a previous build detection script before signing runs' {
        function global:Invoke-CategorySigning {
            param($StageRoot, $ManifestData, $Policy)
            $script:SigningCalls.Stage = @(Get-ChildItem -LiteralPath (Join-Path $StageRoot 'scripts') -Recurse -File -ErrorAction SilentlyContinue).Count
            return [pscustomobject]@{
                PolicyDigest = 'abc'
                Detection    = [pscustomobject]@{ Status = 'NotApplicable' }
                Requirements = [pscustomobject]@{ Status = 'NotApplicable'; Items = @() }
                Deployment   = [pscustomobject]@{ Status = 'NotRequested'; Files = @() }
            }
        }
        $stage = New-TestStage
        New-TestFile -Path (Join-Path $stage 'scripts\detect.ps1') -Content '# stale detector' | Out-Null
        New-TestFile -Path (Join-Path $stage 'scripts\requirements\cpu-arch.ps1') -Content '# stale rule' | Out-Null
        $manifest = New-TestManifest
        $env:APP_PACKAGER_SIGNING = '{"SignDetection":true}'
        try {
            Invoke-StageFinalization -StageRoot $stage -ManifestData $manifest -PackagerScriptPath $null | Out-Null
        }
        finally { Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }

        $script:SigningCalls.Stage | Should -Be 0
        Test-Path -LiteralPath (Join-Path $stage 'scripts\detect.ps1') | Should -BeFalse
    }
}

Describe 'Publication records' {
    It 'appends one JSON line per publication' {
        $id = 'catalog:package-publish'
        Write-PublicationRecord -ApplicationId $id -BuildId 'b1' -Target 'MECM' -Result 'Succeeded' | Out-Null
        $written = Write-PublicationRecord -ApplicationId $id -BuildId 'b2' -Target 'Intune' -Result 'Failed' -Message 'no tenant'
        @(Get-Content -LiteralPath $written.Path).Count | Should -Be 2
    }
}
