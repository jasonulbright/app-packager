#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Export and import of a site file between machines (Q-112h): the allowlist, the manifest, the hashes, the site code
    rules, hostile zips, and create or merge without a silent overwrite.

.DESCRIPTION
    The data root is a temp folder (APP_PACKAGER_DATA_ROOT). Site codes are fictional (AB1, CD2) and hosts use example.com.
    No window and no ConfigMgr. The file format is the one of App Packager Desktop.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking -Global
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $script:PlantedToken = 'ghp_PLANTEDTOKEN0123456789'
    $script:PlantedPassword = 'PlantedPassw0rd!'

    function script:New-Fixture {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) ('sitefile-' + [guid]::NewGuid().ToString('N'))
        $fx = [pscustomobject]@{ Root = $root; Data = (Join-Path $root 'data'); Out = (Join-Path $root 'out') }
        New-Item -ItemType Directory -Path $fx.Out -Force | Out-Null
        $script:Saved = @{}
        foreach ($n in 'APP_PACKAGER_DATA_ROOT', 'APP_PACKAGER_SITE', 'SUITE_CM_SITECODE') { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n) }
        $env:APP_PACKAGER_DATA_ROOT = $fx.Data
        $env:APP_PACKAGER_SITE = $null
        $env:SUITE_CM_SITECODE = $null
        $fx
    }
    function script:Remove-Fixture {
        param($Fixture)
        foreach ($n in $script:Saved.Keys) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
        Remove-Item -LiteralPath $Fixture.Root -Recurse -Force -ErrorAction SilentlyContinue
    }
    function script:Read-Json {
        param([string]$Path)
        if (Test-Path -LiteralPath $Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
    }
    function script:Write-Json {
        param([string]$Path, $Object)
        [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Object -Depth 12), (New-Object Text.UTF8Encoding($false)))
    }
    function script:Get-Settings {
        param([string]$Site)
        Read-Json -Path (Join-Path (Get-SuiteSiteFolder -SiteCode $Site) 'settings.json')
    }
    function script:New-Site {
        param([string]$Name, [string]$Provider, [string]$Settings, [string]$Label)
        $folder = Get-SuiteSiteFolder -SiteCode $Name
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Update-AppSiteIdentity -SiteName $Name -Provider $Provider
        if ($Label) {
            $p = Join-Path $folder 'site.json'
            $o = Read-Json -Path $p
            $o | Add-Member -NotePropertyName label -NotePropertyValue $Label -Force
            Write-Json -Path $p -Object $o
        }
        if ($Settings) { [IO.File]::WriteAllText((Join-Path $folder 'settings.json'), $Settings, (New-Object Text.UTF8Encoding($false))) }
        $folder
    }
    function script:Get-Hex {
        param([byte[]]$Bytes)
        $sha = [Security.Cryptography.SHA256]::Create()
        (([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant())
    }
    function script:ConvertTo-Bytes { param([string]$Text) (New-Object Text.UTF8Encoding($false)).GetBytes($Text) }
    # Builds a zip from [name, bytes] pairs. Duplicate and hostile names are allowed.
    function script:New-Zip {
        param([string]$Path, $Entries)
        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        $fs = [IO.File]::Create($Path)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($e in $Entries) {
            $entry = $zip.CreateEntry([string]$e[0])
            $s = $entry.Open(); $s.Write([byte[]]$e[1], 0, ([byte[]]$e[1]).Length); $s.Dispose()
        }
        $zip.Dispose(); $fs.Dispose()
        $Path
    }
    function script:Get-ManifestBytes {
        param([byte[]]$Settings, [byte[]]$Scope, [string]$Code = 'AB1', [string]$Kind = 'app-packager-site-file', $Version = 1, [string]$SettingsHash, $SettingsBytes)
        if (-not $SettingsHash) { $SettingsHash = Get-Hex $Settings }
        if ($null -eq $SettingsBytes) { $SettingsBytes = $Settings.Length }
        $m = [ordered]@{
            kind = $Kind; formatVersion = $Version; siteCode = $Code; label = 'Made by test'; createdUtc = '2026-11-23T10:00:00Z'
            entries = @(
                [ordered]@{ name = 'settings.json'; sha256 = $SettingsHash; bytes = $SettingsBytes }
                [ordered]@{ name = 'scope.json'; sha256 = (Get-Hex $Scope); bytes = $Scope.Length }
            )
        }
        ConvertTo-Bytes (ConvertTo-Json -InputObject $m -Depth 6)
    }
    $script:Scope = ConvertTo-Bytes '{ "items": [ { "id": "apd.wsus" } ], "neverExported": [] }'
    function script:New-GoodZip {
        param([string]$Path, [string]$Settings = '{ "Wsus": { "ServerName": "wsus.cd2.example.com" } }', [string]$Code = 'CD2')
        $s = ConvertTo-Bytes $Settings
        New-Zip -Path $Path -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope -Code $Code)), @('settings.json', $s), @('scope.json', $script:Scope))
    }
    function script:Get-EntryText {
        param([string]$Path)
        $zip = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $all = foreach ($e in $zip.Entries) { $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
            $all -join "`n"
        }
        finally { $zip.Dispose() }
    }
    function script:Get-TreeCount { param($Root) @(Get-ChildItem -LiteralPath $Root -Recurse -Force -ErrorAction SilentlyContinue).Count }

    $script:AbSettings = @'
{
  "SiteCode": "AB1",
  "ProviderMachineName": "cm.ab1.example.com",
  "DiscoverySqlServer": "sql.ab1.example.com",
  "FileShareRoot": "\\\\files.ab1.example.com\\content$",
  "GitHubToken": "ghp_PLANTEDTOKEN0123456789",
  "IntuneClientSecret": "PlantedPassw0rd!",
  "ContentDistribution": { "AutoDistribute": true, "DpGroupName": "AB1 DPs", "ApiToken": "ghp_PLANTEDTOKEN0123456789" },
  "Wsus": { "ServerName": "wsus.ab1.example.com", "UseSsl": true, "Password": "PlantedPassw0rd!" },
  "SiteHygiene": { "PauseDays": 3 },
  "RecentFiles": [ "C:\\secret\\thing.msi" ],
  "UserName": "someone"
}
'@
}

Describe 'Export-AppSiteFile' {
    BeforeEach { $script:fx = New-Fixture; New-Site -Name 'AB1' -Provider 'cm.ab1.example.com' -Settings $script:AbSettings -Label 'Production site' | Out-Null }
    AfterEach { Remove-Fixture $script:fx }

    It 'writes a zip with exactly three entries and only allowlisted keys' {
        $path = Join-Path $script:fx.Out 'site.zip'
        $r = Export-AppSiteFile -SiteName 'AB1' -Path $path
        $r.Ok | Should -BeTrue -Because $r.Reason
        $zip = [IO.Compression.ZipFile]::OpenRead($path)
        try { @($zip.Entries | ForEach-Object FullName | Sort-Object) | Should -Be @('manifest.json', 'scope.json', 'settings.json') } finally { $zip.Dispose() }
        $read = Read-AppSiteFile -Path $path
        $read.Ok | Should -BeTrue -Because $read.Reason
        $read.Content.SiteCode | Should -BeExactly 'AB1'
        $read.Content.Label | Should -Be 'Production site'
        $read.Content.Settings.Wsus.ServerName | Should -Be 'wsus.ab1.example.com'
        $read.Content.Settings.ContentDistribution.DpGroupName | Should -Be 'AB1 DPs'
        # the PowerShell provider key travels as ProviderMachine
        $read.Content.Settings.ProviderMachine | Should -Be 'cm.ab1.example.com'
    }

    It 'never carries a planted secret, the site code value or other data' {
        $path = Join-Path $script:fx.Out 'site.zip'
        (Export-AppSiteFile -SiteName 'AB1' -Path $path).Ok | Should -BeTrue
        $text = Get-EntryText -Path $path
        $text | Should -Not -Match ([regex]::Escape($script:PlantedToken))
        $text | Should -Not -Match ([regex]::Escape($script:PlantedPassword))
        foreach ($name in 'GitHubToken', 'IntuneClientSecret', 'ApiToken', 'someone', 'thing.msi', 'RecentFiles', 'UserName') { $text | Should -Not -Match $name }
        $read = Read-AppSiteFile -Path $path
        $read.Content.Settings.PSObject.Properties['SiteCode'] | Should -BeNullOrEmpty
        $read.Content.Settings.Wsus.PSObject.Properties['Password'] | Should -BeNullOrEmpty
    }

    It 'does not carry the records of the site and does not change the site' {
        $folder = Get-SuiteSiteFolder -SiteCode 'AB1'
        New-Item -ItemType Directory -Path (Join-Path $folder 'state') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'state\app-state.json') -Value '{"7zip":"24.1"}'
        $before = Get-ChildItem -LiteralPath $folder -Recurse -File | ForEach-Object { $_.FullName + (Get-FileHash $_.FullName).Hash }
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        (Get-EntryText -Path $path) | Should -Not -Match '24\.1'
        $after = Get-ChildItem -LiteralPath $folder -Recurse -File | ForEach-Object { $_.FullName + (Get-FileHash $_.FullName).Hash }
        $after | Should -Be $before
    }

    It 'fails for a site that does not exist or a bad name and writes nothing' {
        $path = Join-Path $script:fx.Out 'none.zip'
        (Export-AppSiteFile -SiteName 'ZZ9' -Path $path).Ok | Should -BeFalse
        (Export-AppSiteFile -SiteName '..\AB1' -Path $path).Ok | Should -BeFalse
        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'replaces an old file and leaves no temporary file' {
        $path = Join-Path $script:fx.Out 'again.zip'
        Set-Content -LiteralPath $path -Value 'old'
        (Export-AppSiteFile -SiteName 'AB1' -Path $path).Ok | Should -BeTrue
        (Read-AppSiteFile -Path $path).Ok | Should -BeTrue
        @(Get-ChildItem -LiteralPath $script:fx.Out -Filter '*.exporting').Count | Should -Be 0
    }

    It 'proposes a file name without path characters' {
        Get-AppSiteFileSuggestedName -SiteName 'AB1~2' -UtcNow ([datetime]'2026-11-23') | Should -Be 'site-AB1-2-20261123.zip'
    }
}

Describe 'Read-AppSiteFile refuses a file that breaks a rule' {
    BeforeEach { $script:fx = New-Fixture; New-Site -Name 'AB1' -Provider 'cm.ab1.example.com' -Settings $script:AbSettings | Out-Null }
    AfterEach { Remove-Fixture $script:fx }

    It 'refuses a file that is not a zip, is empty or is missing' {
        $t = Join-Path $script:fx.Out 'text.zip'; Set-Content -LiteralPath $t -Value 'this is not a zip'
        (Read-AppSiteFile -Path $t).Reason | Should -Match 'not a valid zip'
        $e = Join-Path $script:fx.Out 'empty.zip'; [IO.File]::WriteAllBytes($e, [byte[]]@())
        (Read-AppSiteFile -Path $e).Reason | Should -Match 'empty'
        (Read-AppSiteFile -Path (Join-Path $script:fx.Out 'missing.zip')).Reason | Should -Match 'does not exist'
    }

    It 'refuses a missing entry' {
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'a.zip') -Entries @(, @('settings.json', $s))
        (Read-AppSiteFile -Path $z).Reason | Should -Match "manifest.json' is missing"
    }

    It 'refuses a bad manifest: kind, version, shape' {
        $s = ConvertTo-Bytes '{ "Wsus": { "UseSsl": true } }'
        $cases = @(
            @{ M = (Get-ManifestBytes -Settings $s -Scope $script:Scope -Kind 'other'); R = 'not the manifest of a site file' }
            @{ M = (Get-ManifestBytes -Settings $s -Scope $script:Scope -Version 2); R = 'format version 2' }
            @{ M = (Get-ManifestBytes -Settings $s -Scope $script:Scope -Version 0); R = 'not valid' }
            @{ M = (ConvertTo-Bytes '[1,2]'); R = 'not a JSON object' }
            @{ M = (ConvertTo-Bytes '{ nope'); R = 'not a JSON object' }
        )
        foreach ($c in $cases) {
            $z = New-Zip -Path (Join-Path $script:fx.Out 'b.zip') -Entries @(@('manifest.json', $c.M), @('settings.json', $s), @('scope.json', $script:Scope))
            (Read-AppSiteFile -Path $z).Reason | Should -Match $c.R
        }
    }

    It 'refuses a site code that breaks the name rules: <Code>' -ForEach @(
        @{ Code = 'AB' }, @{ Code = 'ABCD' }, @{ Code = 'ab1' }, @{ Code = 'CON' }, @{ Code = 'SMS' }, @{ Code = 'NUL' },
        @{ Code = '../' }, @{ Code = 'A..' }, @{ Code = 'A/B' }, @{ Code = 'A\B' }, @{ Code = 'AB1~2' }, @{ Code = '' }
    ) {
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'c.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope -Code $Code)), @('settings.json', $s), @('scope.json', $script:Scope))
        $r = Read-AppSiteFile -Path $z
        $r.Ok | Should -BeFalse
        $r.Reason | Should -Match 'site code'
    }

    It 'refuses an entry whose hash or size does not match the manifest' {
        $good = ConvertTo-Bytes '{ "Wsus": { "ServerName": "wsus.cd2.example.com" } }'
        $bad = ConvertTo-Bytes '{ "Wsus": { "ServerName": "evil.example.com" } }'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'd.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $good -Scope $script:Scope)), @('settings.json', $bad), @('scope.json', $script:Scope))
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'SHA-256 of settings.json does not match'
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'e.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope -SettingsBytes 5)), @('settings.json', $s), @('scope.json', $script:Scope))
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'size of settings.json'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'f.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope -SettingsHash 'xyz')), @('settings.json', $s), @('scope.json', $script:Scope))
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'not a SHA-256'
    }

    It 'refuses a secret name or a setting outside the allowlist even when the hash matches' {
        foreach ($json in '{ "Wsus": { "ServerName": "a", "Password": "x" } }', '{ "GitHubToken": "ghp_x" }', '{ "SiteCode": "AB1" }', '{ "UserName": "x" }', '{ "Wsus": { "ServerName": "a" }, "RecentFiles": [] }', '{ "OneClick": { "Other": 1 } }') {
            $s = ConvertTo-Bytes $json
            $z = New-Zip -Path (Join-Path $script:fx.Out 'g.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope)), @('settings.json', $s), @('scope.json', $script:Scope))
            $r = Read-AppSiteFile -Path $z
            $r.Ok | Should -BeFalse -Because $json
            $r.Reason | Should -Match 'secret-like name|does not carry'
        }
    }

    It 'refuses an unknown item in the scope' {
        $s = ConvertTo-Bytes '{}'
        $scope = ConvertTo-Bytes '{ "items": [ { "id": "apd.github-token" } ] }'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'h.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $scope)), @('settings.json', $s), @('scope.json', $scope))
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'does not know'
    }
}

Describe 'Read-AppSiteFile and a hostile zip' {
    BeforeEach { $script:fx = New-Fixture }
    AfterEach { Remove-Fixture $script:fx }

    It 'refuses the entry name <Name> and writes nothing' -ForEach @(
        @{ Name = '../evil.txt' }, @{ Name = '..\evil.txt' }, @{ Name = '..' }, @{ Name = '/etc/evil' }, @{ Name = '\\server\share\evil' },
        @{ Name = 'C:\Windows\evil.dll' }, @{ Name = 'C:evil' }, @{ Name = 'sub/settings.json' }, @{ Name = 'settings.json/../../evil' },
        @{ Name = 'sub\scope.json' }, @{ Name = '~/evil' }, @{ Name = 'evil.txt' }, @{ Name = 'Settings.json' }, @{ Name = 'settings.json ' },
        @{ Name = 'settings.json:stream' }, @{ Name = 'manifest.json/' }
    ) {
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'hostile.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope)), @('settings.json', $s), @('scope.json', $script:Scope), @($Name, (ConvertTo-Bytes 'MZ-evil')))
        $before = Get-TreeCount -Root $script:fx.Root
        $r = Read-AppSiteFile -Path $z
        $r.Ok | Should -BeFalse
        Get-TreeCount -Root $script:fx.Root | Should -Be $before
        Test-Path -LiteralPath (Join-Path $script:fx.Root 'evil.txt') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:fx.Out 'evil.txt') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path (Split-Path $script:fx.Root -Parent) 'evil.txt') | Should -BeFalse
    }

    It 'refuses a hostile name in place of a required entry' {
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'i.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope)), @('../settings.json', $s), @('scope.json', $script:Scope))
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'not a plain file name'
    }

    It 'refuses a duplicate entry' {
        $s = ConvertTo-Bytes '{}'
        $z = New-Zip -Path (Join-Path $script:fx.Out 'j.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope)), @('settings.json', $s), @('scope.json', $script:Scope), @('settings.json', (ConvertTo-Bytes '{"Wsus":{"UseSsl":true}}')))
        (Read-AppSiteFile -Path $z).Ok | Should -BeFalse
    }

    It 'refuses an entry that unpacks too large and a file that is too large' {
        $s = ConvertTo-Bytes '{}'
        $bomb = New-Object byte[] 510000
        for ($i = 0; $i -lt $bomb.Length; $i++) { $bomb[$i] = 32 }
        $z = New-Zip -Path (Join-Path $script:fx.Out 'bomb.zip') -Entries @(@('manifest.json', (Get-ManifestBytes -Settings $s -Scope $script:Scope)), @('settings.json', $bomb), @('scope.json', $script:Scope))
        (Get-Item $z).Length | Should -BeLessThan 1000000
        (Read-AppSiteFile -Path $z).Reason | Should -Match 'larger than'
        $big = Join-Path $script:fx.Out 'big.zip'
        [IO.File]::WriteAllBytes($big, (New-Object byte[] 1000001))
        (Read-AppSiteFile -Path $big).Reason | Should -Match 'larger than'
    }

    It 'refuses a truncated zip' {
        New-Site -Name 'AB1' -Provider 'cm.ab1.example.com' -Settings $script:AbSettings | Out-Null
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $bytes = [IO.File]::ReadAllBytes($path)
        $cut = Join-Path $script:fx.Out 'cut.zip'
        [IO.File]::WriteAllBytes($cut, $bytes[0..([int]($bytes.Length / 2))])
        (Read-AppSiteFile -Path $cut).Ok | Should -BeFalse
    }

    It 'cleans a long label with control characters' {
        $s = ConvertTo-Bytes '{}'
        $m = [ordered]@{
            kind = 'app-packager-site-file'; formatVersion = 1; siteCode = 'AB1'; label = ('Line' + [char]1 + 'one' + ('x' * 200)); createdUtc = '2026-11-23T10:00:00Z'
            entries = @(
                [ordered]@{ name = 'settings.json'; sha256 = (Get-Hex $s); bytes = $s.Length }
                [ordered]@{ name = 'scope.json'; sha256 = (Get-Hex $script:Scope); bytes = $script:Scope.Length }
            )
        }
        $z = New-Zip -Path (Join-Path $script:fx.Out 'l.zip') -Entries @(@('manifest.json', (ConvertTo-Bytes (ConvertTo-Json -InputObject $m -Depth 6))), @('settings.json', $s), @('scope.json', $script:Scope))
        $r = Read-AppSiteFile -Path $z
        $r.Ok | Should -BeTrue -Because $r.Reason
        $r.Content.Label.Length | Should -BeLessOrEqual 80
        $r.Content.Label | Should -Not -Match '[\x00-\x1f]'
    }
}

Describe 'Invoke-AppSiteFileImport' {
    BeforeEach { $script:fx = New-Fixture; New-Site -Name 'AB1' -Provider 'cm.ab1.example.com' -Settings $script:AbSettings -Label 'Production site' | Out-Null }
    AfterEach { Remove-Fixture $script:fx }

    It 'creates a new site (a different environment for a code in use) with the default items and no secret' {
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        Get-AppSiteFileNewName -Content $content | Should -BeExactly 'AB1~2'

        $before = Get-TreeCount -Root $script:fx.Data
        $plan = Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~2' -DryRun
        $plan.Outcome | Should -Be 'Done'
        $plan.Created | Should -BeTrue
        Get-TreeCount -Root $script:fx.Data | Should -Be $before
        Test-Path -LiteralPath (Get-SuiteSiteFolder -SiteCode 'AB1~2') | Should -BeFalse

        $run = Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~2'
        $run.Outcome | Should -Be 'Done'
        $settings = Get-Settings -Site 'AB1~2'
        $settings.ContentDistribution.DpGroupName | Should -Be 'AB1 DPs'
        # environment values are off by default
        $settings.PSObject.Properties['Wsus'] | Should -BeNullOrEmpty
        $settings.PSObject.Properties['ProviderMachine'] | Should -BeNullOrEmpty
        (Get-Content -LiteralPath (Join-Path (Get-SuiteSiteFolder -SiteCode 'AB1~2') 'settings.json') -Raw) | Should -Not -Match 'PLANTED'
        $identity = Read-Json -Path (Join-Path (Get-SuiteSiteFolder -SiteCode 'AB1~2') 'site.json')
        $identity.siteCode | Should -BeExactly 'AB1'
        $identity.label | Should -Be 'Production site'
        $log = Read-Json -Path (Join-Path (Get-SuiteSiteFolder -SiteCode 'AB1~2') 'imported-from.json')
        $log.sourceKind | Should -Be 'file'
    }

    It 'copies the environment values and adds the provider when they are named' {
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        $all = @(Get-AppSiteFileItem | ForEach-Object Id)
        (Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~2' -ItemId $all).Outcome | Should -Be 'Done'
        (Get-Settings -Site 'AB1~2').Wsus.ServerName | Should -Be 'wsus.ab1.example.com'
        @((Read-Json -Path (Join-Path (Get-SuiteSiteFolder -SiteCode 'AB1~2') 'site.json')).providers) | Should -Contain 'cm.ab1.example.com'
    }

    It 'keeps an existing value by default and replaces only on choice, with a verified backup' {
        New-Site -Name 'CD2' -Provider 'cm.cd2.example.com' -Settings '{ "ContentDistribution": { "DpGroupName": "CD2 own group" }, "SiteHygiene": { "PauseDays": 9 } }' | Out-Null
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        $file = Join-Path (Get-SuiteSiteFolder -SiteCode 'CD2') 'settings.json'

        $keep = Invoke-AppSiteFileImport -Content $content -TargetSite 'CD2'
        $s = Read-Json -Path $file
        $s.ContentDistribution.DpGroupName | Should -Be 'CD2 own group'
        $s.SiteHygiene.PauseDays | Should -Be 9
        @($keep.Rows | Where-Object Outcome -eq 'Kept').Count | Should -BeGreaterOrEqual 2
        @(Get-ChildItem -LiteralPath (Get-SuiteSiteFolder -SiteCode 'CD2') -Filter 'settings.json.before-import-*').Count | Should -Be 0

        $replace = Invoke-AppSiteFileImport -Content $content -TargetSite 'CD2' -ReplaceId @('apd.content-distribution') -UtcNow ([datetime]'2026-11-23T10:01:00Z')
        $s = Read-Json -Path $file
        $s.ContentDistribution.DpGroupName | Should -Be 'AB1 DPs'
        $s.SiteHygiene.PauseDays | Should -Be 9
        @($replace.Rows | Where-Object Outcome -eq 'Replaced').Count | Should -Be 1
        @(Get-ChildItem -LiteralPath (Get-SuiteSiteFolder -SiteCode 'CD2') -Filter 'settings.json.before-import-*').Count | Should -Be 1
    }

    It 'changes nothing the second time' {
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        (Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~2').Outcome | Should -Be 'Done'
        $file = Join-Path (Get-SuiteSiteFolder -SiteCode 'AB1~2') 'settings.json'
        $hash = (Get-FileHash $file).Hash
        (Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~2').Outcome | Should -Be 'NothingToDo'
        (Get-FileHash $file).Hash | Should -Be $hash
    }

    It 'writes the provider to ProviderMachineName when the target already uses that key' {
        New-Site -Name 'CD2' -Provider 'cm.cd2.example.com' -Settings '{ "ProviderMachineName": "cm.other.example.com" }' | Out-Null
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        $r = Invoke-AppSiteFileImport -Content $content -TargetSite 'CD2' -ItemId @('apd.provider')
        @($r.Rows | Where-Object Id -eq 'apd.provider')[0].Outcome | Should -Be 'Kept'
        (Read-Json -Path (Join-Path (Get-SuiteSiteFolder -SiteCode 'CD2') 'settings.json')).ProviderMachineName | Should -Be 'cm.other.example.com'
    }

    It 'refuses a target that does not follow the name rules or is a new name that is not the next free one' {
        $path = Join-Path $script:fx.Out 'site.zip'
        Export-AppSiteFile -SiteName 'AB1' -Path $path | Out-Null
        $content = (Read-AppSiteFile -Path $path).Content
        (Invoke-AppSiteFileImport -Content $content -TargetSite '..\CD2').Outcome | Should -Be 'Refused'
        (Invoke-AppSiteFileImport -Content $content -TargetSite 'ZZ9').Outcome | Should -Be 'Refused'
        (Invoke-AppSiteFileImport -Content $content -TargetSite 'AB1~7').Outcome | Should -Be 'Refused'
        Test-Path -LiteralPath (Get-SuiteSiteFolder -SiteCode 'ZZ9') | Should -BeFalse
    }

    It 'names the first code that is free: the code itself when no site has it' {
        New-Site -Name 'CD2' -Provider 'cm.cd2.example.com' -Settings '{}' | Out-Null
        $z = New-GoodZip -Path (Join-Path $script:fx.Out 'g.zip') -Code 'EF3'
        $content = (Read-AppSiteFile -Path $z).Content
        Get-AppSiteFileNewName -Content $content | Should -BeExactly 'EF3'
    }
}
