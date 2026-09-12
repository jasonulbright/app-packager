#Requires -Modules Pester

<#
.SYNOPSIS
    Update preservation: what survives an install-root replacement.

.DESCRIPTION
    install.ps1 replaces the whole install folder on an update and restores
    only the state Get-PreservedStateFile names: every *.json, the Logs
    folder, and - when the extracted release tree is supplied - every file the
    release does not ship at the same relative path, except the paths it lists
    as retired. The function is lifted out of install.ps1 by AST so the test
    runs the shipped rule rather than a copy, and the replacement itself is
    simulated against a fixture tree: preserve, wipe, extract, restore.

    The workbench data root lives outside the install folder, so an update
    must not touch it at all.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $tokens = $null
    $parseErrors = $null
    $installerAst = [System.Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $script:RepoRoot 'install.ps1'), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors) { throw ($parseErrors.Message -join '; ') }
    $function = $installerAst.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PreservedStateFile'
    }, $false)
    if (-not $function) { throw 'install.ps1 no longer defines Get-PreservedStateFile.' }
    . ([scriptblock]::Create($function.Extent.Text))

    function New-InstallFixture {
        <#
            An install folder as a user would have it after running the app:
            shipped files, saved state, logs, and a packager the user wrote
            into the catalog folder.
        #>
        param([Parameter(Mandatory)][string]$Root)

        New-Item -ItemType Directory -Path (Join-Path $Root 'Packagers') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $Root 'Logs') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $Root 'Packagers\Icons') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $Root 'Packagers\Templates') -Force | Out-Null

        Set-Content -LiteralPath (Join-Path $Root 'start-apppackager.ps1') -Value '# shipped v1' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'AppPackager.preferences.json') -Value '{"SiteCode":"MCM"}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'AppPackager.windowstate.json') -Value '{"Width":1200}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Logs\apppackager.log') -Value 'log line' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\AppPackagerCommon.psm1') -Value '# shipped module v1' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\package-contoso-inhouse.ps1') -Value '# user authored packager' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\package-contoso-draft.notps1') -Value '# disabled user packager' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\package-7zip.ps1') -Value '# shipped packager v1' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\package-specexec-mitigations.ps1') -Value '# retired catalog packager' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\Icons\contoso.ico') -Value 'icon bytes' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\Templates\contoso-options.ps1') -Value '# user authored template' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\Templates\options-panel-template.ps1') -Value '# shipped template v1' -Encoding ASCII
    }

    function New-WorkbenchDataFixture {
        # The data root the workbench uses, deliberately outside the install
        # folder.
        param([Parameter(Mandatory)][string]$Root)

        New-Item -ItemType Directory -Path (Join-Path $Root 'applications\catalog_package-7zip\profiles\p1') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $Root 'builds\catalog_package-7zip\p1\20260101-000000-aaaaaaaa') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $Root 'applications\catalog_package-7zip\application.json') -Value '{"ActiveProfileId":"p1"}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'applications\catalog_package-7zip\profiles\p1\profile.json') -Value '{"Name":"Managed","Revision":3}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'builds\catalog_package-7zip\p1\20260101-000000-aaaaaaaa\build.json') -Value '{"BuildId":"20260101-000000-aaaaaaaa"}' -Encoding ASCII
    }

    function New-ReleaseStage {
        # What the release zip extracts to: shipped files only.
        param([Parameter(Mandatory)][string]$Root)

        New-Item -ItemType Directory -Path (Join-Path $Root 'Packagers\Templates') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\Templates\options-panel-template.ps1') -Value '# shipped template v2' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'start-apppackager.ps1') -Value '# shipped v2' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\AppPackagerCommon.psm1') -Value '# shipped module v2' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\package-7zip.ps1') -Value '# shipped packager v2' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $Root 'Packagers\retired-packagers.txt') -Value @(
            '# Catalog scripts removed from the release.'
            'package-specexec-mitigations.ps1'
        ) -Encoding ASCII
    }

    function Invoke-SimulatedUpdate {
        <#
            The install.ps1 update sequence: collect preserved state, remove
            the folder contents, copy the new release in, restore the state.
        #>
        param(
            [Parameter(Mandatory)][string]$InstallPath,
            [Parameter(Mandatory)][string]$StagePath,
            [Parameter(Mandatory)][string]$BackupPath
        )

        New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
        $preserved = @(Get-PreservedStateFile -Root $InstallPath -StagePath $StagePath)
        foreach ($relative in $preserved) {
            $destination = Join-Path $BackupPath $relative
            New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $InstallPath $relative) -Destination $destination -Force
        }

        Get-ChildItem -LiteralPath $InstallPath -Force | Remove-Item -Recurse -Force
        Get-ChildItem -LiteralPath $StagePath -Force | Copy-Item -Destination $InstallPath -Recurse -Force

        foreach ($relative in $preserved) {
            $destination = Join-Path $InstallPath $relative
            New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $BackupPath $relative) -Destination $destination -Force
        }
        return $preserved
    }
}

Describe 'Update preservation' {
    BeforeAll {
        $script:InstallPath = Join-Path $TestDrive 'AppPackager'
        $script:DataRoot = Join-Path $TestDrive 'AppPackagerData\Workbench'
        $script:StagePath = Join-Path $TestDrive 'release-stage'
        $script:BackupPath = Join-Path $TestDrive 'backup'

        New-Item -ItemType Directory -Path $script:InstallPath -Force | Out-Null
        New-InstallFixture -Root $script:InstallPath
        New-WorkbenchDataFixture -Root $script:DataRoot
        New-ReleaseStage -Root $script:StagePath

        $script:DataBefore = @(Get-ChildItem -LiteralPath $script:DataRoot -Recurse -File |
            ForEach-Object { '{0}|{1}' -f $_.FullName.Substring($script:DataRoot.Length), (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } | Sort-Object)

        $script:Preserved = Invoke-SimulatedUpdate -InstallPath $script:InstallPath -StagePath $script:StagePath -BackupPath $script:BackupPath
    }

    It 'replaces the shipped files with the new release' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'start-apppackager.ps1') -Raw | Should -Match 'shipped v2'
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\AppPackagerCommon.psm1') -Raw | Should -Match 'shipped module v2'
    }

    It 'restores every json state file and the log folder' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'AppPackager.preferences.json') -Raw | Should -Match 'MCM'
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'AppPackager.windowstate.json') -Raw | Should -Match '1200'
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Logs\apppackager.log') -Raw | Should -Match 'log line'
        $script:Preserved | Should -Contain 'AppPackager.preferences.json'
        $script:Preserved | Should -Contain 'Logs\apppackager.log'
    }

    It 'leaves the workbench data root byte for byte untouched' {
        $after = @(Get-ChildItem -LiteralPath $script:DataRoot -Recurse -File |
            ForEach-Object { '{0}|{1}' -f $_.FullName.Substring($script:DataRoot.Length), (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } | Sort-Object)
        $after | Should -Be $script:DataBefore
        $script:DataRoot.StartsWith($script:InstallPath.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase) | Should -BeFalse
    }

    It 'does not preserve anything from outside the install folder' {
        @($script:Preserved | Where-Object { $_ -match '(?i)AppPackagerData' }).Count | Should -Be 0
    }

    It 'preserves a user-authored packager inside Packagers' {
        Test-Path -LiteralPath (Join-Path $script:InstallPath 'Packagers\package-contoso-inhouse.ps1') | Should -BeTrue
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\package-contoso-inhouse.ps1') -Raw | Should -Match 'user authored packager'
        $script:Preserved | Should -Contain 'Packagers\package-contoso-inhouse.ps1'
    }

    It 'preserves a disabled user-authored packager' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\package-contoso-draft.notps1') -Raw | Should -Match 'disabled user packager'
    }

    It 'lets the new release replace a shipped packager present in both trees' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\package-7zip.ps1') -Raw | Should -Match 'shipped packager v2'
        $script:Preserved | Should -Not -Contain 'Packagers\package-7zip.ps1'
    }

    It 'drops a packager the new release lists as retired' {
        Test-Path -LiteralPath (Join-Path $script:InstallPath 'Packagers\package-specexec-mitigations.ps1') | Should -BeFalse
        $script:Preserved | Should -Not -Contain 'Packagers\package-specexec-mitigations.ps1'
    }

    It 'keeps a template the release does not ship' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\Templates\contoso-options.ps1') -Raw | Should -Match 'user authored template'
        $script:Preserved | Should -Contain 'Packagers\Templates\contoso-options.ps1'
    }

    It 'keeps icon-pack content the release does not ship' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\Icons\contoso.ico') -Raw | Should -Match 'icon bytes'
        $script:Preserved | Should -Contain 'Packagers\Icons\contoso.ico'
    }

    It 'lets the new release replace a shipped template' {
        Get-Content -LiteralPath (Join-Path $script:InstallPath 'Packagers\Templates\options-panel-template.ps1') -Raw | Should -Match 'shipped template v2'
        $script:Preserved | Should -Not -Contain 'Packagers\Templates\options-panel-template.ps1'
    }

    It 'keeps no unshipped file beyond the state files when the release tree is unknown' {
        # Without the new release there is no way to tell a user-authored file
        # from an older copy of a shipped one, so none is carried.
        $withoutStage = @(Get-PreservedStateFile -Root $script:InstallPath)
        @($withoutStage | Where-Object { $_ -like 'Packagers\package-*' }).Count | Should -Be 0
        @($withoutStage | Where-Object { $_ -like 'Packagers\Templates\*' }).Count | Should -Be 0
        @($withoutStage | Where-Object { $_ -like 'Packagers\Icons\*' }).Count | Should -Be 0
        $withoutStage | Should -Contain 'AppPackager.preferences.json'
    }
}

Describe 'Preserved state selection' {
    BeforeAll {
        $script:Root = Join-Path $TestDrive 'selection'
        New-Item -ItemType Directory -Path (Join-Path $script:Root 'Logs\nested') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:Root 'Packagers') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Root 'AppPackager.preferences.json') -Value '{}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:Root 'Packagers\winrar.json') -Value '{}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:Root 'Logs\nested\old.log') -Value 'x' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:Root 'Packagers\package-x.ps1') -Value '# x' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:Root 'README.md') -Value 'x' -Encoding ASCII
        $script:Selected = @(Get-PreservedStateFile -Root $script:Root)
    }

    It 'keeps every json anywhere under the install folder' {
        $script:Selected | Should -Contain 'AppPackager.preferences.json'
        $script:Selected | Should -Contain 'Packagers\winrar.json'
    }

    It 'keeps every file under Logs, including nested ones' {
        $script:Selected | Should -Contain 'Logs\nested\old.log'
    }

    It 'keeps no shipped file and no script' {
        $script:Selected | Should -Not -Contain 'README.md'
        $script:Selected | Should -Not -Contain 'Packagers\package-x.ps1'
    }

    It 'returns an empty list for a folder that does not exist' {
        @(Get-PreservedStateFile -Root (Join-Path $TestDrive 'no-such-folder')).Count | Should -Be 0
    }

    It 'lists each relative path once' {
        $script:Selected.Count | Should -Be (@($script:Selected | Sort-Object -Unique).Count)
    }
}
