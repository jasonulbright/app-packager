#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    What the first open of the Options window costs.

.DESCRIPTION
    Building all seven panels before the window appears, waiting on a web
    request to api.github.com in the ConfigMgr panel, and parsing the header of
    all 300+ packagers in three panels each delay the first open. These cases
    hold the alternatives in place: the catalog is parsed once per change, a
    panel is built when its section is first shown, and the GitHub status fills
    in from a background runspace.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Get-PackagerMetadata', 'Get-Packagers') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function New-PackagerFile {
        param([string]$Folder, [string]$Name, [string]$App)
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $Folder $Name) -Encoding ASCII -Value @(
            '<#', 'Vendor: Contoso', "App: $App", "CMName: $App", 'UpdateCadenceDays: 14', '', '.SYNOPSIS', "    Packages $App.", '#>')
    }
}

Describe 'Packager catalog' {
    It 'parses the header fields of a packager' {
        $folder = Join-Path $TestDrive 'meta'
        New-PackagerFile -Folder $folder -Name 'package-one.ps1' -App 'One'
        $m = Get-PackagerMetadata -Path (Join-Path $folder 'package-one.ps1')
        $m.Vendor | Should -Be 'Contoso'
        $m.Application | Should -Be 'One'
        $m.UpdateCadenceDays | Should -Be 14
        $m.Description | Should -Be 'Packages One.'
    }

    It 'reads header keys in any letter case' {
        $folder = Join-Path $TestDrive 'case'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'package-case.ps1') -Encoding ASCII -Value @('<#', 'vendor: Contoso', 'APP: Lower', 'cmname: Lower Name', '#>')
        $m = Get-PackagerMetadata -Path (Join-Path $folder 'package-case.ps1')
        $m.Vendor | Should -Be 'Contoso'
        $m.Application | Should -Be 'Lower'
        $m.CMName | Should -Be 'Lower Name'
    }

    It 'reuses the parsed rows until a packager changes, and hands out copies' {
        $folder = Join-Path $TestDrive 'catalog'
        New-PackagerFile -Folder $folder -Name 'package-one.ps1' -App 'One'
        New-PackagerFile -Folder $folder -Name 'package-two.ps1' -App 'Two'

        $first = @(Get-Packagers -Root $folder)
        $first.Count | Should -Be 2
        $first[0].Selected = $true
        $first[0].Status = 'edited by the grid'

        $second = @(Get-Packagers -Root $folder)
        $second[0].Selected | Should -BeFalse
        $second[0].Status | Should -Be 'Ready'
        $second[0].Application | Should -Be $first[0].Application

        # Poison the cached parse: a reread would replace it.
        $script:PackagerCatalogCache[$folder.TrimEnd('\').ToLowerInvariant()].Rows[0].Application = 'from the cache'
        @(Get-Packagers -Root $folder)[0].Application | Should -Be 'from the cache'

        New-PackagerFile -Folder $folder -Name 'package-three.ps1' -App 'Three'
        $third = @(Get-Packagers -Root $folder)
        $third.Count | Should -Be 3
        ($third | Where-Object Script -eq 'package-one.ps1').Application | Should -Be 'One'
    }

    It 'rereads a packager whose file changed' {
        $folder = Join-Path $TestDrive 'changed'
        New-PackagerFile -Folder $folder -Name 'package-one.ps1' -App 'Before'
        @(Get-Packagers -Root $folder)[0].Application | Should -Be 'Before'
        Start-Sleep -Milliseconds 50
        New-PackagerFile -Folder $folder -Name 'package-one.ps1' -App 'After the edit'
        (Get-Item -LiteralPath (Join-Path $folder 'package-one.ps1')).LastWriteTimeUtc = [DateTime]::UtcNow.AddMinutes(1)
        @(Get-Packagers -Root $folder)[0].Application | Should -Be 'After the edit'
    }
}

Describe 'Options window' {
    BeforeAll {
        $script:DialogText = (Get-LauncherFunction 'New-OptionsDialog').Extent.Text
    }

    It 'builds the tree from the section definitions' {
        $script:DialogText | Should -Match 'Get-OptionsSectionDefinition'
        $script:DialogText | Should -Match 'New-Object System.Windows.Controls.TreeViewItem'
    }

    It 'builds a section the first time it is shown' {
        $script:DialogText | Should -Match 'if \(-not \$ctx\.Sections\.Contains\(\$Id\)\)'
        $script:DialogText | Should -Match 'New-OptionsSectionPage -Definition \$def -Context \$ctx'
    }

    It 'commits only the sections that were built' {
        $script:DialogText | Should -Match 'foreach \(\$id in @\(\$ctx\.Sections\.Keys\)\) \{ \$page = \$ctx\.Sections\[\$id\]; if \(\$page -and \$page\.Commit\)'
    }

    It 'does not wait for the GitHub status while the section is built' {
        $init = (Get-LauncherFunction 'Initialize-OptionsGitHubSection').Extent.Text
        $init | Should -Not -Match 'Get-GitHubApiAuthStatus'
        $check = (Get-LauncherFunction 'Start-OptionsGitHubCheck').Extent.Text
        $withoutBackgroundScript = $check -replace '(?s)\[void\]\$shell\.AddScript\(.*?\)\.AddArgument\(\$state\)\.AddArgument\(\(Join-Path \$PSScriptRoot[^
]*', ''
        $withoutBackgroundScript | Should -Not -Match 'Get-GitHubApiAuthStatus'
        $check | Should -Match 'Checking GitHub access'
        $check | Should -Match 'DispatcherTimer'
    }
}
