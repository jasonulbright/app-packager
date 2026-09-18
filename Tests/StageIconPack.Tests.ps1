BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force

    function New-TaggedPackager {
        param([string]$Name, [string]$Tag)
        $path = Join-Path $TestDrive "package-$Name.ps1"
        $lines = @('<#', 'Vendor: Contoso', "App: $Name")
        if ($Tag) { $lines += "IconSource: $Tag" }
        $lines += '#>'
        Set-Content -LiteralPath $path -Value $lines -Encoding ASCII
        $path
    }
}

Describe 'Add-StageIcon icon pack order' {
    BeforeEach {
        $script:Pack = Join-Path $TestDrive 'Icons'
        $script:Stage = Join-Path $TestDrive 'stage'
        Remove-Item -LiteralPath $script:Pack, $script:Stage -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $script:Pack, $script:Stage -Force | Out-Null
        Mock Get-InstallerIcon -ModuleName AppPackagerCommon {
            Set-Content -LiteralPath $OutputPath -Value 'extracted'
            [pscustomobject]@{ Path = $OutputPath }
        }
    }

    It 'uses the pack icon for <Tag> when the pack has an entry' -ForEach @(
        @{ Tag = 'Installer' }, @{ Tag = 'External' }, @{ Tag = 'None' }, @{ Tag = '' }
    ) {
        $script_ = New-TaggedPackager -Name 'widget' -Tag $Tag
        Set-Content -LiteralPath (Join-Path $script:Pack 'widget.png') -Value 'pack'
        Set-Content -LiteralPath (Join-Path $script:Stage 'widget-setup.exe') -Value 'x'
        $data = @{ InstallerFile = 'widget-setup.exe' }
        Add-StageIcon -StageRoot $script:Stage -ManifestData $data -PackagerScriptPath $script_ -IconsDirectory $script:Pack
        $data['Icon'] | Should -Be 'app-icon.png'
        Get-Content -LiteralPath (Join-Path $script:Stage 'app-icon.png') | Should -Be 'pack'
        Should -Invoke Get-InstallerIcon -ModuleName AppPackagerCommon -Times 0
    }

    It 'extracts from the installer for Installer when the pack has no entry' {
        $script_ = New-TaggedPackager -Name 'widget' -Tag 'Installer'
        Set-Content -LiteralPath (Join-Path $script:Stage 'widget-setup.exe') -Value 'x'
        $data = @{ InstallerFile = 'widget-setup.exe' }
        Add-StageIcon -StageRoot $script:Stage -ManifestData $data -PackagerScriptPath $script_ -IconsDirectory $script:Pack
        $data['Icon'] | Should -Be 'app-icon.ico'
        Should -Invoke Get-InstallerIcon -ModuleName AppPackagerCommon -Times 1
    }

    It 'stages no icon for None when the pack has no entry' {
        $script_ = New-TaggedPackager -Name 'widget' -Tag 'None'
        $data = @{ InstallerFile = 'widget-setup.exe' }
        Add-StageIcon -StageRoot $script:Stage -ManifestData $data -PackagerScriptPath $script_ -IconsDirectory $script:Pack
        $data.ContainsKey('Icon') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:Stage -Filter 'app-icon.*').Count | Should -Be 0
    }

    It 'ignores a pack entry for a longer name that shares the prefix' {
        $script_ = New-TaggedPackager -Name 'chrome' -Tag 'None'
        Set-Content -LiteralPath (Join-Path $script:Pack 'chrome.remote.png') -Value 'pack'
        $data = @{}
        Add-StageIcon -StageRoot $script:Stage -ManifestData $data -PackagerScriptPath $script_ -IconsDirectory $script:Pack
        $data.ContainsKey('Icon') | Should -BeFalse
    }
}
