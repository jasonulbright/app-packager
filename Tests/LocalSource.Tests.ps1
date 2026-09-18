BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PackagerMetadata' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Hybrid packager header tag' {
    It 'reads LocalSource: Required' {
        $path = Join-Path $TestDrive 'package-hybrid.ps1'
        Set-Content -LiteralPath $path -Encoding ASCII -Value @'
<#
Vendor: Contoso
App: Widget
LocalSource: Required
#>
'@
        (Get-PackagerMetadata -Path $path).LocalSource | Should -BeTrue
        (Get-PackagerMetadata -Path $path).LocalSourceRequired | Should -BeTrue
    }

    It 'reads LocalSource: Optional as hybrid without the Stage prompt' {
        $path = Join-Path $TestDrive 'package-optional.ps1'
        Set-Content -LiteralPath $path -Encoding ASCII -Value @'
<#
Vendor: Contoso
App: Widget
LocalSource: Optional
#>
'@
        $meta = Get-PackagerMetadata -Path $path
        $meta.LocalSource | Should -BeTrue
        $meta.LocalSourceRequired | Should -BeFalse
    }

    It 'reads a packager without the tag as not hybrid' {
        $path = Join-Path $TestDrive 'package-plain.ps1'
        Set-Content -LiteralPath $path -Encoding ASCII -Value @'
<#
Vendor: Contoso
App: Widget
#>
'@
        (Get-PackagerMetadata -Path $path).LocalSource | Should -BeFalse
    }
}

Describe 'Resolve-LocalSourceInstaller' {
    BeforeEach {
        Mock Get-PackagerPreferences { $script:StoredPreferences } -ModuleName AppPackagerCommon
        $script:StoredPreferences = $null
    }

    It 'refuses when no folder is set' {
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.exe' } | Should -Throw '*No installer source folder is set*'
    }

    It 'refuses a folder that does not exist' {
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.exe' -Override (Join-Path $TestDrive 'missing') } |
            Should -Throw '*folder for package-widget not found*'
    }

    It 'refuses a folder with no matching installer' {
        $folder = Join-Path $TestDrive 'empty'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'readme.txt') -Value 'x'
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.exe' -Override $folder } |
            Should -Throw "*No installer matching '*.exe'*"
    }

    It 'breaks a file-version tie by write time' {
        $folder = Join-Path $TestDrive 'sources'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $old = Join-Path $folder 'Widget-1.0.exe'
        $new = Join-Path $folder 'Widget-1.1.exe'
        Set-Content -LiteralPath $old, $new -Value 'x'
        (Get-Item -LiteralPath $old).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddDays(-2)
        $script:StoredPreferences = [pscustomobject]@{ LocalSourceFolders = [pscustomobject]@{ 'package-widget' = $folder } }
        Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget-*.exe' | Should -Be $new
    }

    It 'lets an override path win over the saved folder' {
        $saved = Join-Path $TestDrive 'saved'
        $override = Join-Path $TestDrive 'override'
        New-Item -ItemType Directory -Path $saved, $override -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $saved 'Widget.exe'), (Join-Path $override 'Widget.exe') -Value 'x'
        $script:StoredPreferences = [pscustomobject]@{ LocalSourceFolders = [pscustomobject]@{ 'package-widget' = $saved } }
        Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget.exe' -Override $override |
            Should -Be (Join-Path $override 'Widget.exe')
    }

    It 'prefers the higher file version over the newer write time' {
        $folder = Join-Path $TestDrive 'versions'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $sources = @(Get-ChildItem -LiteralPath "$env:WINDIR\System32" -Filter '*.exe' -File |
            Where-Object { $v = $null; [version]::TryParse(([string]$_.VersionInfo.FileVersion -split ' ')[0], [ref]$v) } |
            Sort-Object { [version](([string]$_.VersionInfo.FileVersion -split ' ')[0]) } -Unique)
        $low = $sources[0]; $high = $sources[-1]
        [version](($low.VersionInfo.FileVersion -split ' ')[0]) | Should -BeLessThan ([version](($high.VersionInfo.FileVersion -split ' ')[0]))
        Copy-Item -LiteralPath $high.FullName -Destination (Join-Path $folder 'Widget-a.exe')
        Copy-Item -LiteralPath $low.FullName -Destination (Join-Path $folder 'Widget-b.exe')
        (Get-Item -LiteralPath (Join-Path $folder 'Widget-a.exe')).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddDays(-30)
        (Get-Item -LiteralPath (Join-Path $folder 'Widget-b.exe')).LastWriteTimeUtc = (Get-Date).ToUniversalTime()
        Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget-*.exe' -Override $folder |
            Should -Be (Join-Path $folder 'Widget-a.exe')
    }

    It 'skips files that match an exclude pattern' {
        $folder = Join-Path $TestDrive 'exclude'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'Widget.exe'), (Join-Path $folder 'Widget_x64.exe') -Value 'x'
        (Get-Item -LiteralPath (Join-Path $folder 'Widget.exe')).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddDays(-2)
        Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget*.exe' -Exclude @('*_x64*') -Override $folder |
            Should -Be (Join-Path $folder 'Widget.exe')
    }

    It 'refuses when every match is excluded' {
        $folder = Join-Path $TestDrive 'allexcluded'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'Widget_x64.exe') -Value 'x'
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget*.exe' -Exclude @('*_x64*') -Override $folder } |
            Should -Throw "*No installer matching*"
    }
}
