#Requires -Version 5.1
BeforeAll {
    $root = Split-Path $PSScriptRoot -Parent
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-WorkbenchInheritedIconPath' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Get-WorkbenchInheritedIconPath' {
    BeforeEach {
        $script:packDir = Join-Path $TestDrive 'Icons'
        $script:stageDir = Join-Path $TestDrive 'stage\7-Zip\26.03'
        $script:iconSource = 'Installer'
        $script:manifestPath = $null
        $script:downloadSubfolder = '7-Zip'
        $script:searchedDownloadRoot = $false
        New-Item -ItemType Directory -Path $script:packDir, $script:stageDir -Force | Out-Null
        function Get-IconPackRoot { $script:packDir }
        function Get-PackagerIconSource { param($ScriptPath) $script:iconSource }
        function Get-PackagerFolderInfo { param($ScriptPath) @{ DownloadSubfolder = $script:downloadSubfolder; VendorFolder = $null; AppFolder = $null } }
        function Find-NewestStageManifestForPackager { param($PackagerPath, $DownloadRoot) $script:searchedDownloadRoot = $true; $script:manifestPath }
    }
    AfterEach {
        Remove-Item -LiteralPath (Join-Path $TestDrive 'Icons'), (Join-Path $TestDrive 'stage') -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'returns the icon pack entry before the staged app-icon of the newest build' {
        Set-Content -LiteralPath (Join-Path $script:packDir '7zip.png') -Value 'x'
        Set-Content -LiteralPath (Join-Path $script:stageDir 'app-icon.ico') -Value 'x'
        $script:manifestPath = Join-Path $script:stageDir 'stage-manifest.json'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot 'C:\temp\ap' |
            Should -Be (Join-Path $script:packDir '7zip.png')
    }

    It 'falls back to the staged app-icon when the pack has no entry' {
        Set-Content -LiteralPath (Join-Path $script:stageDir 'app-icon.ico') -Value 'x'
        $script:manifestPath = Join-Path $script:stageDir 'stage-manifest.json'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot 'C:\temp\ap' |
            Should -Be (Join-Path $script:stageDir 'app-icon.ico')
    }

    It 'returns the icon pack entry when nothing is staged' {
        Set-Content -LiteralPath (Join-Path $script:packDir '7zip.png') -Value 'x'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot '' |
            Should -Be (Join-Path $script:packDir '7zip.png')
    }

    It 'does not match a longer packager name that shares the prefix' {
        Set-Content -LiteralPath (Join-Path $script:packDir 'chromeremotedesktophost.png') -Value 'x'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-chrome.ps1' -DownloadRoot '' | Should -Be ''
    }

    It 'returns the icon pack entry for a packager tagged None' {
        Set-Content -LiteralPath (Join-Path $script:packDir '7zip.png') -Value 'x'
        $script:iconSource = 'None'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot '' |
            Should -Be (Join-Path $script:packDir '7zip.png')
    }

    It 'returns nothing for a packager tagged None without a pack entry' {
        Set-Content -LiteralPath (Join-Path $script:stageDir 'app-icon.ico') -Value 'x'
        $script:manifestPath = Join-Path $script:stageDir 'stage-manifest.json'
        $script:iconSource = 'None'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot 'C:\temp\ap' | Should -Be ''
    }

    It 'skips the stage manifest search when the download subfolder is unknown' {
        $script:downloadSubfolder = $null
        $script:manifestPath = Join-Path $script:stageDir 'stage-manifest.json'
        Set-Content -LiteralPath (Join-Path $script:stageDir 'app-icon.ico') -Value 'x'
        Get-WorkbenchInheritedIconPath -ScriptPath 'C:\x\package-7zip.ps1' -DownloadRoot 'C:\temp\ap' | Should -Be ''
        $script:searchedDownloadRoot | Should -BeFalse
    }

    It 'returns nothing when the application has no packager script' {
        Get-WorkbenchInheritedIconPath -ScriptPath '' -DownloadRoot '' | Should -Be ''
    }
}
