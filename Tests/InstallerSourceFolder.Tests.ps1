#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
    TeamViewer Host and TortoiseSVN have no download a script can use. Both stage the MSI that the operator puts in
    the installer source folder: the version is the ProductVersion of the file, no web request is sent, a missing file
    is refused with a message that names the folder and the file to put there, and Check Latest says that it has no
    web source. The tests build small real MSI files through the Windows Installer COM object.
#>
BeforeDiscovery {
    $script:Packagers = @(
        @{ Name = 'package-teamviewerhost'; Filter = 'TeamViewer_Host*.msi'; Match = @('TeamViewer_Host.msi', 'TeamViewer_Host (1).msi'); Older = 'TeamViewer_Host (1).msi'; Hint = 'Put the TeamViewer Host MSI (TeamViewer_Host.msi) in that folder.'; NoMatch = @('TeamViewer_Host_Setup_x64.exe', 'TeamViewer_Host.zip') }
        @{ Name = 'package-tortoisesvn'; Filter = 'TortoiseSVN-*-x64-*.msi'; Match = @('TortoiseSVN-1.14.9.29743-x64-svn-1.14.5.msi'); Older = 'TortoiseSVN-1.14.8.1000-x64-svn-1.14.5.msi'; Hint = 'Download the 64-bit MSI from tortoisesvn.net by hand'; NoMatch = @('TortoiseSVN-1.14.9.29743-win32-svn-1.14.5.msi', 'TortoiseSVN-1.14.9.29743-x64-svn-1.14.5.zip') }
    )
}

BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
    $script:PackagersRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\Packagers')).Path
    $script:ProductCode = '{8A6F0B37-3C1D-4E5A-9B2E-7D4C1F0A5E21}'

    function Get-FunctionText {
        param([string]$Path, [string]$Name)
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        return $fn.Extent.Text
    }
    function Get-ScriptAssignment {
        param([string]$Path, [string]$Variable)
        $t = $null; $e = $null
        $tree = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$t, [ref]$e)
        $node = $tree.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$Variable" }, $false)
        return $node.Right.Extent.Text.Trim('"', "'")
    }
    function New-TestMsi {
        param([string]$Path, [string]$Version, [string]$Name = 'Widget')
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 3))
        $run = {
            param($sql)
            $view = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db, @($sql))
            $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
            $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null) | Out-Null
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($view) | Out-Null
        }
        & $run 'CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)'
        $values = [ordered]@{ ProductName = $Name; ProductVersion = $Version; Manufacturer = 'Contoso'; ProductCode = $script:ProductCode }
        foreach ($key in $values.Keys) {
            & $run ('INSERT INTO `Property` (`Property`, `Value`) VALUES (''{0}'', ''{1}'')' -f $key, $values[$key])
        }
        $db.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $db, $null) | Out-Null
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    function Invoke-Packager {
        # Runs a packager in its own Windows PowerShell and returns the exit code and the text it wrote.
        param([string]$Name, [string[]]$Arguments)
        $log = Join-Path $TestDrive ('{0}-{1}.log' -f $Name, [guid]::NewGuid().ToString('N'))
        $text = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $script:PackagersRoot "$Name.ps1") @Arguments -LogPath $log 2>&1 | Out-String
        return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Text = $text }
    }
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PackagerMetadata' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'The packager <Name> stages from the installer source folder' -ForEach $script:Packagers {
    BeforeAll { $script:Path = Join-Path $script:PackagersRoot "$Name.ps1" }

    It 'declares the Required local source in its header' {
        $meta = Get-PackagerMetadata -Path $script:Path
        $meta.LocalSource | Should -BeTrue
        $meta.LocalSourceRequired | Should -BeTrue
    }

    It 'has the file name rule <Filter>' {
        Get-ScriptAssignment -Path $script:Path -Variable 'InstallerFilter' | Should -Be $Filter
    }

    It 'matches the vendor file names' {
        foreach ($file in $Match) { $file | Should -BeLike $Filter }
        foreach ($file in $NoMatch) { $file | Should -Not -BeLike $Filter }
    }

    It 'sends no web request and does not touch the protected page' {
        $code = ((Get-Content -LiteralPath $script:Path -Raw) -replace '(?s)^\s*<#.*?\n#>', '')
        $code | Should -Not -Match 'curl\.exe|Invoke-WebRequest|Invoke-RestMethod|Invoke-DownloadWithRetry|Invoke-CachedDownload|Net\.WebClient|HttpClient'
        $code | Should -Not -Match 'sourceforge|tortoisesvn\.net/downloads|download\.teamviewer|dl\.teamviewer'
    }

    It 'names the folder and the file to put there when the folder is empty' {
        $folder = Join-Path $TestDrive "empty-$Name"
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $run = Invoke-Packager -Name $Name -Arguments @('-GetLatestVersionOnly', '-SourceFolder', $folder)
        $run.ExitCode | Should -Be 1
        $run.Text | Should -BeLike '*has no web version source*'
        $run.Text | Should -BeLike "*No installer matching '$Filter' found in $folder*"
        $run.Text | Should -BeLike "*$Hint*"
    }

    It 'names the missing folder' {
        $folder = Join-Path $TestDrive "missing-$Name"
        $run = Invoke-Packager -Name $Name -Arguments @('-GetLatestVersionOnly', '-SourceFolder', $folder)
        $run.ExitCode | Should -Be 1
        $run.Text | Should -BeLike "*not found: $folder*"
    }

    It 'prints the ProductVersion of the newest MSI as the latest version' {
        $folder = Join-Path $TestDrive "latest-$Name"
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        New-TestMsi -Path (Join-Path $folder $Older) -Version '1.14.8'
        New-TestMsi -Path (Join-Path $folder $Match[0]) -Version '1.14.29743'
        $run = Invoke-Packager -Name $Name -Arguments @('-GetLatestVersionOnly', '-SourceFolder', $folder)
        $run.ExitCode | Should -Be 0
        @($run.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }) | Should -Contain '1.14.29743'
    }
}

Describe 'TortoiseSVN stage' {
    It 'takes the version, the product code and the detection from the MSI' {
        $folder = Join-Path $TestDrive 'stage-source'
        $root = Join-Path $TestDrive 'stage-root'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        New-TestMsi -Path (Join-Path $folder 'TortoiseSVN-1.14.9.29743-x64-svn-1.14.5.msi') -Version '1.14.29743' -Name 'TortoiseSVN 1.14.9.29743 (64 bit)'

        $run = Invoke-Packager -Name 'package-tortoisesvn' -Arguments @('-StageOnly', '-SourceFolder', $folder, '-DownloadRoot', $root)

        $run.ExitCode | Should -Be 0 -Because $run.Text
        $content = Join-Path $root 'TortoiseSVN\1.14.29743'
        $manifest = Get-Content -LiteralPath (Join-Path $content 'stage-manifest.json') -Raw | ConvertFrom-Json
        $manifest.SoftwareVersion | Should -Be '1.14.29743'
        $manifest.InstallerType | Should -Be 'MSI'
        $manifest.ProductCode | Should -Be $script:ProductCode
        $manifest.Detection.RegistryKeyRelative | Should -BeLike "*$($script:ProductCode)"
        $manifest.Detection.ExpectedValue | Should -Be '1.14.29743'
        Test-Path -LiteralPath (Join-Path $content 'install.ps1') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $content 'TortoiseSVN-1.14.9.29743-x64-svn-1.14.5.msi') | Should -BeTrue
        $run.Text | Should -BeLike '*signature is*' -Because 'an unsigned file is logged'
    }
}

Describe 'TeamViewer Host stage' {
    It 'refuses an MSI without a valid signature' {
        $folder = Join-Path $TestDrive 'tv-unsigned'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        New-TestMsi -Path (Join-Path $folder 'TeamViewer_Host.msi') -Version '15.62.3' -Name 'TeamViewer Host'

        $run = Invoke-Packager -Name 'package-teamviewerhost' -Arguments @('-StageOnly', '-SourceFolder', $folder, '-DownloadRoot', (Join-Path $TestDrive 'tv-root'))

        $run.ExitCode | Should -Be 1
        $run.Text | Should -BeLike '*Installer signature is*not Valid*'
    }
}

Describe 'Get-TeamViewerHostMsiArguments' {
    BeforeAll {
        . ([scriptblock]::Create((Get-FunctionText -Path (Join-Path $script:PackagersRoot 'package-teamviewerhost.ps1') -Name 'Get-TeamViewerHostMsiArguments')))
    }

    It 'adds nothing without a configuration' {
        @(Get-TeamViewerHostMsiArguments -Config $null).Count | Should -Be 0
    }

    It 'adds nothing for empty values' {
        $config = [pscustomobject]@{ ApiToken = ''; CustomConfigId = ' '; AssignmentOptions = $null; RemoveDesktopShortcut = $false }
        @(Get-TeamViewerHostMsiArguments -Config $config).Count | Should -Be 0
    }

    It 'turns the four keys into MSI properties' {
        $config = [pscustomobject]@{ ApiToken = 'tok-123'; CustomConfigId = 'cfg456'; AssignmentOptions = '--grant-easy-access --reassign'; RemoveDesktopShortcut = $true }
        $arguments = @(Get-TeamViewerHostMsiArguments -Config $config)
        $arguments | Should -Contain 'APITOKEN=tok-123'
        $arguments | Should -Contain 'CUSTOMCONFIGID=cfg456'
        $arguments | Should -Contain 'ASSIGNMENTOPTIONS="--grant-easy-access --reassign"'
        $arguments | Should -Contain 'DESKTOPSHORTCUTS=0'
        $arguments | Should -Contain 'ADDLOCAL=ALL'
        $arguments | Should -Contain 'REMOVE=f.DesktopShortcut'
    }

    It 'refuses a space in the token or the config id, and a trailing backslash in the options' {
        { Get-TeamViewerHostMsiArguments -Config ([pscustomobject]@{ ApiToken = 'a b' }) } | Should -Throw '*ApiToken*'
        { Get-TeamViewerHostMsiArguments -Config ([pscustomobject]@{ CustomConfigId = 'a b' }) } | Should -Throw '*CustomConfigId*'
        { Get-TeamViewerHostMsiArguments -Config ([pscustomobject]@{ AssignmentOptions = '--x\' }) } | Should -Throw '*backslash*'
    }

    It 'reads the desktop shortcut flag from a text value too' {
        @(Get-TeamViewerHostMsiArguments -Config ([pscustomobject]@{ RemoveDesktopShortcut = 'true' })) | Should -Contain 'DESKTOPSHORTCUTS=0'
    }

    It 'refuses a value with a double quote' {
        { Get-TeamViewerHostMsiArguments -Config ([pscustomobject]@{ ApiToken = 'a"b' }) } | Should -Throw '*double quote*'
    }

    It 'reaches the install wrapper as one argument per property' {
        $config = [pscustomobject]@{ ApiToken = 'tok'; AssignmentOptions = '--reassign' }
        $wrapper = New-MsiWrapperContent -MsiFileName 'TeamViewer_Host.msi' -ExtraInstallArgs @(Get-TeamViewerHostMsiArguments -Config $config)
        $wrapper.Install | Should -BeLike "*`$args += 'APITOKEN=tok'*"
        $wrapper.Install | Should -BeLike '*$args += ''ASSIGNMENTOPTIONS="--reassign"''*'
    }
}

Describe 'Resolve-LocalSourceInstaller with a hint' {
    BeforeEach {
        Mock Get-PackagerPreferences { $null } -ModuleName AppPackagerCommon
    }

    It 'ends every refusal with the hint' {
        $hint = 'Put the file X in that folder.'
        $empty = Join-Path $TestDrive 'hint-empty'
        New-Item -ItemType Directory -Path $empty -Force | Out-Null
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.msi' -Hint $hint } | Should -Throw "*$hint"
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.msi' -Override (Join-Path $TestDrive 'hint-missing') -Hint $hint } | Should -Throw "*not found*$hint"
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.msi' -Override $empty -Hint $hint } | Should -Throw "*No installer matching '*.msi' found in $empty.*$hint"
    }

    It 'names a .zip in the folder and does not read it' {
        $folder = Join-Path $TestDrive 'hint-zip'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'Widget.zip') -Value 'x'
        { Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter '*.msi' -Override $folder } | Should -Throw '*holds a .zip file*Unzip it*'
    }

    It 'ranks an MSI by its ProductVersion, not by its write time' {
        $folder = Join-Path $TestDrive 'hint-rank'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $older = Join-Path $folder 'Widget-a.msi'
        $newer = Join-Path $folder 'Widget-b.msi'
        New-TestMsi -Path $older -Version '2.0.1'
        New-TestMsi -Path $newer -Version '2.0.10'
        (Get-Item -LiteralPath $newer).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddDays(-30)
        (Get-Item -LiteralPath $older).LastWriteTimeUtc = (Get-Date).ToUniversalTime()
        Resolve-LocalSourceInstaller -PackagerName 'package-widget' -Filter 'Widget-*.msi' -Override $folder | Should -Be $newer
    }
}

Describe 'Get-InstallerFileVersion' {
    It 'reads the ProductVersion of an MSI file' {
        $path = Join-Path $TestDrive 'version.msi'
        New-TestMsi -Path $path -Version '3.4.5'
        Get-InstallerFileVersion -Path $path | Should -Be '3.4.5'
    }

    It 'gives an empty text for a file that is no MSI' {
        $path = Join-Path $TestDrive 'broken.msi'
        Set-Content -LiteralPath $path -Value 'not an msi'
        Get-InstallerFileVersion -Path $path | Should -Be ''
    }

    It 'reads the FileVersion of an EXE file' {
        (Get-InstallerFileVersion -Path (Join-Path $env:SystemRoot 'System32\notepad.exe')) | Should -Match '^\d+(\.\d+)+'
    }
}
