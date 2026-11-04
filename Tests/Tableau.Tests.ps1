Describe '<Name>' -ForEach @(
        @{ Name = 'package-tableaureader.ps1';  Filter = 'TableauReader-64bit-*.exe';  Pattern = 'Tableau Reader*'; Sample = 'TableauReader-64bit-2026-2-2.exe';  RemoveInstalled = $false ; Url = 'https://downloads.tableau.com/tssoftware/TableauReader-64bit-2026-2-2.exe' }
        @{ Name = 'package-tableauprep.ps1';    Filter = 'TableauPrep-*.exe';          Pattern = 'Tableau Prep*';   Sample = 'TableauPrep-2026-2-2.exe';          RemoveInstalled = $false ; Url = 'https://downloads.tableau.com/tssoftware/TableauPrep-2026-2-2.exe' }
        @{ Name = 'package-tableaudesktop.ps1'; Filter = 'TableauDesktop-64bit-*.exe'; Pattern = 'Tableau 20*';     Sample = 'TableauDesktop-64bit-2026-2-2.exe'; RemoveInstalled = $true ; Url = 'https://downloads.tableau.com/esdalt/2026.2.2/TableauDesktop-64bit-2026-2-2.exe' }
    ) {
    BeforeAll {
        $path = Join-Path $PSScriptRoot "..\Packagers\$Name"
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }

        foreach ($assign in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Parent -is [System.Management.Automation.Language.NamedBlockAst] }, $false)) {
            $target = $assign.Left.Extent.Text
            if ($target -in @('$InstallerFilter', '$DisplayNamePattern', '$InstallSwitches', '$DownloadUrlTemplate')) {
                . ([scriptblock]::Create($assign.Extent.Text))
            }
        }
        foreach ($fnName in @('Get-TableauVersionFromFileName', 'Get-TableauVersionFromPage', 'New-TableauBundleSearchText', 'New-TableauDetectionScript', 'New-TableauUninstallScript')) {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $fnName }, $false)
            . ([scriptblock]::Create($fn.Extent.Text))
        }
        $headerText = (Get-Content -LiteralPath $path -TotalCount 20) -join "`n"
    }

    It 'reads the highest release from the release page' {
        Get-TableauVersionFromPage -Html '<a href="/desktop/2026.1.3">2026.1.3</a> <a href="/desktop/2026.2">2026.2</a> <a href="/desktop/2026.2.2">2026.2.2</a>' | Should -Be '2026.2.2'
        Get-TableauVersionFromPage -Html '<a>2026.3</a> <a>2026.2.9</a>' | Should -Be '2026.3.0'
        Get-TableauVersionFromPage -Html 'no releases' | Should -BeNullOrEmpty
    }

    It 'builds the download URL from the release' {
        ($DownloadUrlTemplate -f '2026-2-2', '2026.2.2') | Should -Be $Url
    }

    It 'downloads on its own rather than asking for a local source' {
        $headerText | Should -Not -Match '(?m)^LocalSource:'
    }

    It 'matches the expected installer names and display name' {
        $InstallerFilter | Should -Be $Filter
        $DisplayNamePattern | Should -Be $Pattern
        $Sample | Should -BeLike $Filter
    }

    It 'reads the release version from the installer name' {
        Get-TableauVersionFromFileName -FileName $Sample | Should -Be '2026.2.2'
        Get-TableauVersionFromFileName -FileName 'setup.exe' | Should -BeNullOrEmpty
    }

    It 'installs with the production switches' {
        @($InstallSwitches)[0..4] | Should -Be @('/install', '/quiet', '/norestart', 'ACCEPTEULA=1', $(if ($RemoveInstalled) { 'REMOVEINSTALLEDAPP=1' } else { 'SENDTELEMETRY=0' }))
        ($InstallSwitches -contains 'REMOVEINSTALLEDAPP=1') | Should -Be $RemoveInstalled
    }

    It 'detects through the 64-bit bundle entry with the installer version as the minimum' {
        $script = New-TableauDetectionScript -Pattern $Pattern -MinimumVersion '26.2.1782'
        $script | Should -Match 'RegistryView\]::Registry64'
        $script | Should -Not -Match 'Registry32'
        $script | Should -Match ([regex]::Escape("-like '$Pattern'"))
        $script | Should -Match ([regex]::Escape("[version]'26.2.1782'"))
    }

    It 'uninstalls by running the QuietUninstallString as written, never the MSI' {
        $script = New-TableauUninstallScript -Pattern $Pattern
        $script | Should -Match 'RegistryView\]::Registry64'
        $script | Should -Not -Match 'Registry32'
        $script | Should -Match 'QuietUninstallString'
        $script | Should -Match ([regex]::Escape("-ArgumentList `$m.Groups['args'].Value -Wait"))
        $script | Should -Not -Match '/norestart'
        $script | Should -Not -Match '(?i)msiexec'
    }

    It 'generates an uninstall script that parses' {
        $script = New-TableauUninstallScript -Pattern $Pattern
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($script, [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'launches the executable and the arguments of the QuietUninstallString' {
        $exe = Join-Path $TestDrive 'bundle setup.exe'
        Set-Content -LiteralPath $exe -Value 'x'
        $search = New-TableauBundleSearchText -Pattern $Pattern
        $body = (New-TableauUninstallScript -Pattern $Pattern).Substring($search.Length) -replace 'exit \$exitCode\s*$', '$exitCode'
        $stub = '$bundles = @([pscustomobject]@{ DisplayName = ''x''; DisplayVersion = ''1.0''; QuietUninstall = ''"' + $exe + '" /burn.clean.room /uninstall /quiet'' })'
        $global:TableauLaunch = $null
        function global:Start-Process {
            param($FilePath, $ArgumentList, [switch]$Wait, [switch]$PassThru, [switch]$NoNewWindow)
            $global:TableauLaunch = @{ FilePath = $FilePath; ArgumentList = $ArgumentList; Wait = [bool]$Wait }
            [pscustomobject]@{ ExitCode = 0 }
        }
        try { $result = & ([scriptblock]::Create($stub + "`r`n" + $body)) }
        finally { Remove-Item -LiteralPath Function:\global:Start-Process -ErrorAction SilentlyContinue }
        $result | Should -Be 0
        $global:TableauLaunch.FilePath | Should -Be $exe
        $global:TableauLaunch.ArgumentList | Should -Be '/burn.clean.room /uninstall /quiet'
        $global:TableauLaunch.Wait | Should -BeTrue
    }
}
