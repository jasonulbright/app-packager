BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force

    function Get-AuthoredDetection {
        param([string]$File, [string]$Version, [hashtable]$Variables = @{})
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            "$PSScriptRoot\..\Packagers\$File", [ref]$null, [ref]$null)
        $command = $ast.Find({ param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
                $node.GetCommandName() -eq 'Write-StageManifest'
        }, $true)
        $literal = $command.CommandElements | Where-Object {
            $_ -is [System.Management.Automation.Language.HashtableAst]
        }
        $version = $Version
        $installerFileName = 'test.exe'
        foreach ($name in $Variables.Keys) { Set-Variable -Name $name -Value $Variables[$name] }
        $manifest = [pscustomobject]((& ([scriptblock]::Create($literal.Extent.Text))) |
            ConvertTo-Json -Depth 10 | ConvertFrom-Json)
        return $manifest.Detection
    }

    function New-BurnStub {
        param([string]$Path, [guid]$BundleId, [string]$SectionName = '.wixburn')
        $stream = [System.IO.MemoryStream]::new()
        $writer = [System.IO.BinaryWriter]::new($stream)
        $dos = New-Object byte[] 0x40
        $dos[0] = 0x4D; $dos[1] = 0x5A; $dos[0x3C] = 0x40
        $writer.Write($dos)
        $writer.Write([uint32]0x00004550)
        $writer.Write([uint16]0x8664)
        $writer.Write([uint16]1)
        $writer.Write([uint32]0); $writer.Write([uint32]0); $writer.Write([uint32]0)
        $writer.Write([uint16]0)
        $writer.Write([uint16]0)
        $name = New-Object byte[] 8
        [System.Text.Encoding]::ASCII.GetBytes($SectionName).CopyTo($name, 0)
        $writer.Write($name)
        $writer.Write([uint32]0x30); $writer.Write([uint32]0x1000); $writer.Write([uint32]0x30); $writer.Write([uint32]0x80)
        $writer.Write([uint32]0); $writer.Write([uint32]0); $writer.Write([uint32]0); $writer.Write([uint32]0)
        while ($stream.Length -lt 0x80) { $writer.Write([byte]0) }
        $writer.Write([uint32]0x00F14300)
        $writer.Write([uint32]2)
        $writer.Write($BundleId.ToByteArray())
        $writer.Flush()
        [System.IO.File]::WriteAllBytes($Path, $stream.ToArray())
    }
}

Describe 'Version-bearing detection' {
    It 'compares the file version of <File> against the staged version' -ForEach @(
        @{ File = 'package-azurestorageexplorer.ps1'; Name = 'StorageExplorer.exe' }
        @{ File = 'package-rainmeter.ps1'; Name = 'Rainmeter.exe' }
        @{ File = 'package-pdfstudioviewer.ps1'; Name = $null }
        @{ File = 'package-cyberduck.ps1'; Name = 'Cyberduck.exe' }
        @{ File = 'package-daxstudio.ps1'; Name = 'DaxStudio.exe' }
        @{ File = 'package-greenshot.ps1'; Name = 'Greenshot.exe' }
        @{ File = 'package-pspad.ps1'; Name = 'PSPad.exe' }
    ) {
        $d = Get-AuthoredDetection -File $File -Version '4.5.28' -Variables @{ layout = [pscustomobject]@{ LauncherName = 'Launcher.exe'; InstallDir = 'C:\App'; ProcessName = 'App' } }
        $d.Type | Should -Be 'File'
        $d.PropertyType | Should -Be 'Version'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be '4.5.28'
        if ($Name) { $d.FileName | Should -Be $Name }
    }

    It 'drops the packaging suffix from the Positron file version' {
        $d = Get-AuthoredDetection -File 'package-positron.ps1' -Version '2026.10.1-2'
        $d.PropertyType | Should -Be 'Version'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be '2026.10.1'
    }

    It 'compares the uninstall entry version of <File>' -ForEach @(
        @{ File = 'package-aimp.ps1'; Key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AIMP'; Is64Bit = $true }
        @{ File = 'package-cpuz.ps1'; Key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CPUID CPU-Z_is1'; Is64Bit = $true }
        @{ File = 'package-hwmonitor.ps1'; Key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CPUID HWMonitor_is1'; Is64Bit = $true }
        @{ File = 'package-vim.ps1'; Key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Vim'; Is64Bit = $true }
    ) {
        $d = Get-AuthoredDetection -File $File -Version '9.2.1229' -Variables @{ arpRegistryKey = $Key }
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.RegistryKeyRelative | Should -Be $Key
        $d.ValueName | Should -Be 'DisplayVersion'
        $d.PropertyType | Should -Be 'Version'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be '9.2.1229'
        $d.Is64Bit | Should -Be $Is64Bit
    }

    It 'names the pgAdmin uninstall key after the major version' {
        $source = Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\package-pgadmin4.ps1"
        $source | Should -Match 'Uninstall\\pgAdmin 4v\{0\}_is1"\s*-f\s*\$version\.Split\(''\.''\)\[0\]'
        $d = Get-AuthoredDetection -File 'package-pgadmin4.ps1' -Version '9.18' -Variables @{ arpKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\pgAdmin 4v9_is1' }
        $d.PropertyType | Should -Be 'Version'
        $d.ExpectedValue | Should -Be '9.18'
        $d.Is64Bit | Should -BeTrue
    }

    It 'detects the Session Manager plugin by the release bundle key in the 32-bit view' {
        $d = Get-AuthoredDetection -File 'package-awsssmplugin.ps1' -Version '1.2.896.0' -Variables @{ arpKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{cef0d74c-33bc-432b-941f-6e17bea1f9d3}' }
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{cef0d74c-33bc-432b-941f-6e17bea1f9d3}'
        $d.PropertyType | Should -Be 'Version'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be '1.2.896.0'
        $d.Is64Bit | Should -BeFalse
    }

    It 'detects Sysinternals by the version the install writes' {
        $key = 'HKLM:\SOFTWARE\App Packager\Sysinternals Suite'
        $d = Get-AuthoredDetection -File 'package-sysinternals.ps1' -Version '2026.10.15' -Variables @{ DetectionRegistryPath = $key }
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.RegistryKeyRelative | Should -Be 'SOFTWARE\App Packager\Sysinternals Suite'
        $d.ValueName | Should -Be 'Version'
        $d.PropertyType | Should -Be 'Version'
        $d.ExpectedValue | Should -Be '2026.10.15'
        $source = Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\package-sysinternals.ps1"
        $source | Should -Match "New-ItemProperty -Path '\`$DetectionRegistryPath' -Name Version -Value '\`$version'"
        $source | Should -Match "Remove-Item -Path '\`$DetectionRegistryPath'"
    }
}

Describe 'Anaconda' {
    BeforeAll {
        $script:AnacondaSource = Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\package-anaconda.ps1"
    }

    It 'detects by script, comparing the year, month and build of every Anaconda3 entry' {
        $d = Get-AuthoredDetection -File 'package-anaconda.ps1' -Version '2026.07-1' -Variables @{ detectionScript = 'script text' }
        $d.Type | Should -Be 'Script'
        $d.ScriptLanguage | Should -Be 'PowerShell'
        $d.ScriptText | Should -Be 'script text'
        $script:AnacondaSource | Should -Match "StartsWith\(''Anaconda3 ''"
        $script:AnacondaSource | Should -Match 'Write-Output ''''Installed''''; exit 0'
    }

    It 'turns the archive name into a three-number minimum' {
        $m = [regex]::Match('2026.07-1', '^(?<year>\d{4})\.(?<month>\d{2})-(?<build>\d+)$')
        ('{0}.{1}.{2}' -f $m.Groups['year'].Value, [int]$m.Groups['month'].Value, [int]$m.Groups['build'].Value) | Should -Be '2026.7.1'
        ([version]'2026.10.1') -gt ([version]'2026.7.2') | Should -BeTrue
    }

    It 'waits for the uninstaller by keeping it in place' {
        $script:AnacondaSource | Should -Match "_\?=\{0\}"
        $script:AnacondaSource | Should -Match 'Remove-Item -LiteralPath \$installDir -Recurse -Force'
    }
}

Describe 'Android Studio' {
    BeforeAll {
        $script:AndroidSource = Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\package-androidstudio.ps1"
    }

    It 'detects by script on the build number in product-info.json' {
        $d = Get-AuthoredDetection -File 'package-androidstudio.ps1' -Version '2026.2.1.8' -Variables @{ detectionScript = 'script text' }
        $d.Type | Should -Be 'Script'
        $d.ScriptText | Should -Be 'script text'
        $script:AndroidSource | Should -Match 'product-info\.json'
        $script:AndroidSource | Should -Match 'Build\s+= \(\[string\]\$latest\.build\) -replace'
    }

    It 'counts exit code 1223 as success only when studio64.exe exists' {
        $script:AndroidSource | Should -Match 'if \(`\$code -eq 1223 -and \(Test-Path'
        $script:AndroidSource | Should -Match 'bin\\studio64\.exe'
    }
}

Describe 'CutePDF Writer' {
    It 'needs the uninstall entry and the version the install wrote' {
        $d = Get-AuthoredDetection -File 'package-cutepdfwriter.ps1' -Version '4.0.1.5' -Variables @{
            DetectionRegistryKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CutePDF Writer Installation'
            MarkerRegistryKey    = 'SOFTWARE\App Packager\CutePDF Writer'
        }
        $d.Type | Should -Be 'Compound'
        @($d.Clauses).Count | Should -Be 2
        $d.Clauses[0].Type | Should -Be 'RegistryKey'
        $d.Clauses[1].Type | Should -Be 'RegistryKeyValue'
        $d.Clauses[1].RegistryKeyRelative | Should -Be 'SOFTWARE\App Packager\CutePDF Writer'
        $d.Clauses[1].ValueName | Should -Be 'Version'
        $d.Clauses[1].PropertyType | Should -Be 'Version'
        $d.Clauses[1].Operator | Should -Be 'GreaterEquals'
        $d.Clauses[1].ExpectedValue | Should -Be '4.0.1.5'
    }

    It 'writes the marker on a successful install and removes it on uninstall' {
        $source = Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\package-cutepdfwriter.ps1"
        $source | Should -Match "New-ItemProperty -Path 'HKLM:\\\{0\}' -Name Version"
        $source | Should -Match 'if \(\$proc\.ExitCode -eq 0\) \{ Remove-Item -LiteralPath \$marker'
    }
}

Describe 'Install switches proven on a clean machine' {
    It 'passes the install scope to <File>' -ForEach @(
        @{ File = 'package-vim.ps1'; Pattern = "-InstallArgs \`"'/S','/allusers'\`"" }
        @{ File = 'package-vim.ps1'; Pattern = "ArgumentList @\('/S', '/allusers'\)" }
        @{ File = 'package-vim.ps1'; Pattern = 'arpRegistryKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Vim"' }
        @{ File = 'package-pgadmin4.ps1'; Pattern = "'/SP-', '/ALLUSERS'" }
        @{ File = 'package-aimp.ps1'; Pattern = "-InstallArgs \`"'/AUTO','/SILENT'\`"" }
        @{ File = 'package-aimp.ps1'; Pattern = "ArgumentList @\('/AUTO','/SILENT'\)" }
    ) {
        (Get-Content -Raw -LiteralPath "$PSScriptRoot\..\Packagers\$File") | Should -Match $Pattern
    }
}

Describe 'Get-BurnBundleId' {
    It 'reads the bundle id from the .wixburn header as a lower-case braced guid' {
        $id = [guid]'CEF0D74C-33BC-432B-941F-6E17BEA1F9D3'
        $path = Join-Path $TestDrive 'bundle.exe'
        New-BurnStub -Path $path -BundleId $id
        Get-BurnBundleId -Path $path | Should -Be '{cef0d74c-33bc-432b-941f-6e17bea1f9d3}'
    }

    It 'returns nothing for a PE file without a .wixburn section' {
        $path = Join-Path $TestDrive 'plain.exe'
        New-BurnStub -Path $path -BundleId ([guid]::NewGuid()) -SectionName '.text'
        Get-BurnBundleId -Path $path | Should -BeNullOrEmpty
    }

    It 'returns nothing for a file that is not a PE image' {
        $path = Join-Path $TestDrive 'notes.txt'
        Set-Content -LiteralPath $path -Value ('x' * 200)
        Get-BurnBundleId -Path $path | Should -BeNullOrEmpty
    }
}
