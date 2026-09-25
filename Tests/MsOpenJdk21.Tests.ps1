Describe 'Microsoft Build of OpenJDK 21 parsers in <Script>' -ForEach @(
    @{ Script = 'package-ms-openjdk21-msi.ps1' }
    @{ Script = 'package-ms-openjdk21-msi-user.ps1' }
    @{ Script = 'package-ms-openjdk21-exe.ps1' }
    @{ Script = 'package-ms-openjdk21-exe-user.ps1' }
) {
    BeforeAll {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot "..\Packagers\$Script"), [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }
        foreach ($name in 'ConvertFrom-MsOpenJdkDownloadUrl', 'ConvertFrom-MsOpenJdkSha256File') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if (-not $fn) { throw "$name not found in $Script" }
            . ([scriptblock]::Create($fn.Extent.Text))
        }
    }

    It 'reads the version from the redirect target URL' {
        $url = 'https://download.visualstudio.microsoft.com/download/pr/f1e5f23f/fe87dea0/microsoft-jdk-21.0.12.1-windows-x64.msi'
        ConvertFrom-MsOpenJdkDownloadUrl -Url $url -Major 21 -Extension msi | Should -Be '21.0.12.1'
    }

    It 'reads a three-part version from a bare EXE file name' {
        ConvertFrom-MsOpenJdkDownloadUrl -Url 'microsoft-jdk-21.0.12-windows-x64.exe' -Major 21 -Extension exe | Should -Be '21.0.12'
    }

    It 'reads a five-part version from a bare MSI file name' {
        ConvertFrom-MsOpenJdkDownloadUrl -Url 'microsoft-jdk-21.0.12.1.1-windows-x64.msi' -Major 21 -Extension msi | Should -Be '21.0.12.1.1'
    }

    It 'ignores a query string after the file name' {
        ConvertFrom-MsOpenJdkDownloadUrl -Url 'https://host/microsoft-jdk-21.0.4-windows-x64.msi?sig=abc' -Major 21 -Extension msi | Should -Be '21.0.4'
    }

    It 'rejects a name that is not a JDK 21 x64 installer of that type: <Url>' -TestCases @(
        @{ Url = 'https://www.bing.com/?ref=aka&shorturl=download-jdk/microsoft-jdk-21-windows-x64.exe'; Extension = 'exe' }
        @{ Url = 'microsoft-jdk-21-windows-x64.msi'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-21.0.12.1.1.1-windows-x64.msi'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-17.0.20.1-windows-x64.msi'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-210.0.1-windows-x64.msi'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-21.0.12.1-windows-aarch64.msi'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-21.0.12.1-windows-x64.zip'; Extension = 'msi' }
        @{ Url = 'microsoft-jdk-21.0.12.1-windows-x64.msi'; Extension = 'exe' }
        @{ Url = 'microsoft-jdk-21.0.12.1-windows-x64.msi.sha256sum.txt'; Extension = 'msi' }
        @{ Url = ''; Extension = 'msi' }
    ) {
        param($Url, $Extension)
        ConvertFrom-MsOpenJdkDownloadUrl -Url $Url -Major 21 -Extension $Extension | Should -BeNullOrEmpty
    }

    It 'returns the hash listed for the file, lower-cased' {
        $content = "3A3F7E1A9FD9EDD1FCC0545FFBE7DC87787A8417DC1BF303E41C7A697C8490C1  microsoft-jdk-21.0.12.1-windows-x64.msi`n"
        ConvertFrom-MsOpenJdkSha256File -Content $content -FileName 'microsoft-jdk-21.0.12.1-windows-x64.msi' |
            Should -Be '3a3f7e1a9fd9edd1fcc0545ffbe7dc87787a8417dc1bf303e41c7a697c8490c1'
    }

    It 'accepts the binary-mode asterisk and CRLF line ends' {
        $content = "8d06f74ecb94729c353615c4af279681166f667733d307df92d91f63b0844cdf *microsoft-jdk-21.0.12.1-windows-x64.exe`r`n"
        ConvertFrom-MsOpenJdkSha256File -Content $content -FileName 'microsoft-jdk-21.0.12.1-windows-x64.exe' |
            Should -Be '8d06f74ecb94729c353615c4af279681166f667733d307df92d91f63b0844cdf'
    }

    It 'accepts a bare hash with no file name' {
        ConvertFrom-MsOpenJdkSha256File -Content ('a' * 64) -FileName 'x.msi' | Should -Be ('a' * 64)
    }

    It 'returns nothing when no line names the file or the hash is malformed: <Name>' -TestCases @(
        @{ Name = 'other file'; Content = "3a3f7e1a9fd9edd1fcc0545ffbe7dc87787a8417dc1bf303e41c7a697c8490c1  microsoft-jdk-21.0.12.1-windows-x64.exe" }
        @{ Name = 'short hash'; Content = "3a3f7e1a  microsoft-jdk-21.0.12.1-windows-x64.msi" }
        @{ Name = 'HTML page'; Content = "<html><body>Not found</body></html>" }
        @{ Name = 'empty'; Content = '' }
    ) {
        param($Name, $Content)
        ConvertFrom-MsOpenJdkSha256File -Content $Content -FileName 'microsoft-jdk-21.0.12.1-windows-x64.msi' | Should -BeNullOrEmpty
    }
}
