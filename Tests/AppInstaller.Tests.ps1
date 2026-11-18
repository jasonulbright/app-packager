#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-appinstaller.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Select-WinGetRelease', 'Expand-Dependencies', 'New-AppInstallerInstallScript', 'New-AppInstallerUninstallScript', 'New-AppInstallerDetectionScript') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    function New-Release {
        param([string]$Tag, [string[]]$Skip = @(), [switch]$Prerelease)
        $names = @('Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle', 'DesktopAppInstaller_Dependencies.zip', 'DesktopAppInstaller_Dependencies.json', 'e53e159d00e04f729cc2180cffd1c02e_License1.xml', 'DesktopAppInstallerPolicies.zip') | Where-Object { $Skip -notcontains $_ }
        $assets = foreach ($n in $names) { @{ name = $n; browser_download_url = "https://github.com/microsoft/winget-cli/releases/download/$Tag/$n"; digest = 'sha256:' + ('cd' * 32) } }
        @{ tag_name = $Tag; draft = $false; prerelease = [bool]$Prerelease; assets = @($assets) }
    }
    function ConvertTo-ReleaseJson { param([object[]]$Releases) ConvertTo-Json -InputObject $Releases -Depth 6 }
    function Test-Parses { param([string]$Text) $tk = $null; $er = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tk, [ref]$er); @($er).Count -eq 0 }
}

Describe 'App Installer release selection' {
    It 'picks the newest stable release with every file and finds the license by its suffix' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v1.29.380'), (New-Release -Tag 'v1.29.290'))
        $r = Select-WinGetRelease -Json $json
        $r.Version | Should -Be '1.29.380'
        $r.Bundle.Name | Should -Be 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
        $r.License.Name | Should -Be 'e53e159d00e04f729cc2180cffd1c02e_License1.xml'
        $r.Bundle.Sha256 | Should -Be ('CD' * 32)
    }

    It 'skips a preview release' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v1.30.140-preview' -Prerelease), (New-Release -Tag 'v1.29.380'))
        (Select-WinGetRelease -Json $json).Version | Should -Be '1.29.380'
    }

    It 'skips a stable-looking tag that is flagged as a prerelease' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v1.30.1' -Prerelease), (New-Release -Tag 'v1.29.380'))
        (Select-WinGetRelease -Json $json).Version | Should -Be '1.29.380'
    }

    It 'skips a release without the <Missing> asset' -ForEach @(
        @{ Missing = 'DesktopAppInstaller_Dependencies.zip' }
        @{ Missing = 'DesktopAppInstaller_Dependencies.json' }
        @{ Missing = 'e53e159d00e04f729cc2180cffd1c02e_License1.xml' }
        @{ Missing = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle' }
    ) {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v1.29.400' -Skip @($Missing)), (New-Release -Tag 'v1.29.380'))
        (Select-WinGetRelease -Json $json).Version | Should -Be '1.29.380'
    }

    It 'throws when no release qualifies' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v1.30.140-preview' -Prerelease))
        { Select-WinGetRelease -Json $json } | Should -Throw '*No stable winget-cli release*'
    }
}

Describe 'App Installer dependency extraction' {
    BeforeAll {
        function New-DepsZip {
            param([string]$Path, [string[]]$Entries)
            $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($name in $Entries) {
                    $w = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open()); $w.Write('x'); $w.Dispose()
                }
            }
            finally { $zip.Dispose() }
        }
        $script:Json = Join-Path $TestDrive 'deps.json'
        Set-Content -LiteralPath $script:Json -Encoding ASCII -Value '{ "Dependencies": [ { "Name": "Contoso.Lib", "Version": "1.0.0.0" }, { "Name": "Contoso.Runtime", "Version": "2.0.0.0" } ] }'
    }

    It 'extracts each listed dependency for the requested architectures only' {
        $zip = Join-Path $TestDrive 'deps-ok.zip'
        New-DepsZip -Path $zip -Entries @('x64/Contoso.Lib_1.0.0.0_x64.appx', 'x64/Contoso.Runtime_2.0.0.0_x64.appx', 'x86/Contoso.Lib_1.0.0.0_x86.appx', 'arm64/Contoso.Lib_1.0.0.0_arm64.appx', 'arm64/Contoso.Runtime_2.0.0.0_arm64.appx')
        $dest = Join-Path $TestDrive 'out-ok'
        Expand-Dependencies -ZipPath $zip -JsonPath $script:Json -Destination $dest -Architecture @('x64', 'arm64')
        (Get-ChildItem -LiteralPath (Join-Path $dest 'x64') -Filter *.appx).Count | Should -Be 2
        (Get-ChildItem -LiteralPath (Join-Path $dest 'arm64') -Filter *.appx).Count | Should -Be 2
        Test-Path -LiteralPath (Join-Path $dest 'x86') | Should -BeFalse
    }

    It 'throws when a listed dependency is not in the archive' {
        $zip = Join-Path $TestDrive 'deps-missing.zip'
        New-DepsZip -Path $zip -Entries @('x64/Contoso.Lib_1.0.0.0_x64.appx')
        { Expand-Dependencies -ZipPath $zip -JsonPath $script:Json -Destination (Join-Path $TestDrive 'out-missing') -Architecture @('x64') } | Should -Throw '*has no x64/Contoso.Runtime_2.0.0.0_x64.appx*'
    }
}

Describe 'App Installer scripts' {
    It 'writes an install script that parses, provisions with dependencies and the license, and checks the OS build' {
        $text = New-AppInstallerInstallScript -BundleFileName 'Bundle.msixbundle'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "Join-Path \`$PSScriptRoot 'Bundle\.msixbundle'"
        $text | Should -Match 'Add-AppxProvisionedPackage -Online -PackagePath \$bundle -DependencyPackagePath \$dependencies -LicensePath \$license'
        $text | Should -Match "10\.0\.17763"
        $text | Should -Match 'AMD64'
        $text | Should -Match 'ARM64'
        $text | Should -Not -Match 'SkipLicense'
    }

    It 'writes an uninstall script that parses, removes the package for all users, then the provisioning' {
        $text = New-AppInstallerUninstallScript -PackageIdentity 'Microsoft.DesktopAppInstaller'
        Test-Parses $text | Should -BeTrue
        $text.IndexOf('Remove-AppxPackage') | Should -BeLessThan $text.IndexOf('Remove-AppxProvisionedPackage')
        $text | Should -Match "Get-AppxPackage -AllUsers -Name 'Microsoft\.DesktopAppInstaller'"
        $text | Should -Match 'still present'
    }

    It 'writes a detection script that parses and compares the provisioned version' {
        $text = New-AppInstallerDetectionScript -PackageIdentity 'Microsoft.DesktopAppInstaller' -MinimumVersion '2026.917.151.0'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "\[version\]'2026\.917\.151\.0'"
        $text | Should -Match 'Registry64'
        $text | Should -Match '-ge \$minimum'
    }

    It 'detection writes nothing when no package key matches' {
        $script = New-AppInstallerDetectionScript -PackageIdentity 'Contoso.Fake' -MinimumVersion '2.0.0.0'
        # The generated script reads the real store; with no matching key it writes nothing.
        $result = & ([scriptblock]::Create(($script -replace 'exit 0', '')))
        $result | Should -BeNullOrEmpty
    }
}

Describe 'App Installer packager contract' {
    BeforeAll { $script:Text = Get-Content -LiteralPath $script:PackagerPath -Raw }

    It 'tags the header' {
        $script:Text | Should -Match '(?m)^Vendor: Microsoft\r?$'
        $script:Text | Should -Match '(?m)^CMName: App Installer \(WinGet\)\r?$'
        $script:Text | Should -Match '(?m)^CPE: cpe:2\.3:a:microsoft:app_installer:'
    }

    It 'builds no download URL from a version' {
        $script:Text | Should -Not -Match 'releases/download/v\$'
    }

    It 'detects by script' {
        $script:Text | Should -Match 'Type\s*=\s*"Script"'
    }
}
