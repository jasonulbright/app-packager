#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-powershell7msix.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Select-PowerShellMsixRelease', 'Get-BundleInfo', 'New-PowerShellMsixInstallScript', 'New-PowerShellMsixUninstallScript', 'New-PowerShellMsixDetectionScript') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    function New-Release {
        param([string]$Tag, [string[]]$Names, [switch]$Prerelease, [switch]$NoDigest)
        if (-not $Names) { $v = $Tag.TrimStart('v'); $Names = @("PowerShell-$v-win-x64.zip", "PowerShell-$v.msixbundle", "PowerShell-LTS-$v.msixbundle") }
        $assets = foreach ($n in $Names) {
            $a = @{ name = $n; browser_download_url = "https://github.com/PowerShell/PowerShell/releases/download/$Tag/$n" }
            if (-not $NoDigest) { $a.digest = 'sha256:' + ('ab' * 32) }
            $a
        }
        @{ tag_name = $Tag; draft = $false; prerelease = [bool]$Prerelease; assets = @($assets) }
    }
    function ConvertTo-ReleaseJson { param([object[]]$Releases) ConvertTo-Json -InputObject $Releases -Depth 6 }
    function Test-Parses { param([string]$Text) $tk = $null; $er = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tk, [ref]$er); @($er).Count -eq 0 }
}

Describe 'PowerShell 7 MSIX release selection' {
    It 'picks the newest stable release and takes the plain bundle, not the LTS one' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v7.7.0'), (New-Release -Tag 'v7.6.10'))
        $r = Select-PowerShellMsixRelease -Json $json
        $r.Version | Should -Be '7.7.0'
        $r.Name | Should -Be 'PowerShell-7.7.0.msixbundle'
        $r.Sha256 | Should -Be ('AB' * 32)
        $r.DownloadUrl | Should -Be 'https://github.com/PowerShell/PowerShell/releases/download/v7.7.0/PowerShell-7.7.0.msixbundle'
    }

    It 'skips a preview release and a preview bundle' {
        $json = ConvertTo-ReleaseJson @(
            (New-Release -Tag 'v7.8.0-preview.1' -Prerelease -Names @('PowerShellPreview-7.8.0-preview.1.msixbundle')),
            (New-Release -Tag 'v7.7.0-rc.1' -Names @('PowerShell-7.7.0-rc.1.msixbundle')),
            (New-Release -Tag 'v7.6.10'))
        (Select-PowerShellMsixRelease -Json $json).Version | Should -Be '7.6.10'
    }

    It 'skips a release that has no bundle' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v7.7.1' -Names @('PowerShell-7.7.1-win-x64.zip')), (New-Release -Tag 'v7.7.0'))
        (Select-PowerShellMsixRelease -Json $json).Version | Should -Be '7.7.0'
    }

    It 'skips a bundle that has no SHA-256 digest' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v7.7.1' -NoDigest), (New-Release -Tag 'v7.7.0'))
        (Select-PowerShellMsixRelease -Json $json).Version | Should -Be '7.7.0'
    }

    It 'throws when no release qualifies' {
        $json = ConvertTo-ReleaseJson @((New-Release -Tag 'v7.8.0-preview.1' -Prerelease))
        { Select-PowerShellMsixRelease -Json $json } | Should -Throw '*No stable PowerShell release*'
    }
}

Describe 'PowerShell 7 MSIX bundle identity' {
    BeforeAll {
        function New-Bundle {
            param([string]$Path, [string]$Name = 'Microsoft.PowerShell', [string]$Architectures = 'arm64,x64,x86')
            $xml = '<?xml version="1.0"?><Bundle xmlns="http://schemas.microsoft.com/appx/2013/bundle"><Identity Name="' + $Name + '" Publisher="CN=Microsoft Corporation" Version="2026.1110.418.0"/><Packages>' +
                (($Architectures -split ',' | ForEach-Object { '<Package Type="application" Version="7.7.0.0" Architecture="' + $_ + '" FileName="p-' + $_ + '.msix"/>' }) -join '') + '</Packages></Bundle>'
            $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
            try { $w = New-Object System.IO.StreamWriter($zip.CreateEntry('AppxMetadata/AppxBundleManifest.xml').Open()); $w.Write($xml); $w.Dispose() } finally { $zip.Dispose() }
        }
    }

    It 'reads the identity name and the x64 package version, not the date-based bundle version' {
        $p = Join-Path $TestDrive 'ok.msixbundle'
        New-Bundle -Path $p
        $i = Get-BundleInfo -Path $p
        $i.Name | Should -Be 'Microsoft.PowerShell'
        $i.PackageVersion | Should -Be '7.7.0.0'
        $i.BundleVersion | Should -Be '2026.1110.418.0'
    }

    It 'throws when the bundle has no x64 package' {
        $p = Join-Path $TestDrive 'noX64.msixbundle'
        New-Bundle -Path $p -Architectures 'arm64'
        { Get-BundleInfo -Path $p } | Should -Throw '*no x64 application package*'
    }
}

Describe 'PowerShell 7 MSIX scripts' {
    It 'writes an install script that parses and provisions the bundle for all users' {
        $text = New-PowerShellMsixInstallScript -BundleFileName 'PowerShell-7.7.0.msixbundle'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "Join-Path \`$PSScriptRoot 'PowerShell-7\.7\.0\.msixbundle'"
        $text | Should -Match 'Add-AppxProvisionedPackage -Online -PackagePath \$bundle -SkipLicense'
        $text | Should -Match '10\.0\.17763'
    }

    It 'writes an uninstall script that parses, removes the package for all users, then the provisioning' {
        $text = New-PowerShellMsixUninstallScript -PackageIdentity 'Microsoft.PowerShell'
        Test-Parses $text | Should -BeTrue
        $text.IndexOf('Remove-AppxPackage') | Should -BeLessThan $text.IndexOf('Remove-AppxProvisionedPackage')
        $text | Should -Match "Get-AppxPackage -AllUsers -Name 'Microsoft\.PowerShell'"
        $text | Should -Match 'still present'
    }

    It 'writes a detection script that parses and compares the provisioning and the x64 package folder version' {
        $text = New-PowerShellMsixDetectionScript -PackageIdentity 'Microsoft.PowerShell' -MinimumVersion '7.7.0.0'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "\[version\]'7\.7\.0\.0'"
        $text | Should -Match 'Registry64'
        $text | Should -Match '_neutral_~_8wekyb3d8bbwe'
        $text | Should -Match '_x64__8wekyb3d8bbwe'
        $text | Should -Match 'WindowsApps'
        $text | Should -Match '-ge \$minimum'
    }

    It 'detection writes nothing when no package key matches' {
        $script = New-PowerShellMsixDetectionScript -PackageIdentity 'Contoso.Fake' -MinimumVersion '2.0.0.0'
        $result = & ([scriptblock]::Create(($script -replace 'exit 0', '')))
        $result | Should -BeNullOrEmpty
    }
}

Describe 'PowerShell 7 MSIX packager contract' {
    BeforeAll { $script:Text = Get-Content -LiteralPath $script:PackagerPath -Raw }

    It 'tags the header' {
        $script:Text | Should -Match '(?m)^Vendor: Microsoft\r?$'
        $script:Text | Should -Match '(?m)^CMName: PowerShell 7 \(MSIX\)\r?$'
        $script:Text | Should -Match '(?m)^IconSource: External\r?$'
        $script:Text | Should -Match '(?m)^CPE: cpe:2\.3:a:microsoft:powershell:'
    }

    It 'verifies the download by digest and by a Microsoft signature' {
        $script:Text | Should -Match 'SHA256 mismatch'
        $script:Text | Should -Match "CN=Microsoft Corporation"
        $script:Text | Should -Match "Status -ne 'Valid'"
    }

    It 'documents the single-user, no-remoting and all-users-scope limits of the MSIX package' {
        $script:Text | Should -Match 'single-user installation'
        $script:Text | Should -Match 'does not support PowerShell remoting'
        $script:Text | Should -Match 'Register-PSSessionConfiguration'
        $script:Text | Should -Match 'Set-ExecutionPolicy -Scope LocalMachine'
    }

    It 'detects by script' {
        $script:Text | Should -Match 'Type\s*=\s*"Script"'
    }

    It 'builds no download URL from a version' {
        $script:Text | Should -Not -Match 'releases/download/v\$'
    }
}
