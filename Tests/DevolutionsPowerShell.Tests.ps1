#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-devolutionspowershell.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Select-DevolutionsPowerShellAsset') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    function New-ReleaseJson {
        param([object[]]$Releases)
        ConvertTo-Json -InputObject $Releases -Depth 6
    }
    function New-Release {
        param([string]$Tag, [string[]]$Assets, [switch]$Prerelease, [switch]$NoDigest)
        $list = foreach ($a in $Assets) {
            $digest = if ($NoDigest) { $null } else { 'sha256:' + ('ab' * 32) }
            @{ name = $a; browser_download_url = "https://github.com/Devolutions/pwsh-distro/releases/download/$Tag/$a"; digest = $digest }
        }
        @{ tag_name = $Tag; draft = $false; prerelease = [bool]$Prerelease; assets = @($list) }
    }
}

Describe 'Devolutions PowerShell release selection' {
    It 'picks the Microsoft-payload x64 MSI of the newest release that lists one' {
        $json = New-ReleaseJson @(
            (New-Release -Tag 'v7.7.0.1' -Assets @('Devolutions.PowerShell-7.7.0.1-win-x64.msi', 'Microsoft.PowerShell-7.7.0.1-win-arm64.msi', 'Microsoft.PowerShell-7.7.0.1-win-x64.msi')),
            (New-Release -Tag 'v7.7.0.0' -Assets @('Microsoft.PowerShell-7.7.0.0-win-x64.msi'))
        )
        $r = Select-DevolutionsPowerShellAsset -Json $json
        $r.Version | Should -Be '7.7.0.1'
        $r.Tag | Should -Be 'v7.7.0.1'
        $r.AssetName | Should -Be 'Microsoft.PowerShell-7.7.0.1-win-x64.msi'
        $r.DownloadUrl | Should -Be 'https://github.com/Devolutions/pwsh-distro/releases/download/v7.7.0.1/Microsoft.PowerShell-7.7.0.1-win-x64.msi'
        $r.Sha256 | Should -Be (('AB' * 32))
    }

    It 'skips a newer release that has only the Devolutions-flavored MSI' {
        $json = New-ReleaseJson @(
            (New-Release -Tag 'v7.7.0.2' -Assets @('Devolutions.PowerShell-7.7.0.2-win-x64.msi')),
            (New-Release -Tag 'v7.7.0.1' -Assets @('Microsoft.PowerShell-7.7.0.1-win-x64.msi'))
        )
        (Select-DevolutionsPowerShellAsset -Json $json).Version | Should -Be '7.7.0.1'
    }

    It 'skips a prerelease' {
        $json = New-ReleaseJson @(
            (New-Release -Tag 'v7.8.0.0' -Assets @('Microsoft.PowerShell-7.8.0.0-win-x64.msi') -Prerelease),
            (New-Release -Tag 'v7.7.0.1' -Assets @('Microsoft.PowerShell-7.7.0.1-win-x64.msi'))
        )
        (Select-DevolutionsPowerShellAsset -Json $json).Version | Should -Be '7.7.0.1'
    }

    It 'skips an asset without a digest because the download cannot be verified' {
        $json = New-ReleaseJson @(
            (New-Release -Tag 'v7.7.0.2' -Assets @('Microsoft.PowerShell-7.7.0.2-win-x64.msi') -NoDigest),
            (New-Release -Tag 'v7.7.0.1' -Assets @('Microsoft.PowerShell-7.7.0.1-win-x64.msi'))
        )
        (Select-DevolutionsPowerShellAsset -Json $json).Version | Should -Be '7.7.0.1'
    }

    It 'throws when no release lists the Microsoft-payload MSI' {
        $json = New-ReleaseJson @(
            (New-Release -Tag 'v7.6.10.0' -Assets @('Devolutions-PowerShell-7.6.10.0-win-x64.msi'))
        )
        { Select-DevolutionsPowerShellAsset -Json $json } | Should -Throw '*Microsoft-payload*'
    }
}

Describe 'Devolutions PowerShell packager contract' {
    BeforeAll { $script:Text = Get-Content -LiteralPath $script:PackagerPath -Raw }

    It 'builds no download URL from a version' {
        $script:Text | Should -Not -Match 'releases/download/v\$'
    }

    It 'detects by the pwsh.exe file version at or above the packaged version' {
        $script:Text | Should -Match 'Type\s*=\s*"File"'
        $script:Text | Should -Match 'FileName\s*=\s*"pwsh\.exe"'
        $script:Text | Should -Match 'Operator\s*=\s*"GreaterEquals"'
    }

    It 'installs and uninstalls the MSI silently without a restart' {
        $script:Text | Should -Match "''/qn'', ''/norestart'', ''ADD_PATH=1''"
        $script:Text | Should -Match "''/x''.*''/qn'', ''/norestart''"
    }

    It 'tags the header' {
        $script:Text | Should -Match '(?m)^Vendor: Devolutions\r?$'
        $script:Text | Should -Match '(?m)^CMName: PowerShell 7 \(Devolutions\)\r?$'
        $script:Text | Should -Match '(?m)^WsusSupport: Yes\r?$'
    }
}

Describe 'Devolutions PowerShell version text' {
    It 'says that the latest version is the release number, not the pwsh.exe file version' {
        $text = Get-Content -LiteralPath $script:PackagerPath -Raw
        $text | Should -Match 'pwsh-distro release number, not the pwsh\.exe file version'
        $text | Should -Match 'not\s+the file version of pwsh\.exe'
    }
}
