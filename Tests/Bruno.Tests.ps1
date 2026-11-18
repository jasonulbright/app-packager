#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:path = Join-Path $PSScriptRoot '..\Packagers\package-bruno.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:path, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-LatestBrunoRelease' }, $false)
    if (-not $fn) { throw 'function Get-LatestBrunoRelease not found' }
    . ([scriptblock]::Create($fn.Extent.Text))
    function Write-Log { param($Message, $Level, [switch]$Quiet) }
    function Get-GitHubApiCurlArgs { @() }
    $script:GitHubApiUrl = 'https://example.invalid/releases/latest'
    $script:releaseJson = @'
{ "tag_name": "v4.3.0", "assets": [
  { "name": "bruno_4.3.0_arm64_win.msi", "browser_download_url": "https://example.invalid/arm.msi", "digest": "sha256:cb2e70556a80c2d3dd8132f439c162ba3db3124e6e5cf6e1d632d9e140f19b97" },
  { "name": "bruno_4.3.0_x64_win.exe", "browser_download_url": "https://example.invalid/x.exe", "digest": "sha256:e1149fba7def0bd5b9654d745c769a16e62212e2683c5a6523807262b369c3c3" },
  { "name": "bruno_4.3.0_x64_win.msi", "browser_download_url": "https://example.invalid/x64.msi", "digest": "sha256:551768b18db4232133511d2422c4a407c40f7867dac001368ed8eafb2cee1efa" }
] }
'@
}

Describe 'Bruno release lookup' {
    It 'picks the x64 MSI, strips the v from the tag and reads the digest' {
        function curl.exe { $global:LASTEXITCODE = 0; $script:releaseJson }
        $r = Get-LatestBrunoRelease -Quiet
        $r.Version | Should -Be '4.3.0'
        $r.FileName | Should -Be 'bruno_4.3.0_x64_win.msi'
        $r.DownloadUrl | Should -Be 'https://example.invalid/x64.msi'
        $r.Sha256 | Should -Be '551768B18DB4232133511D2422C4A407C40F7867DAC001368ED8EAFB2CEE1EFA'
    }

    It 'returns null when the release has no x64 MSI' {
        function curl.exe { $global:LASTEXITCODE = 0; '{ "tag_name": "v4.3.0", "assets": [ { "name": "bruno_4.3.0_arm64_win.msi" } ] }' }
        Get-LatestBrunoRelease -Quiet | Should -BeNullOrEmpty
    }

    It 'returns null for a tag that is not a version' {
        function curl.exe { $global:LASTEXITCODE = 0; '{ "tag_name": "nightly", "assets": [ { "name": "bruno_4.3.0_x64_win.msi", "browser_download_url": "u" } ] }' }
        Get-LatestBrunoRelease -Quiet | Should -BeNullOrEmpty
    }

    It 'returns no digest when GitHub publishes none' {
        function curl.exe { $global:LASTEXITCODE = 0; '{ "tag_name": "v4.3.0", "assets": [ { "name": "bruno_4.3.0_x64_win.msi", "browser_download_url": "u" } ] }' }
        $r = Get-LatestBrunoRelease -Quiet
        $r.Version | Should -Be '4.3.0'
        $r.FileName | Should -Be 'bruno_4.3.0_x64_win.msi'
        $r.Sha256 | Should -BeNullOrEmpty
    }
}
