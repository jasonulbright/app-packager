BeforeAll {
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-unityhub.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-LatestUnityHubRelease' }, $true)
    . ([scriptblock]::Create($fn.Extent.Text))
    $HubCdnRoot = 'https://cdn.invalid/hub/prod/'
    $LatestYmlUrl = $HubCdnRoot + 'latest.yml'
    function Write-Log { param($Message, $Level, [switch]$Quiet) }
}

Describe 'Unity Hub update feed' {
    It 'reads the x64 installer from the files list' {
        function curl.exe {
            $global:LASTEXITCODE = 0
            'version: 3.22.2', 'files:', '  - url: 3.22.2/UnityHubSetup-3.22.2-x64.exe', '    sha512: AAAA', '    size: 1',
            '  - url: 3.22.2/UnityHubSetup-3.22.2-arm64.exe', '    sha512: BBBB', '    size: 1', 'releaseDate: 2026-10-02T20:18:59.543Z'
        }
        $r = Get-LatestUnityHubRelease -Quiet
        $r.Version | Should -Be '3.22.2'
        $r.FileName | Should -Be 'UnityHubSetup-3.22.2-x64.exe'
        $r.DownloadUrl | Should -Be 'https://cdn.invalid/hub/prod/3.22.2/UnityHubSetup-3.22.2-x64.exe'
    }

    It 'still reads a feed with a single path entry' {
        function curl.exe {
            $global:LASTEXITCODE = 0
            'version: 3.12.0', 'files:', '  - url: UnityHubSetup-x64.exe', 'path: UnityHubSetup-x64.exe', 'sha512: AAAA'
        }
        (Get-LatestUnityHubRelease -Quiet).FileName | Should -Be 'UnityHubSetup-x64.exe'
    }

    It 'returns nothing when the feed names no x64 installer' {
        function curl.exe { $global:LASTEXITCODE = 0; 'version: 3.22.2', 'files:', '  - url: 3.22.2/UnityHubSetup-3.22.2-arm64.exe' }
        Get-LatestUnityHubRelease -Quiet | Should -BeNullOrEmpty
    }
}
