#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The GitHub token that Options, Tools stores: its order of use and its secrecy.
#>

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ap-ghtoken-' + [guid]::NewGuid().ToString('N'))
    $packagers = Join-Path $script:Temp 'Packagers'
    [void](New-Item -ItemType Directory -Path $packagers -Force)
    Copy-Item -LiteralPath (Join-Path $script:Root 'Packagers\AppPackagerCommon.psm1') -Destination $packagers
    Copy-Item -LiteralPath (Join-Path $script:Root 'Lib') -Destination $script:Temp -Recurse
    $script:Secret = 'ghp_unit_test_token_123'
    $protected = ConvertFrom-SecureString -SecureString (ConvertTo-SecureString -String $script:Secret -AsPlainText -Force)
    @{ GitHub = @{ TokenProtected = $protected } } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $script:Temp 'AppPackager.preferences.json') -Encoding UTF8
    Import-Module (Join-Path $packagers 'AppPackagerCommon.psm1') -Force -DisableNameChecking
    $script:SavedGithub = $env:GITHUB_TOKEN
    $script:SavedGh = $env:GH_TOKEN
}

AfterAll {
    $env:GITHUB_TOKEN = $script:SavedGithub
    $env:GH_TOKEN = $script:SavedGh
    Remove-Module AppPackagerCommon -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $script:Temp -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'The stored GitHub token' {
    BeforeEach { $env:GITHUB_TOKEN = $null; $env:GH_TOKEN = $null }

    It 'reads the DPAPI text of the preferences file' {
        Get-StoredGitHubToken | Should -Be $script:Secret
    }

    It 'returns an empty string when the text cannot be read' {
        $path = Join-Path $script:Temp 'AppPackager.preferences.json'
        $keep = Get-Content -LiteralPath $path -Raw
        try {
            '{ "GitHub": { "TokenProtected": "not-dpapi" } }' | Set-Content -LiteralPath $path -Encoding UTF8
            Get-StoredGitHubToken | Should -Be ''
        }
        finally { Set-Content -LiteralPath $path -Value $keep -Encoding UTF8 }
    }

    It 'is used after GITHUB_TOKEN and GH_TOKEN' {
        $env:GH_TOKEN = 'from-gh-token'
        (Get-GitHubApiCurlArgs) -join ' ' | Should -Match 'Bearer from-gh-token'
        $env:GITHUB_TOKEN = 'from-github-token'
        (Get-GitHubApiCurlArgs) -join ' ' | Should -Match 'Bearer from-github-token'
    }

    It 'is used before the gh login' {
        (Get-GitHubApiCurlArgs) -join ' ' | Should -Match ('Bearer ' + [regex]::Escape($script:Secret))
        $body = (Get-Command Get-GitHubApiCurlArgs).Definition
        $body.IndexOf('Get-StoredGitHubToken') | Should -BeLessThan $body.IndexOf('auth token')
    }

    It 'never goes to a log' {
        $text = Get-Content -LiteralPath (Join-Path $script:Root 'Packagers\AppPackagerCommon.psm1') -Raw
        $function = [regex]::Match($text, '(?s)function Get-StoredGitHubToken \{.*?\r?\n\}\r?\n').Value
        $function | Should -Not -Match 'Write-Log|Write-Host|Write-Output|Write-Verbose'
        $args = [regex]::Match($text, '(?s)function Get-GitHubApiCurlArgs \{.*?\r?\n\}\r?\n').Value
        $args | Should -Not -Match 'Write-Log|Write-Host|Write-Verbose'
    }
}
