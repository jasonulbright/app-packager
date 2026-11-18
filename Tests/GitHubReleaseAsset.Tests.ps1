#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Select-GitHubReleaseWithAsset and the PowerShell 7 packager: the newest release that carries the MSI.

.DESCRIPTION
    A fake release list. PowerShell 7.7.0 publishes no MSI, so the packager takes 7.6.10, from the asset list and not from a
    file name built out of the version.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function New-FakeRelease {
        param([string]$Tag, [string]$Published, [string[]]$Assets, [switch]$Draft, [switch]$Prerelease)
        [pscustomobject]@{
            tag_name = $Tag; published_at = $Published; draft = [bool]$Draft; prerelease = [bool]$Prerelease
            assets = @($Assets | ForEach-Object { [pscustomobject]@{ name = $_; browser_download_url = "https://example.test/$Tag/$_" } })
        }
    }
    $script:Msi = '^PowerShell-.+-win-x64\.msi$'
    $script:Without = @('PowerShell-7.7.0-win-x64.zip', 'PowerShell-7.7.0-win.msixbundle', 'powershell_7.7.0-1.deb_amd64.deb')
    $script:With = @('PowerShell-7.6.10-win-x64.msi', 'PowerShell-7.6.10-win-x86.msi', 'PowerShell-7.6.10-win-arm64.msi', 'PowerShell-7.6.10-win-x64.zip')
}

Describe 'Select-GitHubReleaseWithAsset' {
    It 'takes the older release when the newest one carries no MSI, and names the skipped release' {
        $releases = @(
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' $script:Without),
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With))
        $r = Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi
        $r.Version | Should -Be '7.6.10'
        $r.AssetName | Should -Be 'PowerShell-7.6.10-win-x64.msi'
        $r.DownloadUrl | Should -Be 'https://example.test/v7.6.10/PowerShell-7.6.10-win-x64.msi'
        $r.SkippedNewer | Should -Be @('7.7.0')
    }

    It 'takes the newest release when it carries the MSI' {
        $releases = @(
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' @('PowerShell-7.7.0-win-x64.msi', 'PowerShell-7.7.0-win-x64.zip')),
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With))
        $r = Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi
        $r.Version | Should -Be '7.7.0'
        $r.SkippedNewer.Count | Should -Be 0
    }

    It 'orders by the publish date, not by the order of the list' {
        $releases = @(
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With),
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' $script:Without))
        (Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
        $releases = @(
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With),
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' @('PowerShell-7.7.0-win-x64.msi')))
        (Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi).Version | Should -Be '7.7.0'
    }

    It 'passes over a draft and a pre-release' {
        $releases = @(
            (New-FakeRelease 'v7.8.0' '2026-11-12T10:00:00Z' @('PowerShell-7.8.0-win-x64.msi') -Draft),
            (New-FakeRelease 'v7.8.0-preview.1' '2026-11-11T10:00:00Z' @('PowerShell-7.8.0-preview.1-win-x64.msi') -Prerelease),
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With))
        $r = Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi
        $r.Version | Should -Be '7.6.10'
        $r.SkippedNewer.Count | Should -Be 0
    }

    It 'picks the variant that the pattern names' {
        $releases = @((New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With))
        (Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern '^PowerShell-.+-win-x86\.msi$').AssetName | Should -Be 'PowerShell-7.6.10-win-x86.msi'
        (Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern '^PowerShell-.+-win-arm64\.msi$').AssetName | Should -Be 'PowerShell-7.6.10-win-arm64.msi'
    }

    It 'reads the list that Windows PowerShell 5.1 returns from a JSON array' {
        $json = ConvertTo-Json -Depth 6 -InputObject @(
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' $script:Without),
            (New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' $script:With))
        $parsed = ConvertFrom-Json $json
        (Select-GitHubReleaseWithAsset -Releases @($parsed) -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
        (Select-GitHubReleaseWithAsset -Releases @(,@($parsed)) -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
    }

    It 'returns nothing when no release carries the asset, and for an empty list' {
        $releases = @((New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' $script:Without))
        Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi | Should -BeNullOrEmpty
        Select-GitHubReleaseWithAsset -Releases @() -AssetPattern $script:Msi | Should -BeNullOrEmpty
    }

    It 'takes the file name from the asset and not from the tag' {
        $releases = @((New-FakeRelease 'v7.6.10' '2026-10-01T10:00:00Z' @('PowerShell-7.6.10.0-win-x64.msi')))
        (Select-GitHubReleaseWithAsset -Releases $releases -AssetPattern $script:Msi).AssetName | Should -Be 'PowerShell-7.6.10.0-win-x64.msi'
    }
}

Describe 'package-powershell7.ps1' {
    BeforeAll { $script:Text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-powershell7.ps1')) }

    It 'reads the release list and the asset, and builds no file name or URL from the version' {
        $script:Text | Should -Match 'Select-GitHubReleaseWithAsset'
        $script:Text | Should -Not -Match 'releases/latest'
        $script:Text | Should -Not -Match 'PowerShell-\$\{version\}-win-x64\.msi'
        $script:Text | Should -Not -Match 'releases/download/v\$\{version\}'
    }

    It 'names the release that has no MSI' {
        $script:Text | Should -Match 'publishes no MSI; packaging'
    }

    It 'installs with Microsoft Update on and detects with a greater or equal version rule' {
        $script:Text | Should -Match "'USE_MU=1', ''ENABLE_MU=1'|''USE_MU=1'', ''ENABLE_MU=1''"
        $script:Text | Should -Match 'USE_MU=1 ENABLE_MU=1'
        $script:Text | Should -Not -Match 'USE_MU=0'
        $script:Text | Should -Match 'Operator\s*=\s*"GreaterEquals"'
        $script:Text | Should -Not -Match 'Operator\s*=\s*"(IsEquals|Equals)"'
    }
}

Describe 'Get-CurrentNewerNote' {
    It 'explains a newer installed version of PowerShell 7' {
        Get-CurrentNewerNote -PackagerScript 'package-powershell7.ps1' | Should -Match 'Microsoft Update'
        Get-CurrentNewerNote -PackagerScript 'C:\x\Packagers\package-powershell7.ps1' | Should -Match 'compliant'
    }

    It 'gives no note for another packager or an empty name' {
        Get-CurrentNewerNote -PackagerScript 'package-7zip.ps1' | Should -Be ''
        Get-CurrentNewerNote -PackagerScript '' | Should -Be ''
        Get-CurrentNewerNote -PackagerScript $null | Should -Be ''
    }

    It 'is called after both checks in the start script' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
        ([regex]::Matches($text, 'Get-CurrentNewerNote')).Count | Should -Be 2
    }
}

Describe 'Assert-DownloadAvailable' {
    BeforeAll {
        # A local server: /ok answers 206 to a range request, anything else 404.
        $script:Port = Get-Random -Minimum 20000 -Maximum 40000
        $script:Listener = New-Object System.Net.HttpListener
        $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
        $script:Listener.Start()
        $script:Ps = [powershell]::Create()
        [void]$script:Ps.AddScript({
                param($listener)
                while ($listener.IsListening) {
                    try { $ctx = $listener.GetContext() } catch { break }
                    if ($ctx.Request.Url.AbsolutePath -eq '/ok') { $ctx.Response.StatusCode = 206 } else { $ctx.Response.StatusCode = 404 }
                    $ctx.Response.Close()
                }
            }).AddArgument($script:Listener)
        $script:Handle = $script:Ps.BeginInvoke()
    }
    AfterAll {
        try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
        try { $script:Ps.Dispose() } catch { $null = $_ }
    }

    It 'passes when the file answers' {
        { Assert-DownloadAvailable -Url "http://localhost:$($script:Port)/ok" } | Should -Not -Throw
    }

    It 'stops with a clear message when the file is not published' {
        { Assert-DownloadAvailable -Url "http://localhost:$($script:Port)/missing.msi" -What 'The Fixture installer' } |
            Should -Throw -ExpectedMessage '*The Fixture installer is not published at*(HTTP status 404)*Nothing was downloaded.*'
    }
}

Describe 'Packagers that build a download URL' {
    It 'check the file before the download: <Name>' -ForEach @(
        @{ Name = 'positron' }, @{ Name = 'keepass' }, @{ Name = 'winscp' }, @{ Name = 'soapui' }, @{ Name = 'cryptomator' },
        @{ Name = 'daxstudio' }, @{ Name = 'greenshot' }, @{ Name = 'colourcontrastanalyser' }, @{ Name = 'paintdotnet' }
    ) {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot "Packagers\package-$Name.ps1"))
        $text.IndexOf('Assert-DownloadAvailable') | Should -BeGreaterThan 0
        $text.IndexOf('Assert-DownloadAvailable') | Should -BeLessThan $text.IndexOf('Invoke-DownloadWithRetry -Url')
    }

    It 'reads the Paint.NET installer from the asset list of the release' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-paintdotnet.ps1'))
        $text | Should -Match 'Select-GitHubReleaseWithAsset'
        $text | Should -Not -Match 'releases/latest'
        $text | Should -Not -Match 'paint\.net\.\$\{version\}'
    }
}
