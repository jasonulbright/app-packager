#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# The digest that GitHub lists for a release asset is checked when a packager downloads the asset (Q-70, step A).

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $script:Root 'Packagers\AppPackagerCommon.psd1') -Force
    $script:Work = Join-Path ([System.IO.Path]::GetTempPath()) ('asset-digest-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory $script:Work | Out-Null
    $script:Bytes = [byte[]](1..250)
    $script:Hash = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash($script:Bytes)).Replace('-', '')
    # A local server that serves the bytes at /asset.bin.
    $script:Port = Get-Random -Minimum 20000 -Maximum 40000
    $script:Listener = New-Object System.Net.HttpListener
    $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
    $script:Listener.Start()
    $script:Ps = [powershell]::Create()
    [void]$script:Ps.AddScript({
            param($listener, $bytes)
            while ($listener.IsListening) {
                try { $ctx = $listener.GetContext() } catch { break }
                $ctx.Response.StatusCode = 200
                $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                $ctx.Response.Close()
            }
        }).AddArgument($script:Listener).AddArgument($script:Bytes)
    $script:Handle = $script:Ps.BeginInvoke()
    $script:Url = "http://localhost:$($script:Port)/asset.bin"
    function New-Release([string]$Digest) { [pscustomobject]@{ assets = @([pscustomobject]@{ name = 'asset.bin'; browser_download_url = $script:Url; digest = $Digest }) } }
}

AfterAll {
    try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
    try { $script:Ps.Dispose() } catch { $null = $_ }
    Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Invoke-DownloadWithRetry checks the digest of a registered release asset' {
    It 'keeps a file that matches the digest' {
        Register-ReleaseAssetDigest -Release (New-Release ('sha256:' + $script:Hash.ToLowerInvariant()))
        $out = Join-Path $script:Work 'ok.bin'
        { Invoke-DownloadWithRetry -Url $script:Url -OutFile $out -Quiet } | Should -Not -Throw
        Test-Path -LiteralPath $out | Should -BeTrue
    }
    It 'removes a file that does not match and says nothing was staged' {
        Register-ReleaseAssetDigest -Release (New-Release ('sha256:' + ('0' * 64)))
        $out = Join-Path $script:Work 'bad.bin'
        { Invoke-DownloadWithRetry -Url $script:Url -OutFile $out -Quiet } |
            Should -Throw -ExpectedMessage '*does not match the SHA-256 digest that GitHub lists*removed and nothing was staged*'
        Test-Path -LiteralPath $out | Should -BeFalse
    }
    It 'goes on with a warning when the release lists no digest' {
        Register-ReleaseAssetDigest -Release (New-Release '')
        $out = Join-Path $script:Work 'nodigest.bin'
        { Invoke-DownloadWithRetry -Url $script:Url -OutFile $out } | Should -Not -Throw
        Test-Path -LiteralPath $out | Should -BeTrue
    }
    It 'does not check a download that is not a registered release asset' {
        Register-ReleaseAssetDigest -Release $null
        $other = "http://localhost:$($script:Port)/other.bin"
        $out = Join-Path $script:Work 'other.bin'
        { Invoke-DownloadWithRetry -Url $other -OutFile $out -Quiet } | Should -Not -Throw
    }
}

Describe 'The release selectors register the digests of the release they choose' {
    It 'Select-GitHubReleaseWithAsset, ByAsset and Json call Register-ReleaseAssetDigest' {
        $text = Get-Content -LiteralPath (Join-Path $script:Root 'Packagers\AppPackagerCommon.psm1') -Raw
        ([regex]::Matches($text, 'Register-ReleaseAssetDigest -Release')).Count | Should -BeGreaterOrEqual 5
    }
    It 'a release chosen by Select-GitHubReleaseByAsset makes its asset checked' {
        $rel = [pscustomobject]@{ tag_name = 'v1'; draft = $false; prerelease = $false; published_at = '2026-01-01T00:00:00Z'
            assets = @([pscustomobject]@{ name = 'asset.bin'; browser_download_url = $script:Url; digest = ('sha256:' + ('1' * 64)) }) }
        $null = Select-GitHubReleaseByAsset -Releases @($rel) -AssetFilter { $_.name -eq 'asset.bin' }
        $out = Join-Path $script:Work 'selected.bin'
        { Invoke-DownloadWithRetry -Url $script:Url -OutFile $out -Quiet } | Should -Throw -ExpectedMessage '*does not match the SHA-256 digest*'
    }
}