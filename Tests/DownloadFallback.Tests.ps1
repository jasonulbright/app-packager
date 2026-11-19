#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    A vendor that serves one file under two addresses (Tableau: esdalt, then tssoftware): the second address is tried when
    the first fails for any reason, and the run fails only when both fail.

.DESCRIPTION
    A local HTTP server plays the host: a good installer (starts with MZ), 403, 404, 500, a web page with status 200 (a sign-in
    page) and a body that is no installer. A closed port plays a connection failure. curl.exe does the transfers, as in the packagers.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Work = Join-Path ([System.IO.Path]::GetTempPath()) ('dlfallback-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:Work | Out-Null
    $script:Port = Get-Random -Minimum 20000 -Maximum 40000
    $script:State = [hashtable]::Synchronized(@{ Requests = [System.Collections.ArrayList]::new() })
    $script:Listener = New-Object System.Net.HttpListener
    $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
    $script:Listener.Start()
    $script:Ps = [powershell]::Create()
    [void]$script:Ps.AddScript({
            param($listener, $state)
            $good = New-Object byte[] 4096
            $good[0] = 0x4D; $good[1] = 0x5A
            for ($i = 2; $i -lt $good.Length; $i++) { $good[$i] = [byte](($i * 7) % 251) }
            while ($listener.IsListening) {
                try { $ctx = $listener.GetContext() } catch { break }
                $res = $ctx.Response
                $path = $ctx.Request.Url.AbsolutePath
                [void]$state.Requests.Add(('{0} {1}' -f $ctx.Request.HttpMethod, $path))
                try {
                    switch ($path) {
                        '/good.bin' { $res.StatusCode = 200; $res.OutputStream.Write($good, 0, $good.Length) }
                        '/good2.bin' { $res.StatusCode = 200; $res.OutputStream.Write($good, 0, $good.Length) }
                        '/forbidden' { $res.StatusCode = 403 }
                        '/missing' { $res.StatusCode = 404 }
                        '/error' { $res.StatusCode = 500 }
                        '/signin' { $res.StatusCode = 200; $res.ContentType = 'text/html'; $b = [Text.Encoding]::ASCII.GetBytes('<html><body>Sign in</body></html>'); $res.OutputStream.Write($b, 0, $b.Length) }
                        '/text' { $res.StatusCode = 200; $b = [Text.Encoding]::ASCII.GetBytes('not an installer'); $res.OutputStream.Write($b, 0, $b.Length) }
                        default { $res.StatusCode = 500 }
                    }
                    $res.Close()
                }
                catch { try { $res.Abort() } catch { $null = $_ } }
            }
        }).AddArgument($script:Listener).AddArgument($script:State)
    $script:Handle = $script:Ps.BeginInvoke()
    $script:Base = "http://localhost:$($script:Port)"
    $script:Dead = "http://localhost:$(Get-Random -Minimum 41000 -Maximum 45000)"
}

AfterAll {
    try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
    try { $script:Ps.Dispose() } catch { $null = $_ }
    if ($script:Work -and (Test-Path -LiteralPath $script:Work)) { Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Invoke-DownloadFromAddresses' {
    BeforeEach {
        $script:State.Requests.Clear()
        $global:FallbackLog = [System.Collections.ArrayList]::new()
        Mock -ModuleName AppPackagerCommon Write-Log { [void]$global:FallbackLog.Add(('{0}|{1}' -f $Level, $Message)) }
        $script:Out = Join-Path $script:Work ('out-' + [guid]::NewGuid().ToString('N') + '.bin')
    }

    It 'takes the first address when it gives an installer, and never asks the second' {
        $used = Invoke-DownloadFromAddresses -Urls @("$($script:Base)/good.bin", "$($script:Base)/good2.bin") -OutFile $script:Out
        $used | Should -Be "$($script:Base)/good.bin"
        (Get-Item -LiteralPath $script:Out).Length | Should -Be 4096
        @($script:State.Requests | Where-Object { $_ -match 'good2' }).Count | Should -Be 0
        ($global:FallbackLog -join "`n") | Should -Match 'Download address used\s+: .*/good\.bin'
    }

    It 'tries the second address after a 403 and says why the first was skipped' {
        $used = Invoke-DownloadFromAddresses -Urls @("$($script:Base)/forbidden", "$($script:Base)/good.bin") -OutFile $script:Out
        $used | Should -Be "$($script:Base)/good.bin"
        (Get-Item -LiteralPath $script:Out).Length | Should -Be 4096
        ($global:FallbackLog -join "`n") | Should -Match 'Address skipped.*/forbidden.*HTTP 403'
    }

    It 'tries the second address after a 404' {
        Invoke-DownloadFromAddresses -Urls @("$($script:Base)/missing", "$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
        ($global:FallbackLog -join "`n") | Should -Match 'HTTP 404'
    }

    It 'tries the second address after another HTTP error (500)' {
        Invoke-DownloadFromAddresses -Urls @("$($script:Base)/error", "$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
    }

    It 'tries the second address after a connection failure' {
        Invoke-DownloadFromAddresses -Urls @("$($script:Dead)/x.bin", "$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
        ($global:FallbackLog -join "`n") | Should -Match 'no HTTP answer'
    }

    It 'tries the second address when the first gives a web page (a sign-in page) with status 200, and keeps no page' {
        Invoke-DownloadFromAddresses -Urls @("$($script:Base)/signin", "$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
        $bytes = [System.IO.File]::ReadAllBytes($script:Out); ($bytes[0], $bytes[1]) -join ',' | Should -Be '77,90'
        ($global:FallbackLog -join "`n") | Should -Match 'not an installer'
    }

    It 'tries the second address when the first gives a body that is no installer' {
        Invoke-DownloadFromAddresses -Urls @("$($script:Base)/text", "$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
    }

    It 'fails only when both addresses fail, and the message holds both addresses and both reasons' {
        $thrown = $null
        try { Invoke-DownloadFromAddresses -Urls @("$($script:Base)/forbidden", "$($script:Base)/missing") -OutFile $script:Out } catch { $thrown = $_.Exception.Message }
        $thrown | Should -Match 'could not be downloaded from any of its addresses'
        $thrown | Should -Match '/forbidden.*HTTP 403'
        $thrown | Should -Match '/missing.*HTTP 404'
        Test-Path -LiteralPath $script:Out | Should -BeFalse
    }

    It 'leaves no partial file when the first address fails and the second one fails too' {
        { Invoke-DownloadFromAddresses -Urls @("$($script:Base)/signin", "$($script:Base)/text") -OutFile $script:Out } | Should -Throw
        Test-Path -LiteralPath $script:Out | Should -BeFalse
    }

    It 'with one address behaves as a plain download and names that address in the failure' {
        { Invoke-DownloadFromAddresses -Urls @("$($script:Base)/forbidden") -OutFile $script:Out } | Should -Throw '*/forbidden*HTTP 403*'
        Invoke-DownloadFromAddresses -Urls @("$($script:Base)/good.bin") -OutFile $script:Out | Should -Be "$($script:Base)/good.bin"
    }
}

Describe 'The Tableau packagers use the fallback' {
    It 'download through both addresses in <_>' -ForEach @('package-tableaudesktop.ps1', 'package-tableaureader.ps1', 'package-tableauprep.ps1') {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot "..\Packagers\$_") -Raw
        $text | Should -Match 'Invoke-DownloadFromAddresses -Urls @\(\$release\.DownloadUrl, \$release\.FallbackUrl\) -OutFile \$localExe'
        $text | Should -Not -Match 'Select-AvailableDownloadUrl'
        $text | Should -Match '\$DownloadUrlTemplate = "https://downloads\.tableau\.com/esdalt/'
        $text | Should -Match '\$FallbackUrlTemplate = "https://downloads\.tableau\.com/tssoftware/'
    }
}
