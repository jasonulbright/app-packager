#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Invoke-ResumableDownload: a large installer that resumes, logs its progress and stops on a status that no retry changes.

.DESCRIPTION
    A local HTTP server serves one file with Range support. The first plain request loses its connection halfway; the retry
    must continue at the byte where it stopped. A 403 and a 404 stop at once. curl.exe does the transfer, as in the packagers.
#>

BeforeDiscovery {
    $script:CurlOk = [bool](Get-Command curl.exe -ErrorAction SilentlyContinue)
}

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Work = Join-Path ([System.IO.Path]::GetTempPath()) ('resumable-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:Work | Out-Null
    $script:Source = Join-Path $script:Work 'source.bin'
    $rng = New-Object System.Random 7
    $bytes = New-Object byte[] (3MB)
    $rng.NextBytes($bytes)
    [System.IO.File]::WriteAllBytes($script:Source, $bytes)
    $script:SourceHash = (Get-FileHash -LiteralPath $script:Source -Algorithm SHA256).Hash

    $script:Port = Get-Random -Minimum 20000 -Maximum 40000
    $script:State = [hashtable]::Synchronized(@{ Requests = [System.Collections.ArrayList]::new(); Dropped = $false; Slow = $false })
    $script:Listener = New-Object System.Net.HttpListener
    $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
    $script:Listener.Start()
    $script:Ps = [powershell]::Create()
    [void]$script:Ps.AddScript({
            param($listener, $path, $state)
            $data = [System.IO.File]::ReadAllBytes($path)
            while ($listener.IsListening) {
                try { $ctx = $listener.GetContext() } catch { break }
                $req = $ctx.Request; $res = $ctx.Response
                $range = $req.Headers['Range']
                [void]$state.Requests.Add(('{0} {1} {2}' -f $req.HttpMethod, $req.Url.AbsolutePath, $range))
                try {
                    if ($req.Url.AbsolutePath -eq '/forbidden') { $res.StatusCode = 403; $res.Close(); continue }
                    if ($req.Url.AbsolutePath -eq '/gone') { $res.StatusCode = 404; $res.Close(); continue }
                    $res.AddHeader('Accept-Ranges', 'bytes')
                    $from = 0
                    if ($range -match '^bytes=(\d+)-') { $from = [int]$Matches[1] }
                    if ($req.HttpMethod -eq 'HEAD') { $res.StatusCode = 200; $res.ContentLength64 = $data.Length; $res.Close(); continue }
                    if ($from -ge $data.Length) { $res.StatusCode = 416; $res.Close(); continue }
                    $length = $data.Length - $from
                    if ($range) { $res.StatusCode = 206; $res.AddHeader('Content-Range', ('bytes {0}-{1}/{2}' -f $from, ($data.Length - 1), $data.Length)) } else { $res.StatusCode = 200 }
                    $res.ContentType = 'application/octet-stream'
                    $res.ContentLength64 = $length
                    if (-not $range -and -not $state.Dropped -and $req.Url.AbsolutePath -eq '/drop.bin') {
                        $state.Dropped = $true
                        $res.OutputStream.Write($data, 0, [int]($data.Length * 0.4))
                        $res.OutputStream.Flush()
                        $res.Abort()
                        continue
                    }
                    if ($state.Slow) {
                        $chunk = 256KB; $sent = 0
                        while ($sent -lt $length) { $n = [Math]::Min($chunk, $length - $sent); $res.OutputStream.Write($data, $from + $sent, $n); $res.OutputStream.Flush(); $sent += $n; Start-Sleep -Milliseconds 400 }
                    }
                    else { $res.OutputStream.Write($data, $from, $length) }
                    $res.Close()
                }
                catch { try { $res.Abort() } catch { $null = $_ } }
            }
        }).AddArgument($script:Listener).AddArgument($script:Source).AddArgument($script:State)
    $script:Handle = $script:Ps.BeginInvoke()
    $script:Base = "http://localhost:$($script:Port)"
}

AfterAll {
    try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
    try { $script:Ps.Dispose() } catch { $null = $_ }
    if ($script:Work -and (Test-Path -LiteralPath $script:Work)) { Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Invoke-ResumableDownload' {
    BeforeEach {
        $script:State.Requests.Clear()
        $script:State.Dropped = $false
        $script:State.Slow = $false
        $global:ResumableLines = [System.Collections.ArrayList]::new()
        Mock -ModuleName AppPackagerCommon Write-Log { [void]$global:ResumableLines.Add([string]$Message) }
    }

    It 'continues at the byte where a lost connection stopped, and the file is whole' -Skip:(-not $script:CurlOk) {
        $out = Join-Path $script:Work 'resumed.bin'
        try { Invoke-ResumableDownload -Url "$($script:Base)/drop.bin" -OutFile $out -RetryDelaySec 1 -ProgressSeconds 1 } catch { throw ($_.Exception.Message + ' LOG: ' + ($global:ResumableLines -join ' | ')) }
        (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash | Should -Be $script:SourceHash
        Test-Path -LiteralPath ($out + '.part') | Should -BeFalse
        @($script:State.Requests | Where-Object { $_ -match '^GET /drop\.bin bytes=\d+-' }).Count | Should -BeGreaterThan 0
        ($global:ResumableLines -join "`n") | Should -Match 'continuing at'
    }

    It 'continues a .part file that an earlier run left' -Skip:(-not $script:CurlOk) {
        $out = Join-Path $script:Work 'later.bin'
        $all = [System.IO.File]::ReadAllBytes($script:Source)
        [System.IO.File]::WriteAllBytes($out + '.part', $all[0..(1MB - 1)])
        Invoke-ResumableDownload -Url "$($script:Base)/file.bin" -OutFile $out -ProgressSeconds 1
        (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash | Should -Be $script:SourceHash
        @($script:State.Requests | Where-Object { $_ -match '^GET /file\.bin bytes=1048576-' }).Count | Should -Be 1
    }

    It 'logs the announced size and the progress with size and speed' -Skip:(-not $script:CurlOk) {
        $script:State.Slow = $true
        $out = Join-Path $script:Work 'slow.bin'
        Invoke-ResumableDownload -Url "$($script:Base)/file.bin" -OutFile $out -ProgressSeconds 1
        $text = $global:ResumableLines -join "`n"
        $text | Should -Match 'Announced size\s+: 3 MB'
        $text | Should -Match 'Downloaded\s+: \d+ MB \(\d+%\), [\d.,]+ MB/s'
        (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash | Should -Be $script:SourceHash
    }

    It 'stops at once on a 403 with the status, and leaves no finished file' -Skip:(-not $script:CurlOk) {
        $out = Join-Path $script:Work 'forbidden.bin'
        { Invoke-ResumableDownload -Url "$($script:Base)/forbidden" -OutFile $out -RetryDelaySec 1 -ProgressSeconds 1 } | Should -Throw '*HTTP 403*No retry can change that*'
        Test-Path -LiteralPath $out | Should -BeFalse
        @($script:State.Requests | Where-Object { $_ -match '^GET /forbidden' }).Count | Should -Be 1
    }

    It 'stops at once on a 404' -Skip:(-not $script:CurlOk) {
        $out = Join-Path $script:Work 'gone.bin'
        { Invoke-ResumableDownload -Url "$($script:Base)/gone" -OutFile $out -RetryDelaySec 1 -ProgressSeconds 1 } | Should -Throw '*HTTP 404*'
        @($script:State.Requests | Where-Object { $_ -match '^GET /gone' }).Count | Should -Be 1
    }

    It 'keeps the partial file when every attempt fails and says that the next run continues it' -Skip:(-not $script:CurlOk) {
        $out = Join-Path $script:Work 'unreachable.bin'
        $closed = Get-Random -Minimum 41000 -Maximum 45000
        { Invoke-ResumableDownload -Url "http://localhost:$closed/none.bin" -OutFile $out -MaxAttempts 2 -RetryDelaySec 1 -ProgressSeconds 1 } | Should -Throw '*next run continues it*'
        Test-Path -LiteralPath $out | Should -BeFalse
    }
}
