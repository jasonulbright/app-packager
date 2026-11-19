#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# The Authenticode status of a downloaded installer is checked and recorded (Q-70, step B).

BeforeDiscovery {
    # A file with a signature inside the file (a catalog-signed system file reports NotSigned when it is changed).
    $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    $script:SignedOk = [bool]($pwsh -and ((Get-AuthenticodeSignature -LiteralPath $pwsh).Status -eq 'Valid'))
}

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $script:Root 'Packagers\AppPackagerCommon.psd1') -Force
    $script:Work = Join-Path ([System.IO.Path]::GetTempPath()) ('dl-sig-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory $script:Work | Out-Null
    $script:Signed = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    function New-Tampered([string]$Name) {
        $file = Join-Path $script:Work $Name
        $bytes = [System.IO.File]::ReadAllBytes($script:Signed)
        $bytes[[int]($bytes.Length / 2)] = $bytes[[int]($bytes.Length / 2)] -bxor 0xFF
        [System.IO.File]::WriteAllBytes($file, $bytes)
        $file
    }
}

AfterAll { Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Test-DownloadedInstallerSignature' {
    It 'accepts a valid signature, records the signer and the thumbprint' -Skip:(-not $script:SignedOk) {
        $file = Join-Path $script:Work 'valid.exe'; Copy-Item -LiteralPath $script:Signed -Destination $file
        { Test-DownloadedInstallerSignature -Path $file -Quiet } | Should -Not -Throw
        $record = Get-RecordedInstallerSignature -FileName 'valid.exe'
        $record.Status | Should -Be 'Valid'
        $record.Signer | Should -Not -BeNullOrEmpty
        $record.Thumbprint | Should -Match '^[0-9A-F]{40}$'
    }
    It 'warns about an unsigned installer and goes on, keeping the file' {
        $file = Join-Path $script:Work 'unsigned.msi'; Set-Content -LiteralPath $file -Value 'not signed'
        { Test-DownloadedInstallerSignature -Path $file -Quiet } | Should -Not -Throw
        Test-Path -LiteralPath $file | Should -BeTrue
        (Get-RecordedInstallerSignature -FileName 'unsigned.msi').Signer | Should -Be ''
    }
    It 'removes a signed file that was changed after signing and says nothing was staged' -Skip:(-not $script:SignedOk) {
        $file = New-Tampered 'changed.exe'
        { Test-DownloadedInstallerSignature -Path $file -Quiet } | Should -Throw -ExpectedMessage '*Authenticode signature that does not pass (status: HashMismatch)*removed and nothing was staged*'
        Test-Path -LiteralPath $file | Should -BeFalse
    }
    It 'leaves a file that is not an installer alone' {
        $file = Join-Path $script:Work 'data.zip'; Set-Content -LiteralPath $file -Value 'zip'
        { Test-DownloadedInstallerSignature -Path $file -Quiet } | Should -Not -Throw
        Get-RecordedInstallerSignature -FileName 'data.zip' | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-DownloadWithRetry applies the check' {
    BeforeAll {
        $script:Port = Get-Random -Minimum 20000 -Maximum 40000
        $script:Listener = New-Object System.Net.HttpListener
        $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
        $script:Listener.Start()
        $script:Ps = [powershell]::Create()
        [void]$script:Ps.AddScript({
                param($listener, $good, $bad)
                while ($listener.IsListening) {
                    try { $ctx = $listener.GetContext() } catch { break }
                    $bytes = if ($ctx.Request.Url.AbsolutePath -eq '/good.exe') { [System.IO.File]::ReadAllBytes($good) } else { [System.IO.File]::ReadAllBytes($bad) }
                    $ctx.Response.StatusCode = 200
                    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $ctx.Response.Close()
                }
            }).AddArgument($script:Listener).AddArgument($script:Signed).AddArgument((New-Tampered 'served-bad.exe'))
        $script:Handle = $script:Ps.BeginInvoke()
    }
    AfterAll {
        try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
        try { $script:Ps.Dispose() } catch { $null = $_ }
    }
    It 'keeps a validly signed download' -Skip:(-not $script:SignedOk) {
        $out = Join-Path $script:Work 'dl-good.exe'
        { Invoke-DownloadWithRetry -Url "http://localhost:$($script:Port)/good.exe" -OutFile $out -Quiet } | Should -Not -Throw
        Test-Path -LiteralPath $out | Should -BeTrue
        (Get-RecordedInstallerSignature -FileName 'dl-good.exe').Status | Should -Be 'Valid'
    }
    It 'stops on a download that was changed after signing and removes it' -Skip:(-not $script:SignedOk) {
        $out = Join-Path $script:Work 'dl-bad.exe'
        { Invoke-DownloadWithRetry -Url "http://localhost:$($script:Port)/bad.exe" -OutFile $out -Quiet } | Should -Throw -ExpectedMessage '*HashMismatch*nothing was staged*'
        Test-Path -LiteralPath $out | Should -BeFalse
    }
}

Describe 'The stage manifest records the signature' {
    It 'Write-StageManifest adds InstallerSignature from the record before the icon and the hashes' {
        $text = Get-Content -LiteralPath (Join-Path $script:Root 'Packagers\AppPackagerCommon.psm1') -Raw
        $i = $text.IndexOf("`$ManifestData['InstallerSignature'] = `$recordedSignature")
        $i | Should -BeGreaterThan 0
        $i | Should -BeLessThan $text.IndexOf('Add-StageIcon -StageRoot $stageRoot -ManifestData $ManifestData -PackagerScriptPath $PackagerScriptPath')
    }
}