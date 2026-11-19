#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Installer verification (Q-70): the Authenticode check and the digest check of a downloaded installer.

BeforeDiscovery {
    $script:SignedOk = ((Get-AuthenticodeSignature -LiteralPath (Join-Path $env:SystemRoot 'System32\notepad.exe')).Status -eq 'Valid')
}

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $script:Root 'Packagers\AppPackagerCommon.psd1') -Force
    $script:Work = Join-Path ([System.IO.Path]::GetTempPath()) ('installer-verify-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory $script:Work | Out-Null
    # A file that Windows signs, and its signer name as the check reads it.
    $script:Signed = Join-Path $env:SystemRoot 'System32\notepad.exe'
    $sig = Get-AuthenticodeSignature -LiteralPath $script:Signed
    $script:SignedOk = ($sig.Status -eq 'Valid')
    $script:Signer = if ($script:SignedOk -and $sig.SignerCertificate.Subject -match 'CN=(?:"(?<q>[^"]+)"|(?<n>[^,]+))') { $(if ($Matches['q']) { $Matches['q'] } else { $Matches['n'] }).Trim() } else { '' }
}

AfterAll { Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Assert-InstallerSignature' {
    It 'returns the signer of a validly signed file with the expected signer' -Skip:(-not $script:SignedOk) {
        (Assert-InstallerSignature -Path $script:Signed -ExpectedSigner $script:Signer) | Should -Be $script:Signer
    }
    It 'refuses another signer and names both' -Skip:(-not $script:SignedOk) {
        { Assert-InstallerSignature -Path $script:Signed -ExpectedSigner 'Nobody Corporation' } |
            Should -Throw -ExpectedMessage "*signed by '$($script:Signer)', not by 'Nobody Corporation'*Nothing was staged*"
    }
    It 'refuses an unsigned file, says the status, and keeps the file unless asked' {
        $file = Join-Path $script:Work 'unsigned.msi'
        Set-Content -LiteralPath $file -Value 'not an installer'
        { Assert-InstallerSignature -Path $file -ExpectedSigner 'Microsoft Corporation' } | Should -Throw -ExpectedMessage '*no valid Authenticode signature (status: *)*Nothing was staged*'
        Test-Path -LiteralPath $file | Should -BeTrue
    }
    It 'deletes the file on failure with -RemoveOnFailure, so the next run downloads it again' {
        $file = Join-Path $script:Work 'unsigned2.msi'
        Set-Content -LiteralPath $file -Value 'not an installer'
        { Assert-InstallerSignature -Path $file -ExpectedSigner 'Microsoft Corporation' -RemoveOnFailure } | Should -Throw
        Test-Path -LiteralPath $file | Should -BeFalse
    }
    It 'refuses a signed file that was changed after signing' -Skip:(-not $script:SignedOk) {
        $file = Join-Path $script:Work 'tampered.exe'
        Copy-Item -LiteralPath $script:Signed -Destination $file
        $bytes = [System.IO.File]::ReadAllBytes($file); $bytes[200] = $bytes[200] -bxor 0xFF; [System.IO.File]::WriteAllBytes($file, $bytes)
        { Assert-InstallerSignature -Path $file -ExpectedSigner $script:Signer } | Should -Throw -ExpectedMessage '*no valid Authenticode signature*'
    }
    It 'says so when the file is missing' {
        { Assert-InstallerSignature -Path (Join-Path $script:Work 'missing.msi') -ExpectedSigner 'X' } | Should -Throw -ExpectedMessage '*is missing*Nothing was staged*'
    }
}

Describe 'Assert-InstallerDigest' {
    BeforeAll {
        $script:Data = Join-Path $script:Work 'data.bin'
        [System.IO.File]::WriteAllBytes($script:Data, [byte[]](1..200))
        $script:Hash = (Get-FileHash -LiteralPath $script:Data -Algorithm SHA256).Hash
    }
    It 'accepts the digest, with or without the sha256: prefix and in any case' {
        (Assert-InstallerDigest -Path $script:Data -Sha256 $script:Hash) | Should -Be $script:Hash
        (Assert-InstallerDigest -Path $script:Data -Sha256 ('sha256:' + $script:Hash.ToLowerInvariant())) | Should -Be $script:Hash
    }
    It 'refuses a different digest and removes the file on request' {
        $copy = Join-Path $script:Work 'data-copy.bin'; Copy-Item -LiteralPath $script:Data -Destination $copy
        { Assert-InstallerDigest -Path $copy -Sha256 ('0' * 64) -RemoveOnFailure } | Should -Throw -ExpectedMessage '*does not match the SHA-256 digest*Nothing was staged*'
        Test-Path -LiteralPath $copy | Should -BeFalse
    }
}

Describe 'The PowerShell 7 packager checks the signature before it stages' {
    It 'asks for the signer Microsoft Corporation before it reads the MSI' {
        $text = Get-Content -LiteralPath (Join-Path $script:Root 'Packagers\package-powershell7.ps1') -Raw
        $text | Should -Match "Assert-InstallerSignature -Path \`$localMsi -ExpectedSigner 'Microsoft Corporation' -RemoveOnFailure"
        $text.IndexOf('Assert-InstallerSignature') | Should -BeLessThan $text.IndexOf('Get-MsiPropertyMap -MsiPath')
    }
}