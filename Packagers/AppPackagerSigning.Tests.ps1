#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for the AppPackagerSigning service.

.DESCRIPTION
    Creates throwaway self-signed code-signing certificates in
    Cert:\CurrentUser\My and removes them in AfterAll. No trusted-root or
    trusted-publisher store is touched: adding a root there raises an
    interactive Windows confirmation dialog, and endpoint trust is not what
    these tests prove.

    Signature assertions therefore use SignatureIntact (hash and signature
    block still match the content) rather than the Valid/trusted status. A
    self-signed certificate that this host does not trust reports
    Status 'UnknownError' with an untrusted-root message and SignatureIntact
    true; a changed byte reports 'HashMismatch' and SignatureIntact false,
    which is what the transport and tamper proofs need.

.EXAMPLE
    Invoke-Pester .\AppPackagerSigning.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\AppPackagerSigning.psd1" -Force

    $script:Root = Join-Path ([System.IO.Path]::GetTempPath()) ('apsign-' + [guid]::NewGuid().ToString('N').Substring(0, 10))
    New-Item -ItemType Directory -Path $script:Root -Force | Out-Null

    $script:Cert = New-SelfSignedCertificate -Type CodeSigningCert `
        -Subject 'CN=AppPackager Unit Test Signing' `
        -CertStoreLocation Cert:\CurrentUser\My `
        -NotAfter (Get-Date).AddDays(1)
    $script:Thumbprint = $script:Cert.Thumbprint.ToUpperInvariant()

    # A second certificate imported public-key-only, so Cert:\CurrentUser\My
    # holds an entry with no usable private key.
    $script:PublicOnlyThumbprint = ''
    try {
        $keyless = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject 'CN=AppPackager Unit Test Keyless' `
            -CertStoreLocation Cert:\CurrentUser\My `
            -NotAfter (Get-Date).AddDays(1) -ErrorAction Stop
        $keylessCer = Join-Path $script:Root 'keyless.cer'
        [System.IO.File]::WriteAllBytes($keylessCer, $keyless.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        Remove-Item -LiteralPath "Cert:\CurrentUser\My\$($keyless.Thumbprint)" -Force -DeleteKey -ErrorAction SilentlyContinue
        $imported = Import-Certificate -FilePath $keylessCer -CertStoreLocation Cert:\CurrentUser\My -ErrorAction Stop
        $script:PublicOnlyThumbprint = $imported.Thumbprint.ToUpperInvariant()
    }
    catch { $script:PublicOnlyThumbprint = '' }

    $script:ExpiredThumbprint = ''
    try {
        $expired = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject 'CN=AppPackager Unit Test Expired' `
            -CertStoreLocation Cert:\CurrentUser\My `
            -NotBefore (Get-Date).AddDays(-30) -NotAfter (Get-Date).AddDays(-1) -ErrorAction Stop
        $script:ExpiredThumbprint = $expired.Thumbprint.ToUpperInvariant()
    }
    catch { $script:ExpiredThumbprint = '' }

    $script:Policy = {
        param([bool]$SD, [bool]$SR, [bool]$SP, [bool]$RD = $false, [bool]$RR = $false, [bool]$RP = $false,
              [string]$Timestamp = '', [bool]$TimestampRequired = $false, [string]$Thumb = $script:Thumbprint)
        [pscustomobject]@{
            SignDetection = $SD; SignRequirements = $SR; SignDeployment = $SP
            RequireDetection = $RD; RequireRequirements = $RR; RequireDeployment = $RP
            CertificateThumbprint = $Thumb; StoreLocation = 'CurrentUser'
            TimestampServer = $Timestamp; TimestampRequired = $TimestampRequired; HashAlgorithm = 'SHA256'
        }
    }

    $script:NewFixture = {
        param([string]$Name, [bool]$SignedLaunchers)
        $stage = Join-Path $script:Root $Name
        New-Item -ItemType Directory -Path (Join-Path $stage 'scripts\requirements') -Force | Out-Null
        $detect = "if (Test-Path 'C:\Program Files\Fixture\app.exe') { Write-Output 'Detected' }`r`nexit 0"
        Set-Content -LiteralPath (Join-Path $stage 'scripts\detect.ps1') -Value $detect -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stage 'scripts\requirements\r1.ps1') -Value "Write-Output 'true'`r`nexit 0" -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stage 'install.ps1') -Value "Import-Module `"`$PSScriptRoot\helper.psm1`"`r`nexit 0" -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stage 'uninstall.ps1') -Value 'exit 0' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stage 'helper.psm1') -Value "function Get-FixtureState { 'ok' }" -Encoding ASCII
        foreach ($entry in 'install', 'uninstall') {
            $launcher = New-DeploymentLauncherCommand -Script "$entry.ps1" -Signed $SignedLaunchers
            Set-Content -LiteralPath (Join-Path $stage "$entry.bat") -Value $launcher.BatBody -Encoding ASCII
        }
        [pscustomobject]@{
            StageRoot = $stage
            Manifest  = @{
                Detection            = @{ Type = 'Script'; ScriptText = $detect }
                Requirements         = @(@{ RuleId = 'r1'; ScriptText = "Write-Output 'true'`r`nexit 0" })
                InstallCommandLine   = (New-DeploymentLauncherCommand -Script 'install.bat' -Signed $SignedLaunchers).CommandLine
                UninstallCommandLine = (New-DeploymentLauncherCommand -Script 'uninstall.bat' -Signed $SignedLaunchers).CommandLine
            }
        }
    }
}

AfterAll {
    foreach ($thumb in @($script:Thumbprint, $script:PublicOnlyThumbprint, $script:ExpiredThumbprint)) {
        if ($thumb) { Remove-Item -LiteralPath "Cert:\CurrentUser\My\$thumb" -Force -ErrorAction SilentlyContinue }
    }
    if ($script:Root -and (Test-Path $script:Root)) {
        Remove-Item -LiteralPath $script:Root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Module surface' {
    It 'exports every contracted function' {
        $expected = @(
            'Get-SigningPolicy', 'Get-CodeSigningCertificateCandidates', 'Resolve-SigningCertificate',
            'Test-SigningConfiguration', 'Invoke-ScriptSigning', 'Test-ScriptSignature',
            'Test-ScriptSignatureBytes', 'Get-SignedScriptRepresentation', 'Test-SignedScriptRoundTrip',
            'New-DeploymentLauncherCommand', 'Test-DeploymentLauncherChain', 'Invoke-CategorySigning'
        )
        $actual = (Get-Command -Module AppPackagerSigning).Name
        foreach ($name in $expected) { $actual | Should -Contain $name }
    }

    It 'reports the ConfigMgr detection script size limit' {
        Get-ConfigMgrDetectionScriptMaxBytes | Should -Be 32768
    }
}

Describe 'Get-SigningPolicy' {
    It 'defaults every switch off' {
        $p = Get-SigningPolicy
        $p.SignDetection | Should -BeFalse
        $p.SignRequirements | Should -BeFalse
        $p.SignDeployment | Should -BeFalse
        $p.RequireDetection | Should -BeFalse
        $p.RequireRequirements | Should -BeFalse
        $p.RequireDeployment | Should -BeFalse
        $p.HashAlgorithm | Should -Be 'SHA256'
        $p.StoreLocation | Should -Be 'CurrentUser'
    }

    It 'normalizes a policy object passed directly' {
        $p = Get-SigningPolicy -Policy ([pscustomobject]@{ SignDetection = 'true'; StoreLocation = 'localmachine'; CertificateThumbprint = 'aa bb:cc' })
        $p.SignDetection | Should -BeTrue
        $p.StoreLocation | Should -Be 'LocalMachine'
        $p.CertificateThumbprint | Should -Be 'AABBCC'
        $p.PolicyDigest.Length | Should -Be 64
    }

    It 'is idempotent when given its own output' {
        $once = Get-SigningPolicy -Policy ([pscustomobject]@{ SignDeployment = $true; CertificateThumbprint = $script:Thumbprint })
        (Get-SigningPolicy -Policy $once).PolicyDigest | Should -Be $once.PolicyDigest
    }

    It 'reads the child bridge environment variable' {
        $saved = $env:APP_PACKAGER_SIGNING
        try {
            $env:APP_PACKAGER_SIGNING = '{"SignDetection":true,"CertificateThumbprint":"aa bb:cc"}'
            $p = Get-SigningPolicy
            $p.SignDetection | Should -BeTrue
            $p.CertificateThumbprint | Should -Be 'AABBCC'
        }
        finally {
            if ($null -eq $saved) { Remove-Item Env:APP_PACKAGER_SIGNING -ErrorAction SilentlyContinue }
            else { $env:APP_PACKAGER_SIGNING = $saved }
        }
    }

    It 'reads the run snapshot block' {
        $snapshot = [pscustomobject]@{ SigningPolicy = [pscustomobject]@{ SignDeployment = $true; StoreLocation = 'localmachine' } }
        $p = Get-SigningPolicy -Snapshot $snapshot
        $p.SignDeployment | Should -BeTrue
        $p.StoreLocation | Should -Be 'LocalMachine'
    }

    It 'refuses an unsupported store location' {
        { Get-SigningPolicy -Policy ([pscustomobject]@{ StoreLocation = 'Machine' }) } | Should -Throw '*StoreLocation*'
    }

    It 'produces a digest that changes with the policy' {
        $a = Get-SigningPolicy -Policy ([pscustomobject]@{ SignDetection = $true })
        $b = Get-SigningPolicy -Policy ([pscustomobject]@{ SignDetection = $false })
        $a.PolicyDigest | Should -Not -Be $b.PolicyDigest
    }

    It 'never enables a switch from unparsable input' {
        (Get-SigningPolicy -Policy ([pscustomobject]@{ SignDeployment = 'maybe' })).SignDeployment | Should -BeFalse
    }
}

Describe 'Get-CodeSigningCertificateCandidates' {
    It 'lists the throwaway certificate as usable with an accessible key' {
        $c = @(Get-CodeSigningCertificateCandidates -StoreLocation CurrentUser | Where-Object { $_.Thumbprint -eq $script:Thumbprint })
        $c.Count | Should -Be 1
        $c[0].HasPrivateKey | Should -BeTrue
        $c[0].KeyAccessible | Should -BeTrue
        $c[0].Usable | Should -BeTrue
        $c[0].MayPrompt | Should -BeFalse
    }

    It 'marks an expired certificate unusable with a dated reason' {
        $c = @(Get-CodeSigningCertificateCandidates -StoreLocation CurrentUser | Where-Object { $_.Thumbprint -eq $script:ExpiredThumbprint })
        $c.Count | Should -Be 1
        $c[0].Usable | Should -BeFalse
        $c[0].Reason | Should -Match 'expired'
    }

    It 'marks a public-only certificate unusable' {
        $c = @(Get-CodeSigningCertificateCandidates -StoreLocation CurrentUser | Where-Object { $_.Thumbprint -eq $script:PublicOnlyThumbprint })
        $c.Count | Should -Be 1
        $c[0].HasPrivateKey | Should -BeFalse
        $c[0].Usable | Should -BeFalse
    }
}

Describe 'Resolve-SigningCertificate' {
    It 'selects the configured thumbprint' {
        (Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = $script:Thumbprint })).Thumbprint.ToUpperInvariant() |
            Should -Be $script:Thumbprint
    }

    It 'throws when no certificate is selected' {
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = '' }) } |
            Should -Throw '*SigningCertificateUnavailable: no certificate selected*'
    }

    It 'throws when the thumbprint is not in the store' {
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = ('0' * 40) }) } |
            Should -Throw '*was not found*'
    }

    It 'throws for an expired certificate' {
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = $script:ExpiredThumbprint }) } |
            Should -Throw '*expired on*'
    }

    It 'throws for a certificate without a private key' {
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = $script:PublicOnlyThumbprint }) } |
            Should -Throw '*no associated private key*'
    }

    It 'throws for an ambiguous configuration' {
        Mock -ModuleName AppPackagerSigning Get-ChildItem {
            $c = Get-Item "Cert:\CurrentUser\My\$($script:Thumbprint)"
            return @($c, $c)
        } -ParameterFilter { $Path -like 'Cert:*' }
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{ CertificateThumbprint = $script:Thumbprint }) } |
            Should -Throw '*ambiguous configuration*'
    }

    It 'never selects the first certificate when the thumbprint is blank' {
        @(Get-CodeSigningCertificateCandidates -StoreLocation CurrentUser).Count | Should -BeGreaterThan 0
        { Resolve-SigningCertificate -Policy ([pscustomobject]@{}) } | Should -Throw '*no certificate selected*'
    }
}

Describe 'Test-SigningConfiguration' {
    It 'signs and verifies a temporary file without leaving it behind' {
        $before = @(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter 'appackager-sign-*.ps1' -ErrorAction SilentlyContinue).Count
        $r = Test-SigningConfiguration -Policy (& $script:Policy $true $true $true) -TimeoutSeconds 120
        $r.Ok | Should -BeTrue
        $r.SignatureIntact | Should -BeTrue
        $r.Thumbprint | Should -Be $script:Thumbprint
        $r.TimestampVerified | Should -BeFalse
        $r.TimedOut | Should -BeFalse
        @(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter 'appackager-sign-*.ps1' -ErrorAction SilentlyContinue).Count | Should -Be $before
    }

    It 'reports the certificate problem instead of throwing' {
        $r = Test-SigningConfiguration -Policy (& $script:Policy $true $true $true -Thumb ('0' * 40))
        $r.Ok | Should -BeFalse
        $r.Reason | Should -Match 'was not found'
    }

    It 'fails when a required timestamp server is unreachable' {
        $r = Test-SigningConfiguration -Policy (& $script:Policy $true $true $true -Timestamp 'http://timestamp.invalid.example/tsa' -TimestampRequired $true) -TimeoutSeconds 120
        $r.Ok | Should -BeFalse
        $r.Reason | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-ScriptSigning' {
    BeforeEach {
        $script:CaseDir = Join-Path $script:Root ('case-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $script:CaseDir -Force | Out-Null
        $script:CasePath = Join-Path $script:CaseDir 'detect.ps1'
        Set-Content -LiteralPath $script:CasePath -Value "Write-Output 'Detected'`r`nexit 0" -Encoding ASCII
    }

    It 'returns NotRequested when neither switch is set' {
        $r = Invoke-ScriptSigning -Path $script:CasePath -Category Detection -Policy ([pscustomobject]@{})
        $r.Status | Should -Be 'NotRequested'
        $r.Sha256.Length | Should -Be 64
        (Test-ScriptSignature -Path $script:CasePath).Status | Should -Be 'NotSigned'
    }

    It 'returns NotApplicable for a missing file' {
        (Invoke-ScriptSigning -Path (Join-Path $script:CaseDir 'absent.ps1') -Category Detection -Policy (& $script:Policy $true $false $false)).Status |
            Should -Be 'NotApplicable'
    }

    It 'fails a require-only category when the file is unsigned' {
        $r = Invoke-ScriptSigning -Path $script:CasePath -Category Requirements -Policy ([pscustomobject]@{ RequireRequirements = $true })
        $r.Status | Should -Be 'Failed'
        $r.Reason | Should -Match 'required'
    }

    It 'signs the file and records the signer thumbprint' {
        $r = Invoke-ScriptSigning -Path $script:CasePath -Category Detection -Policy (& $script:Policy $true $false $false)
        $r.Status | Should -Be 'SignedAndVerified'
        $r.Thumbprint | Should -Be $script:Thumbprint
        (Get-Content -LiteralPath $script:CasePath -Raw) | Should -Match 'SIG # Begin signature block'
    }

    It 'acts only on the requested category' {
        $r = Invoke-ScriptSigning -Path $script:CasePath -Category Deployment -Policy (& $script:Policy $true $false $false)
        $r.Status | Should -Be 'NotRequested'
        (Get-Content -LiteralPath $script:CasePath -Raw) | Should -Not -Match 'SIG # Begin signature block'
    }

    It 'preserves an existing signature on an unchanged third-party file' {
        $null = Invoke-ScriptSigning -Path $script:CasePath -Category Deployment -Policy (& $script:Policy $false $false $true)
        $hash = (Get-FileHash -Path $script:CasePath -Algorithm SHA256).Hash
        $again = Invoke-ScriptSigning -Path $script:CasePath -Category Deployment -Policy (& $script:Policy $false $false $true) -PreserveExisting
        $again.Status | Should -Be 'ExistingSignatureValid'
        (Get-FileHash -Path $script:CasePath -Algorithm SHA256).Hash | Should -Be $hash
    }

    It 'accepts an already valid signature when automatic signing is off but required' {
        $null = Invoke-ScriptSigning -Path $script:CasePath -Category Detection -Policy (& $script:Policy $true $false $false)
        (Invoke-ScriptSigning -Path $script:CasePath -Category Detection -Policy (& $script:Policy $false $false $false -RD $true)).Status |
            Should -Be 'ExistingSignatureValid'
    }

    It 'throws when a required timestamp cannot be obtained' {
        { Invoke-ScriptSigning -Path $script:CasePath -Category Detection `
            -Policy (& $script:Policy $true $false $false -Timestamp 'http://timestamp.invalid.example/tsa' -TimestampRequired $true) } |
            Should -Throw '*timestamp*'
    }
}

Describe 'Signature verification' {
    It 'reports NotSigned for a plain script' {
        $p = Join-Path $script:Root 'plain.ps1'
        Set-Content -LiteralPath $p -Value 'exit 0' -Encoding ASCII
        $r = Test-ScriptSignature -Path $p
        $r.Valid | Should -BeFalse
        $r.SignatureIntact | Should -BeFalse
        $r.Status | Should -Be 'NotSigned'
        $r.Timestamped | Should -BeFalse
        $r.TimestampVerified | Should -BeFalse
    }

    It 'reports NotFound for a missing path' {
        (Test-ScriptSignature -Path (Join-Path $script:Root 'nope.ps1')).Status | Should -Be 'NotFound'
    }

    It 'separates an intact signature from host trust' {
        $p = Join-Path $script:Root 'intact.ps1'
        Set-Content -LiteralPath $p -Value 'exit 0' -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $p -Category Detection -Policy (& $script:Policy $true $false $false)
        $r = Test-ScriptSignature -Path $p
        $r.SignatureIntact | Should -BeTrue
        $r.Thumbprint | Should -Be $script:Thumbprint
        $r.Timestamped | Should -BeFalse
        ($r.Valid -eq $r.TrustedOnThisHost) | Should -BeTrue
    }

    It 'reports HashMismatch after a one-byte change made after signing' {
        $p = Join-Path $script:Root 'tamper.ps1'
        Set-Content -LiteralPath $p -Value "Write-Output 'Detected'`r`nexit 0" -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $p -Category Detection -Policy (& $script:Policy $true $false $false)
        $bytes = [System.IO.File]::ReadAllBytes($p)
        $bytes[0] = [byte]0x20
        [System.IO.File]::WriteAllBytes($p, $bytes)
        $r = Test-ScriptSignature -Path $p
        $r.Status | Should -Be 'HashMismatch'
        $r.SignatureIntact | Should -BeFalse
    }

    It 'verifies bytes without leaving the temporary file behind' {
        $p = Join-Path $script:Root 'bytes.ps1'
        Set-Content -LiteralPath $p -Value "Write-Output 'Detected'`r`nexit 0" -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $p -Category Detection -Policy (& $script:Policy $true $false $false)
        $before = @(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter 'appackager-sign-*.ps1' -ErrorAction SilentlyContinue).Count
        $r = Test-ScriptSignatureBytes -Bytes ([System.IO.File]::ReadAllBytes($p))
        $r.SignatureIntact | Should -BeTrue
        $r.Sha256.Length | Should -Be 64
        @(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter 'appackager-sign-*.ps1' -ErrorAction SilentlyContinue).Count | Should -Be $before
    }

    It 'reports a tampered byte array as HashMismatch' {
        $p = Join-Path $script:Root 'bytes2.ps1'
        Set-Content -LiteralPath $p -Value "Write-Output 'Detected'`r`nexit 0" -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $p -Category Detection -Policy (& $script:Policy $true $false $false)
        $bytes = [System.IO.File]::ReadAllBytes($p)
        $bytes[1] = [byte]0x20
        (Test-ScriptSignatureBytes -Bytes $bytes).Status | Should -Be 'HashMismatch'
    }
}

Describe 'Transport and encoding' {
    BeforeAll {
        $script:Encodings = [ordered]@{
            ASCII     = (New-Object System.Text.ASCIIEncoding)
            UTF8BOM   = (New-Object System.Text.UTF8Encoding($true))
            UTF8NoBOM = (New-Object System.Text.UTF8Encoding($false))
            UTF16LE   = (New-Object System.Text.UnicodeEncoding($false, $true))
        }
        $script:EncRoot = Join-Path $script:Root 'encodings'
        New-Item -ItemType Directory -Path $script:EncRoot -Force | Out-Null
        foreach ($name in $script:Encodings.Keys) {
            $body = if ($name -eq 'ASCII') { "Write-Output 'Detected'`r`nexit 0" }
                    else { "Write-Output 'Detected caf" + [char]0xE9 + "'`r`nexit 0" }
            $path = Join-Path $script:EncRoot "$name.ps1"
            [System.IO.File]::WriteAllText($path, $body, $script:Encodings[$name])
            $null = Invoke-ScriptSigning -Path $path -Category Detection -Policy (& $script:Policy $true $false $false)
        }
    }

    It 'detects the encoding of the signed <_> fixture' -ForEach @('ASCII', 'UTF8BOM', 'UTF8NoBOM', 'UTF16LE') {
        (Get-SignedScriptRepresentation -Path (Join-Path $script:EncRoot "$_.ps1")).Encoding | Should -Be $_
    }

    It 'keeps the <_> signature intact through the byte and file transports' -ForEach @('ASCII', 'UTF8BOM', 'UTF8NoBOM', 'UTF16LE') {
        $path = Join-Path $script:EncRoot "$_.ps1"
        (Test-ScriptSignature -Path $path).SignatureIntact | Should -BeTrue
        foreach ($transport in 'File', 'Base64') {
            $r = Test-SignedScriptRoundTrip -Path $path -Transport $transport
            $r.BytesIdentical | Should -BeTrue
            $r.SignatureIntact | Should -BeTrue
        }
        (Test-ScriptSignatureBytes -Bytes ([System.IO.File]::ReadAllBytes($path))).SignatureIntact | Should -BeTrue
    }

    It 'keeps the <_> signature intact through a text transport that preserves the file encoding' -ForEach @('ASCII', 'UTF8BOM', 'UTF8NoBOM', 'UTF16LE') {
        $r = Test-SignedScriptRoundTrip -Path (Join-Path $script:EncRoot "$_.ps1") -Transport Text
        $r.BytesIdentical | Should -BeTrue
        $r.SignatureIntact | Should -BeTrue
    }

    It 'breaks <_> when a text transport forces UTF-8 without a BOM' -ForEach @('UTF8BOM', 'UTF16LE') {
        $r = Test-SignedScriptRoundTrip -Path (Join-Path $script:EncRoot "$_.ps1") `
            -Transport Text -TextEncoding (New-Object System.Text.UTF8Encoding($false))
        $r.BytesIdentical | Should -BeFalse
        $r.Status | Should -Be 'HashMismatch'
    }

    It 'stays under the ConfigMgr detection script limit after signing' {
        $rep = Get-SignedScriptRepresentation -Path (Join-Path $script:EncRoot 'ASCII.ps1')
        $rep.Length | Should -BeLessThan 32768
        $rep.ExceedsCMLimit | Should -BeFalse
        $rep.CMScriptMaxBytes | Should -Be 32768
    }

    It 'flags content above the ConfigMgr limit' {
        $p = Join-Path $script:Root 'big.ps1'
        [System.IO.File]::WriteAllText($p, ('#' + ('x' * 40000)), (New-Object System.Text.ASCIIEncoding))
        (Get-SignedScriptRepresentation -Path $p).ExceedsCMLimit | Should -BeTrue
    }
}

Describe 'New-DeploymentLauncherCommand' {
    It 'reproduces the existing unsigned BAT body' {
        (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $false).BatBody |
            Should -Be ("@echo off`r`nPowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File `"%~dp0install.ps1`"`r`nexit /b %ERRORLEVEL%")
    }

    It 'keeps the exit-code override shape' {
        $r = New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $false -BatExitCode '3010'
        $r.BatBody | Should -Match 'if %ERRORLEVEL% EQU 0 exit /b 3010'
        $r.BatBody | Should -Match 'exit /b %ERRORLEVEL%$'
    }

    It 'emits no execution-policy argument in signed mode' {
        foreach ($scriptHost in 'x64', 'x86') {
            foreach ($exitCode in '%ERRORLEVEL%', '3010') {
                $r = New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true -ScriptHost $scriptHost -BatExitCode $exitCode
                foreach ($text in @($r.CommandLine, $r.BatBody, $r.BatInvoke)) {
                    $text | Should -Not -Match '(?i)executionpolicy'
                    $text | Should -Not -Match '(?i)\bbypass\b'
                    $text | Should -Not -Match '(?i)\-ep\b'
                    $text | Should -Not -Match '(?i)encodedcommand'
                }
                $r.CommandLine | Should -Match '\-NoProfile \-NonInteractive \-File'
            }
        }
    }

    It 'selects the 32-bit host explicitly' {
        $r = New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true -ScriptHost x86
        $r.Executable | Should -Be '%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        $r.CommandLine | Should -BeLike '%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe *'
    }

    It 'passes its own signed output through the launcher chain check' {
        $stage = Join-Path $script:Root ('chain-clean-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $stage 'install.bat') `
            -Value (New-DeploymentLauncherCommand -Script 'install.ps1' -Signed $true).BatBody -Encoding ASCII
        $r = Test-DeploymentLauncherChain -StageRoot $stage -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{ SignDeployment = $true })
        $r.BypassFree | Should -BeTrue
        $r.Findings.Count | Should -Be 0
    }
}

Describe 'Test-DeploymentLauncherChain' {
    BeforeEach {
        $script:ChainRoot = Join-Path $script:Root ('chain-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $script:ChainRoot -Force | Out-Null
    }

    It 'catches every execution-policy spelling and abbreviation' {
        foreach ($spelling in '-ExecutionPolicy Bypass', '-ep Bypass', '-ex Bypass', '-ExecutionPo Bypass', '-executionpolicy bypass', '/ExecutionPolicy Bypass') {
            $file = Join-Path $script:ChainRoot 'ep.bat'
            Set-Content -LiteralPath $file -Value ('PowerShell.exe -NonInteractive {0} -File "%~dp0install.ps1"' -f $spelling) -Encoding ASCII
            $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
            @($r.Findings | Where-Object { $_.Code -eq 'LauncherExecutionPolicy' }).Count | Should -BeGreaterThan 0
            Remove-Item $file -Force
        }
    }

    It 'catches encoded command invocations' {
        foreach ($spelling in '-EncodedCommand', '-enc', '-e') {
            $file = Join-Path $script:ChainRoot 'enc.bat'
            Set-Content -LiteralPath $file -Value ('PowerShell.exe {0} ZQB4AGkAdAAgADAA' -f $spelling) -Encoding ASCII
            $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
            @($r.Findings | Where-Object { $_.Code -eq 'LauncherEncodedCommand' }).Count | Should -BeGreaterThan 0
            Remove-Item $file -Force
        }
    }

    It 'flags a custom command carrying bypass' {
        $manifest = [pscustomobject]@{ InstallCommandLine = 'cmd.exe /c setup.exe && powershell -nop -executionpolicy bypass -file extra.ps1' }
        $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest $manifest -Policy ([pscustomobject]@{})
        @($r.Findings | Where-Object { $_.Code -eq 'CustomCommandBypass' }).Count | Should -BeGreaterThan 0
        @($r.Findings | Where-Object { $_.Code -eq 'LauncherExecutionPolicy' }).Count | Should -BeGreaterThan 0
    }

    It 'inspects deployment type entries too' {
        $manifest = [pscustomobject]@{
            DeploymentTypes = @([pscustomobject]@{ InstallCommandLine = 'powershell.exe -ExecutionPolicy Bypass -File "Invoke-AppDeployToolkit.ps1"' })
        }
        $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest $manifest -Policy ([pscustomobject]@{})
        $r.CommandsInspected | Should -Be 1
        @($r.Findings | Where-Object { $_.Code -eq 'LauncherExecutionPolicy' }).Count | Should -BeGreaterThan 0
    }

    It 'accepts the signed Extend orchestrator child-launch shape' {
        Set-Content -LiteralPath (Join-Path $script:ChainRoot 'install.ps1') -Encoding ASCII -Value @(
            '$steps = @("before.ps1", "generated.ps1", "after.ps1")',
            'foreach ($step in $steps) {',
            '    $p = Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-NonInteractive", "-File", "$PSScriptRoot\$step") -Wait -PassThru',
            '    if ($p.ExitCode -ne 0) { exit $p.ExitCode }',
            '}',
            'exit 0'
        )
        (Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{ SignDeployment = $true })).BypassFree |
            Should -BeTrue
    }

    It 'rejects the Extend orchestrator when it still passes a policy argument in signed mode' {
        Set-Content -LiteralPath (Join-Path $script:ChainRoot 'install.ps1') -Encoding ASCII -Value @(
            '$p = Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "step.ps1") -Wait -PassThru',
            'exit $p.ExitCode'
        )
        { Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{ SignDeployment = $true }) } |
            Should -Throw '*Signed deployment mode refuses this build*'
    }

    It 'token-scans a launcher that never names powershell literally' {
        $cases = @(
            'pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"',
            'start "" "%PSEXE%" -ep Bypass -File "%~dp0install.ps1"',
            '& $exe -enc ZQB4AGkAdAAgADAA'
        )
        foreach ($case in $cases) {
            $file = Join-Path $script:ChainRoot 'indirect.cmd'
            Set-Content -LiteralPath $file -Value $case -Encoding ASCII
            $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
            $r.BypassFree | Should -BeFalse
            Remove-Item $file -Force
        }
    }

    It 'does not report findings inside an Authenticode signature block' {
        $file = Join-Path $script:ChainRoot 'signed.ps1'
        Set-Content -LiteralPath $file -Value "exit 0" -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $file -Category Deployment -Policy (& $script:Policy $false $false $true)
        $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
        $r.Findings.Count | Should -Be 0
    }

    It 'enforces findings when RequireDeployment is set without automatic signing' {
        $file = Join-Path $script:ChainRoot 'install.bat'
        Set-Content -LiteralPath $file -Value 'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"' -Encoding ASCII
        { Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) `
            -Policy ([pscustomobject]@{ SignDeployment = $false; RequireDeployment = $true }) } |
            Should -Throw '*refuses this build*'
    }

    It 'is informational in unsigned mode and throws in signed mode' {
        $file = Join-Path $script:ChainRoot 'install.bat'
        Set-Content -LiteralPath $file -Value 'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"' -Encoding ASCII
        $unsigned = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
        $unsigned.BypassFree | Should -BeFalse
        $unsigned.Enforced | Should -BeFalse
        { Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{ SignDeployment = $true }) } |
            Should -Throw '*Signed deployment mode refuses this build*'
    }

    It 'reports the offending file and line' {
        $file = Join-Path $script:ChainRoot 'install.bat'
        Set-Content -LiteralPath $file -Value "@echo off`r`nPowerShell.exe -ExecutionPolicy Bypass -File `"%~dp0install.ps1`"" -Encoding ASCII
        $r = Test-DeploymentLauncherChain -StageRoot $script:ChainRoot -Manifest ([pscustomobject]@{}) -Policy ([pscustomobject]@{})
        $finding = @($r.Findings | Where-Object { $_.Code -eq 'LauncherExecutionPolicy' })[0]
        $finding.File | Should -Be $file
        $finding.LineNumber | Should -Be 2
    }
}

Describe 'Invoke-CategorySigning' {
    It 'signs only the selected categories for combination <Key>' -ForEach @(
        @{ Key = 'D0R0P0'; D = $false; R = $false; P = $false }
        @{ Key = 'D0R0P1'; D = $false; R = $false; P = $true }
        @{ Key = 'D0R1P0'; D = $false; R = $true;  P = $false }
        @{ Key = 'D0R1P1'; D = $false; R = $true;  P = $true }
        @{ Key = 'D1R0P0'; D = $true;  R = $false; P = $false }
        @{ Key = 'D1R0P1'; D = $true;  R = $false; P = $true }
        @{ Key = 'D1R1P0'; D = $true;  R = $true;  P = $false }
        @{ Key = 'D1R1P1'; D = $true;  R = $true;  P = $true }
    ) {
        $fixture = & $script:NewFixture "combo-$Key" $P
        $block = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $D $R $P)
        $block.Detection.Status | Should -Be $(if ($D) { 'SignedAndVerified' } else { 'NotRequested' })
        $block.Requirements.Status | Should -Be $(if ($R) { 'SignedAndVerified' } else { 'NotRequested' })
        $block.Deployment.Status | Should -Be $(if ($P) { 'SignedAndVerified' } else { 'NotRequested' })
        $block.Deployment.LaunchersBypassFree | Should -Be $P
        @($block.Deployment.Files | Where-Object { $_.Owner -eq 'AppPackager' }).Count | Should -Be 3
        @($block.Requirements.Items).Count | Should -Be 1
        $block.Detection.File | Should -Be 'scripts\detect.ps1'
        $block.PolicyDigest.Length | Should -Be 64
    }

    It 'gives each switch combination its own policy digest' {
        $digests = foreach ($d in $false, $true) {
            foreach ($r in $false, $true) {
                foreach ($p in $false, $true) { (Get-SigningPolicy -Policy (& $script:Policy $d $r $p)).PolicyDigest }
            }
        }
        @($digests | Sort-Object -Unique).Count | Should -Be 8
    }

    It 'replaces the manifest detection text with the signed text' {
        $fixture = & $script:NewFixture 'signed-text' $true
        $null = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $true $false $false)
        $fixture.Manifest.Detection.ScriptFile | Should -Be 'scripts\detect.ps1'
        $fixture.Manifest.Detection.ScriptText | Should -Match 'SIG # Begin signature block'
        (Test-ScriptSignatureBytes -Bytes ([System.Text.Encoding]::UTF8.GetBytes($fixture.Manifest.Detection.ScriptText))).SignatureIntact |
            Should -BeTrue
    }

    It 'rewrites a stale detection script from the current manifest text' {
        $fixture = & $script:NewFixture 'stale-detect' $true
        $stale = "Write-Output 'STALE'`r`nexit 0"
        Set-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\detect.ps1') -Value $stale -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\requirements\r1.ps1') -Value $stale -Encoding ASCII
        $null = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $true $true $false)
        $detect = Get-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\detect.ps1') -Raw
        $detect | Should -Not -Match 'STALE'
        $detect | Should -Match 'Program Files\\Fixture'
        $fixture.Manifest.Detection.ScriptText | Should -Not -Match 'STALE'
        (Get-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\requirements\r1.ps1') -Raw) | Should -Not -Match 'STALE'
    }

    It 'rewrites a stale detection script even when no signing switch is set' {
        $fixture = & $script:NewFixture 'stale-detect-unsigned' $false
        Set-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\detect.ps1') -Value "Write-Output 'STALE'" -Encoding ASCII
        $null = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy ([pscustomobject]@{})
        (Get-Content -LiteralPath (Join-Path $fixture.StageRoot 'scripts\detect.ps1') -Raw) | Should -Not -Match 'STALE'
    }

    It 'materializes pre-signed manifest text whose origin encoding was not UTF-8' -ForEach @('UTF8BOM', 'UTF16LE') {
        $encoding = if ($_ -eq 'UTF8BOM') { New-Object System.Text.UTF8Encoding($true) }
                    else { New-Object System.Text.UnicodeEncoding($false, $true) }
        $fixture = & $script:NewFixture ('presigned-' + $_) $false
        $source = Join-Path $script:Root ('src-' + $_ + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.ps1')
        [System.IO.File]::WriteAllText($source, ("Write-Output 'Detected caf" + [char]0xE9 + "'`r`nexit 0"), $encoding)
        $null = Invoke-ScriptSigning -Path $source -Category Detection -Policy (& $script:Policy $true $false $false)
        $fixture.Manifest.Detection.ScriptText = (Get-SignedScriptRepresentation -Path $source).Text

        $block = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $false -RD $true)
        $block.Detection.Status | Should -Be 'ExistingSignatureValid'
        (Test-ScriptSignature -Path (Join-Path $fixture.StageRoot 'scripts\detect.ps1')).SignatureIntact | Should -BeTrue
    }

    It 'refuses a build whose launcher carries bypass when only RequireDeployment is set' {
        $fixture = & $script:NewFixture 'require-only-launcher' $false
        { Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest `
            -Policy (& $script:Policy $false $false $false -RP $true) } |
            Should -Throw '*refuses this build*'
    }

    It 'refuses to publish when RequireDetection is set and the script is unsigned' {
        $fixture = & $script:NewFixture 'strict-unsigned' $false
        { Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $false -RD $true) } |
            Should -Throw '*RequireDetection*'
    }

    It 'accepts pre-signed detection text from the manifest when automatic signing is off' {
        # The manifest carries the signed representation; a signed file dropped
        # into the stage folder is not an input, because the folder is rewritten.
        $fixture = & $script:NewFixture 'strict-presigned' $false
        $source = Join-Path $script:Root ('presigned-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.ps1')
        [System.IO.File]::WriteAllText($source, $fixture.Manifest.Detection.ScriptText, (New-Object System.Text.UTF8Encoding($false)))
        $null = Invoke-ScriptSigning -Path $source -Category Detection -Policy (& $script:Policy $true $false $false)
        $fixture.Manifest.Detection.ScriptText = (Get-SignedScriptRepresentation -Path $source).Text

        $block = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $false -RD $true)
        $block.Detection.Status | Should -Be 'ExistingSignatureValid'
    }

    It 'preserves an unchanged third-party toolkit signature byte for byte' {
        $fixture = & $script:NewFixture 'thirdparty' $true
        $vendorDir = Join-Path $fixture.StageRoot 'Toolkit'
        New-Item -ItemType Directory -Path $vendorDir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $vendorDir 'Invoke-AppDeployToolkit.ps1') -Value 'exit 0' -Encoding ASCII
        $helper = Join-Path $vendorDir 'AppDeployToolkitMain.ps1'
        Set-Content -LiteralPath $helper -Value "function Get-Vendor { 'v' }" -Encoding ASCII
        $null = Invoke-ScriptSigning -Path $helper -Category Deployment -Policy (& $script:Policy $false $false $true)
        $hashBefore = (Get-FileHash -Path $helper -Algorithm SHA256).Hash

        $block = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $true)
        $entry = @($block.Deployment.Files | Where-Object { $_.RelativePath -like '*AppDeployToolkitMain.ps1' })[0]
        $entry.Status | Should -Be 'ExistingSignatureValid'
        $entry.Owner | Should -Be 'ThirdParty'
        (Get-FileHash -Path $helper -Algorithm SHA256).Hash | Should -Be $hashBefore
    }

    It 'refuses a signed build whose launcher still carries bypass' {
        $fixture = & $script:NewFixture 'bypass-launcher' $false
        { Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $true) } |
            Should -Throw '*Signed deployment mode refuses this build*'
    }

    It 'refuses a signed build when a required timestamp server is unreachable' {
        $fixture = & $script:NewFixture 'timestamp' $true
        { Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest `
            -Policy (& $script:Policy $true $false $false -Timestamp 'http://timestamp.invalid.example/tsa' -TimestampRequired $true) } |
            Should -Throw '*timestamp*'
    }

    It 'refuses a detection script that exceeds the ConfigMgr limit after signing' {
        $fixture = & $script:NewFixture 'oversize' $true
        $pad = '#' + ('x' * 200)
        $fixture.Manifest['Detection'] = @{ Type = 'Script'; ScriptText = (((1..200 | ForEach-Object { $pad }) -join "`r`n") + "`r`nWrite-Output 'Detected'`r`nexit 0") }
        Remove-Item (Join-Path $fixture.StageRoot 'scripts\detect.ps1') -Force
        { Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $true $false $false) } |
            Should -Throw '*32768*'
    }

    It 'does not sign detection or requirement scripts under the deployment switch' {
        $fixture = & $script:NewFixture 'category-isolation' $true
        $null = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest -Policy (& $script:Policy $false $false $true)
        (Test-ScriptSignature -Path (Join-Path $fixture.StageRoot 'scripts\detect.ps1')).Status | Should -Be 'NotSigned'
        (Test-ScriptSignature -Path (Join-Path $fixture.StageRoot 'scripts\requirements\r1.ps1')).Status | Should -Be 'NotSigned'
        (Test-ScriptSignature -Path (Join-Path $fixture.StageRoot 'install.ps1')).SignatureIntact | Should -BeTrue
    }

    It 'requires no certificate when every switch is off' {
        $fixture = & $script:NewFixture 'no-cert' $false
        $block = Invoke-CategorySigning -StageRoot $fixture.StageRoot -ManifestData $fixture.Manifest `
            -Policy ([pscustomobject]@{ CertificateThumbprint = '' })
        $block.Detection.Status | Should -Be 'NotRequested'
        $block.Deployment.LaunchersBypassFree | Should -BeFalse
    }
}
