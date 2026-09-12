#Requires -Modules Pester

<#
.SYNOPSIS
    End-to-end script-signing coverage for a staged build.

.DESCRIPTION
    Every combination of the three signing switches runs through the real
    Write-StageManifest finalization on a fixture stage that carries a
    detection script, a requirement script, install/uninstall wrappers and a
    helper script.

    Trust boundary: the throwaway certificate lives in Cert:\CurrentUser\My
    and nowhere else. No test writes to a root or trusted-publisher store in
    any scope, so assertions use signature intactness (hash and signature
    block agree), never host trust, and nothing can raise a Windows trust
    dialog. The certificate is removed in AfterAll.
#>

BeforeAll {
    # Common imports its companions only when they are absent, so a suite
    # that ran earlier in the same process can leave a stale instance behind;
    # importing both explicitly keeps the signing functions resolvable here.
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerSigning.psd1" -Force -Global -DisableNameChecking
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerWorkbench.psd1" -Force -Global -DisableNameChecking
    Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force

    $script:PreviousSigning = $env:APP_PACKAGER_SIGNING
    $script:PreviousSnapshot = $env:APP_PACKAGER_RUN_SNAPSHOT
    $script:PreviousRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
    $env:APP_PACKAGER_RUN_SNAPSHOT = $null
    $env:APP_PACKAGER_WORKBENCH_ROOT = Join-Path $TestDrive 'signing data'

    $script:CertSubject = 'CN=AppPackager G Signing Fixture'
    $script:Cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $script:CertSubject `
        -CertStoreLocation 'Cert:\CurrentUser\My' -KeyUsage DigitalSignature `
        -NotAfter (Get-Date).AddDays(2) -ErrorAction Stop
    $script:Thumbprint = $script:Cert.Thumbprint

    function New-SigningPolicyJson {
        param(
            [bool]$SignDetection, [bool]$SignRequirements, [bool]$SignDeployment,
            [bool]$RequireDetection = $false, [bool]$RequireRequirements = $false, [bool]$RequireDeployment = $false,
            [string]$Thumbprint = $script:Thumbprint
        )
        return (@{
            SignDetection = $SignDetection; SignRequirements = $SignRequirements; SignDeployment = $SignDeployment
            RequireDetection = $RequireDetection; RequireRequirements = $RequireRequirements; RequireDeployment = $RequireDeployment
            CertificateThumbprint = $Thumbprint; StoreLocation = 'CurrentUser'
            TimestampServer = ''; TimestampRequired = $false; HashAlgorithm = 'SHA256'
        } | ConvertTo-Json -Depth 4 -Compress)
    }

    function New-SigningFixtureStage {
        <#
            A stage root under a folder whose name contains a space, holding
            every signing category: detector, requirement rule, generated
            wrappers, and an AppPackager-owned helper.
        #>
        param([string]$Name, [switch]$NonAscii)

        $root = Join-Path (Join-Path $TestDrive 'stage fixtures') $Name
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $root 'setup.exe'), 'not a real installer')

        $helper = "function Write-GFixtureState {`r`n    param([string]`$Message)`r`n    Write-Output `$Message`r`n}`r`n"
        if ($NonAscii) {
            # Built from code points so this test file itself stays ASCII.
            $accented = [string][char]0x00E9 + [string][char]0x00FC + [string][char]0x2014
            $helper += ("# caption: {0}`r`n" -f $accented)
        }
        # UTF-8 without a BOM: the encoding the signing module reports as
        # UTF8NoBOM and the one a signed script must keep byte for byte.
        [System.IO.File]::WriteAllText((Join-Path $root 'helper.ps1'), $helper, (New-Object System.Text.UTF8Encoding($false)))

        Write-ContentWrappers -OutputPath $root `
            -InstallPs1Content 'exit 0' `
            -UninstallPs1Content 'exit 0' 6>$null | Out-Null

        return $root
    }

    function New-SigningFixtureManifest {
        return @{
            AppName         = 'G Signing Fixture'
            Publisher       = 'Contoso'
            SoftwareVersion = '1.0.0'
            InstallerFile   = 'setup.exe'
            InstallerType   = 'EXE'
            Architecture    = 'x64'
            Detection       = @{
                Type       = 'Script'
                ScriptText = "if (Test-Path 'HKLM:\SOFTWARE\GFixture') { Write-Output 'Installed' }`r`n"
            }
            Requirements    = @(
                @{ RuleId = 'gfixture-ready'; ScriptText = "Write-Output `$true`r`n" }
            )
        }
    }

    function Invoke-SigningFixtureStage {
        <#
            Generates the wrappers and writes the manifest under one policy,
            the way a packager child does, and returns the finalized manifest.
        #>
        param([string]$Name, [string]$PolicyJson, [switch]$NonAscii)

        $env:APP_PACKAGER_SIGNING = $PolicyJson
        try {
            $root = New-SigningFixtureStage -Name $Name -NonAscii:$NonAscii
            $manifest = New-SigningFixtureManifest
            Write-StageManifest -Path (Join-Path $root 'stage-manifest.json') -ManifestData $manifest `
                -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null | Out-Null
            return [pscustomobject]@{ StageRoot = $root; Manifest = $manifest }
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }
}

AfterAll {
    $env:APP_PACKAGER_SIGNING = $script:PreviousSigning
    $env:APP_PACKAGER_RUN_SNAPSHOT = $script:PreviousSnapshot
    $env:APP_PACKAGER_WORKBENCH_ROOT = $script:PreviousRoot
    if ($script:Thumbprint) {
        Get-ChildItem -Path 'Cert:\CurrentUser\My' |
            Where-Object { $_.Thumbprint -eq $script:Thumbprint } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Signing switch combinations' {
    It 'signs only the selected categories (<Label>)' -ForEach @(
        @{ Label = 'none';                 Det = $false; Req = $false; Dep = $false }
        @{ Label = 'detection';            Det = $true;  Req = $false; Dep = $false }
        @{ Label = 'requirements';         Det = $false; Req = $true;  Dep = $false }
        @{ Label = 'deployment';           Det = $false; Req = $false; Dep = $true }
        @{ Label = 'detection+req';        Det = $true;  Req = $true;  Dep = $false }
        @{ Label = 'detection+dep';        Det = $true;  Req = $false; Dep = $true }
        @{ Label = 'req+dep';              Det = $false; Req = $true;  Dep = $true }
        @{ Label = 'all three';            Det = $true;  Req = $true;  Dep = $true }
    ) {
        $policy = New-SigningPolicyJson -SignDetection $Det -SignRequirements $Req -SignDeployment $Dep
        $result = Invoke-SigningFixtureStage -Name ('combo-' + ($Label -replace '[^a-z0-9]', '-')) -PolicyJson $policy
        $signing = $result.Manifest.ScriptSigning

        $expectedDetection = if ($Det) { 'SignedAndVerified' } else { 'NotRequested' }
        $expectedRequirements = if ($Req) { 'SignedAndVerified' } else { 'NotRequested' }
        $expectedDeployment = if ($Dep) { 'SignedAndVerified' } else { 'NotRequested' }

        $signing.Detection.Status | Should -Be $expectedDetection
        $signing.Requirements.Status | Should -Be $expectedRequirements
        $signing.Deployment.Status | Should -Be $expectedDeployment

        if ($Det) {
            $signing.Detection.Thumbprint | Should -Be $script:Thumbprint
            (Test-ScriptSignature -Path (Join-Path $result.StageRoot 'scripts\detect.ps1')).SignatureIntact | Should -BeTrue
        }
        if ($Req) {
            @($signing.Requirements.Items).Count | Should -Be 1
            (Test-ScriptSignature -Path (Join-Path $result.StageRoot 'scripts\requirements\gfixture-ready.ps1')).SignatureIntact | Should -BeTrue
        }
        if ($Dep) {
            foreach ($name in 'install.ps1', 'uninstall.ps1', 'helper.ps1') {
                (Test-ScriptSignature -Path (Join-Path $result.StageRoot $name)).SignatureIntact | Should -BeTrue
            }
        }
        else {
            (Test-ScriptSignature -Path (Join-Path $result.StageRoot 'install.ps1')).Status | Should -Be 'NotSigned'
        }
    }

    It 'keeps every signed build free of an execution-policy launcher argument' {
        $result = Invoke-SigningFixtureStage -Name 'signed-launchers' -PolicyJson (New-SigningPolicyJson -SignDetection $true -SignRequirements $true -SignDeployment $true)
        foreach ($name in 'install.bat', 'uninstall.bat') {
            $body = [System.IO.File]::ReadAllText((Join-Path $result.StageRoot $name))
            $body | Should -Not -Match '(?i)-ExecutionPolicy'
            $body | Should -Not -Match '(?i)\bbypass\b'
            $body | Should -Not -Match '(?i)-EncodedCommand'
            $body | Should -Match '(?i)-NoProfile -NonInteractive -File'
        }
        $result.Manifest.ScriptSigning.Deployment.LaunchersBypassFree | Should -BeTrue
    }

    It 'leaves the historical launcher string in place when deployment signing is off' {
        $result = Invoke-SigningFixtureStage -Name 'unsigned-launchers' -PolicyJson (New-SigningPolicyJson -SignDetection $false -SignRequirements $false -SignDeployment $false)
        [System.IO.File]::ReadAllText((Join-Path $result.StageRoot 'install.bat')) | Should -Match '(?i)-NonInteractive -ExecutionPolicy Bypass -File'
    }

    It 'covers the signed scripts with the stage file hashes' {
        $result = Invoke-SigningFixtureStage -Name 'hash-coverage' -PolicyJson (New-SigningPolicyJson -SignDetection $true -SignRequirements $true -SignDeployment $true)
        $paths = @($result.Manifest.FileHashes | ForEach-Object { $_.RelativePath })
        $paths | Should -Contain 'install.ps1'
        $paths | Should -Contain 'helper.ps1'
        $paths | Should -Contain 'scripts\detect.ps1'
        $paths | Should -Contain 'scripts\requirements\gfixture-ready.ps1'
        $comparison = Compare-StageFileHashes -Root $result.StageRoot -Expected $result.Manifest.FileHashes -Exclude @('stage-manifest.json')
        $comparison.Pass | Should -BeTrue
    }

    It 'carries the exact signed detection text into the manifest' {
        $result = Invoke-SigningFixtureStage -Name 'detection-text' -PolicyJson (New-SigningPolicyJson -SignDetection $true -SignRequirements $false -SignDeployment $false)
        $result.Manifest.Detection.ScriptFile | Should -Be 'scripts\detect.ps1'
        $result.Manifest.Detection.ScriptText | Should -Match 'SIG # Begin signature block'
        $representation = Get-SignedScriptRepresentation -Path (Join-Path $result.StageRoot 'scripts\detect.ps1')
        $result.Manifest.Detection.ScriptText | Should -Be $representation.Text
    }
}

Describe 'Strict signature requirements' {
    It 'throws before writing any script when a required category has no certificate' {
        $root = Join-Path (Join-Path $TestDrive 'stage fixtures') 'require-no-cert'
        $env:APP_PACKAGER_SIGNING = New-SigningPolicyJson -SignDetection $true -SignRequirements $false -SignDeployment $false -RequireDetection $true -Thumbprint ''
        try {
            $stage = New-SigningFixtureStage -Name 'require-no-cert'
            $manifest = New-SigningFixtureManifest
            { Write-StageManifest -Path (Join-Path $stage 'stage-manifest.json') -ManifestData $manifest -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null } |
                Should -Throw '*SigningCertificateUnavailable*'
            Test-Path -LiteralPath (Join-Path $stage 'scripts') | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $stage 'stage-manifest.json') | Should -BeFalse
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }

    It 'throws when a thumbprint names a certificate that is not in the store' {
        $env:APP_PACKAGER_SIGNING = New-SigningPolicyJson -SignDetection $true -SignRequirements $false -SignDeployment $false -Thumbprint ('A' * 40)
        try {
            $stage = New-SigningFixtureStage -Name 'missing-cert'
            { Write-StageManifest -Path (Join-Path $stage 'stage-manifest.json') -ManifestData (New-SigningFixtureManifest) -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null } |
                Should -Throw '*was not found*'
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }

    It 'refuses to publish a required category that carries no signature' {
        # RequireDetection without SignDetection: the detector is materialized
        # unsigned, verifies as NotSigned, and the category check refuses.
        $env:APP_PACKAGER_SIGNING = New-SigningPolicyJson -SignDetection $false -SignRequirements $false -SignDeployment $false -RequireDetection $true
        try {
            $stage = New-SigningFixtureStage -Name 'require-without-sign'
            { Write-StageManifest -Path (Join-Path $stage 'stage-manifest.json') -ManifestData (New-SigningFixtureManifest) -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null } |
                Should -Throw '*refusing to publish*'
            Test-Path -LiteralPath (Join-Path $stage 'stage-manifest.json') | Should -BeFalse
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }

    It 'refuses a required deployment whose wrappers were generated unsigned' {
        # RequireDeployment alone still enforces the launcher chain, so the
        # historical unsigned wrappers are refused before the category check.
        $env:APP_PACKAGER_SIGNING = New-SigningPolicyJson -SignDetection $false -SignRequirements $false -SignDeployment $false -RequireDeployment $true
        try {
            $stage = New-SigningFixtureStage -Name 'require-deployment-unsigned'
            { Write-StageManifest -Path (Join-Path $stage 'stage-manifest.json') -ManifestData (New-SigningFixtureManifest) -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null } |
                Should -Throw '*refuses this build*'
            Test-Path -LiteralPath (Join-Path $stage 'stage-manifest.json') | Should -BeFalse
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }

    It 'refuses a signed deployment whose launcher still relaxes the execution policy' {
        $env:APP_PACKAGER_SIGNING = New-SigningPolicyJson -SignDetection $false -SignRequirements $false -SignDeployment $true
        try {
            $stage = New-SigningFixtureStage -Name 'relaxed-launcher'
            # A hand-authored launcher that the generated wrappers would never
            # produce: the chain check has to catch it, not trust the generator.
            [System.IO.File]::WriteAllText((Join-Path $stage 'legacy-launch.bat'),
                "@echo off`r`npowershell.exe -ExecutionPolicy Bypass -File `"%~dp0install.ps1`"`r`nexit /b %ERRORLEVEL%`r`n")
            { Write-StageManifest -Path (Join-Path $stage 'stage-manifest.json') -ManifestData (New-SigningFixtureManifest) -PackagerScriptPath (Join-Path $TestDrive 'package-gsign.ps1') 6>$null } |
                Should -Throw '*Signed deployment mode refuses this build*'
        }
        finally { $env:APP_PACKAGER_SIGNING = $script:PreviousSigning }
    }
}

Describe 'Post-sign mutation and encoding' {
    BeforeAll {
        $script:Signed = Invoke-SigningFixtureStage -Name 'mutation' -PolicyJson (New-SigningPolicyJson -SignDetection $true -SignRequirements $true -SignDeployment $true) -NonAscii
    }

    It 'keeps a non-ASCII helper intact through signing' {
        $helper = Join-Path $script:Signed.StageRoot 'helper.ps1'
        $bytes = [System.IO.File]::ReadAllBytes($helper)
        @($bytes | Where-Object { $_ -gt 0x7F }).Count | Should -BeGreaterThan 0
        (Test-ScriptSignature -Path $helper).SignatureIntact | Should -BeTrue
        (Get-SignedScriptRepresentation -Path $helper).Encoding | Should -Be 'UTF8NoBOM'
    }

    It 'reports HashMismatch after one byte changes behind the signature' {
        $target = Join-Path $script:Signed.StageRoot 'install.ps1'
        (Test-ScriptSignature -Path $target).SignatureIntact | Should -BeTrue

        $bytes = [System.IO.File]::ReadAllBytes($target)
        $bytes[0] = [byte]([int]$bytes[0] -bxor 0x20)
        [System.IO.File]::WriteAllBytes($target, $bytes)

        $after = Test-ScriptSignature -Path $target
        $after.Status | Should -Be 'HashMismatch'
        $after.SignatureIntact | Should -BeFalse
    }

    It 'reports HashMismatch for the same mutation through the byte transport' {
        $target = Join-Path $script:Signed.StageRoot 'uninstall.ps1'
        $bytes = [System.IO.File]::ReadAllBytes($target)
        # Intactness, not host trust: this workstation never trusts the
        # throwaway signer, and it is not required to.
        (Test-ScriptSignatureBytes -Bytes $bytes).SignatureIntact | Should -BeTrue
        $bytes[0] = [byte]([int]$bytes[0] -bxor 0x20)
        (Test-ScriptSignatureBytes -Bytes $bytes).Status | Should -Be 'HashMismatch'
    }

    It 'signs a stage whose path contains a space' {
        $script:Signed.StageRoot | Should -Match '\s'
        (Test-ScriptSignature -Path (Join-Path $script:Signed.StageRoot 'scripts\detect.ps1')).SignatureIntact | Should -BeTrue
    }
}
