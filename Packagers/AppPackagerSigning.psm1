<#
.SYNOPSIS
    Authenticode signing service for staged AppPackager content.

.DESCRIPTION
    One service used by the catalog packagers, BYO, the CLI and both target
    adapters. Three independent categories (Detection, Requirements,
    Deployment) each have a "sign" switch and a "require" switch. Nothing
    here falls back to unsigned output when a require switch is set.

    Deployment launcher strings produced by New-DeploymentLauncherCommand:

      unsigned, x64 BAT body line
        PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"
      unsigned, x64 command line
        PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "install.ps1"
      signed, x64 BAT body line
        PowerShell.exe -NoProfile -NonInteractive -File "%~dp0install.ps1"
      signed, x64 command line
        PowerShell.exe -NoProfile -NonInteractive -File "install.ps1"
      x86 replaces PowerShell.exe with
        %SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe

    Signed mode carries no execution-policy argument in any spelling and no
    encoded command: the endpoint's configured policy governs.

    Verification distinguishes an intact signature from a locally trusted one.
    SignatureIntact means the hash and the signature block still match the
    content; TrustedOnThisHost means the signing host also trusts the chain.
    A build host is not required to trust the signer, and trusting it here
    proves nothing about the endpoints.
#>

Set-StrictMode -Version 2.0

# ConfigMgr refuses a detection script larger than this; the Authenticode
# block counts toward it.
$script:ConfigMgrDetectionScriptMaxBytes = 32768

$script:CodeSigningEku = '1.3.6.1.5.5.7.3.3'

# Key storage providers that hold the private key on removable hardware or
# behind an interactive PIN; signing with one cannot complete unattended.
$script:InteractiveKeyProviderPatterns = @(
    'Smart Card',
    'SmartCard',
    'Token',
    'eToken',
    'SafeNet',
    'nCipher',
    'Luna',
    'YubiKey',
    'Gemalto',
    'Thales'
)

$script:DefaultSigningPolicy = [ordered]@{
    SignDetection         = $false
    SignRequirements      = $false
    SignDeployment        = $false
    RequireDetection      = $false
    RequireRequirements   = $false
    RequireDeployment     = $false
    CertificateThumbprint = ''
    StoreLocation         = 'CurrentUser'
    TimestampServer       = ''
    TimestampRequired     = $false
    HashAlgorithm         = 'SHA256'
}

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

function Get-SigningPropertyValue {
    param($InputObject, [string]$Name, $Default)

    if ($null -eq $InputObject) { return $Default }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            $v = $InputObject[$Name]
            if ($null -eq $v) { return $Default }
            return $v
        }
        return $Default
    }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $Default }
    if ($null -eq $prop.Value) { return $Default }
    return $prop.Value
}

function ConvertTo-SigningBoolean {
    param($Value, [bool]$Default = $false)

    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
    switch -Regex ($text.Trim()) {
        '^(?i)(true|1|yes|on)$'  { return $true }
        '^(?i)(false|0|no|off)$' { return $false }
        default                  { return $Default }
    }
}

function Get-SigningSha256 {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($Bytes)) -replace '-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Get-SigningFileSha256 {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-SigningSha256 -Bytes ([System.IO.File]::ReadAllBytes($Path)))
}

function New-SigningTempScriptPath {
    $name = 'appackager-sign-{0}.ps1' -f ([guid]::NewGuid().ToString('N'))
    return (Join-Path ([System.IO.Path]::GetTempPath()) $name)
}

function Test-SigningCertificateIsCodeSigning {
    param([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $sawEku = $false
    foreach ($ext in $Certificate.Extensions) {
        if ($ext -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]) {
            $sawEku = $true
            foreach ($oid in $ext.EnhancedKeyUsages) {
                if ($oid.Value -eq $script:CodeSigningEku) { return $true }
            }
        }
    }
    # A certificate without an EKU extension is usable for any purpose.
    return (-not $sawEku)
}

function Get-SigningPrivateKeyInfo {
    param([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $info = [pscustomobject]@{
        HasPrivateKey = [bool]$Certificate.HasPrivateKey
        Accessible    = $false
        MayPrompt     = $false
        Provider      = ''
        Reason        = ''
    }
    if (-not $Certificate.HasPrivateKey) {
        $info.Reason = 'Certificate has no associated private key.'
        return $info
    }

    $key = $null
    $ownsKey = $true
    try {
        $key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
        if ($null -eq $key) {
            $key = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($Certificate)
        }
    }
    catch {
        # A key the current process created is still bound to an ephemeral
        # handle that the extension methods refuse to reopen. The key material
        # is present in the store entry and CryptAcquireCertificatePrivateKey,
        # which Set-AuthenticodeSignature uses, still reaches it.
        $handleError = $_.Exception.Message
        $ownsKey = $false
        try { $key = $Certificate.PrivateKey }
        catch { $key = $null }
        if ($null -eq $key) {
            if ($handleError -match '(?i)ephemeral') {
                $info.Accessible = $true
                return $info
            }
            $info.Reason = 'Private key handle could not be opened: {0}' -f $handleError
            return $info
        }
    }

    if ($null -eq $key) {
        $info.Reason = 'Private key handle could not be opened.'
        return $info
    }

    try {
        $info.Accessible = $true
        # CNG keys expose Key.Provider; CAPI keys expose CspKeyContainerInfo.
        if ($key.PSObject.Properties['Key'] -and $null -ne $key.Key) {
            $info.Provider = [string]$key.Key.Provider.Provider
        }
        elseif ($key.PSObject.Properties['CspKeyContainerInfo'] -and $null -ne $key.CspKeyContainerInfo) {
            $info.Provider = [string]$key.CspKeyContainerInfo.ProviderName
            if ($key.CspKeyContainerInfo.HardwareDevice) { $info.MayPrompt = $true }
            if ($key.CspKeyContainerInfo.Protected) { $info.MayPrompt = $true }
        }
        foreach ($pattern in $script:InteractiveKeyProviderPatterns) {
            if ($info.Provider -and $info.Provider -like ('*{0}*' -f $pattern)) { $info.MayPrompt = $true }
        }
        if ($info.MayPrompt) {
            $info.Reason = 'Private key provider may require a PIN or hardware token; unattended signing can block.'
        }
    }
    catch {
        $info.Reason = 'Private key metadata could not be read: {0}' -f $_.Exception.Message
    }
    finally {
        if ($ownsKey -and $key -is [System.IDisposable]) { $key.Dispose() }
    }

    return $info
}

function Get-SigningTokenFindings {
    <#
        Token-level inspection of a command line. PowerShell binds a parameter
        by unique prefix, so -e, -ex, -executionpo all reach -ExecutionPolicy
        and -e, -enc all reach -EncodedCommand.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $findings = New-Object System.Collections.ArrayList
    foreach ($raw in ($Text -split '[\s=:]+')) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        if ($raw[0] -ne '-' -and $raw[0] -ne '/') { continue }
        $name = $raw.TrimStart('-', '/').Trim('"', "'").ToLowerInvariant()
        if ([string]::IsNullOrEmpty($name)) { continue }
        if ('executionpolicy'.StartsWith($name) -or $name -eq 'ep') {
            [void]$findings.Add([pscustomobject]@{ Code = 'LauncherExecutionPolicy'; Token = $raw })
        }
        if ('encodedcommand'.StartsWith($name)) {
            [void]$findings.Add([pscustomobject]@{ Code = 'LauncherEncodedCommand'; Token = $raw })
        }
    }
    # The bare policy names only count on a line that names the execution
    # policy (Set-ExecutionPolicy, the preference variable, a split
    # argument); a comment or a string that merely contains the word is not
    # a launcher.
    if ($Text -match '(?i)executionpolicy' -and $Text -match '(?i)\b(bypass|unrestricted|remotesigned)\b') {
        [void]$findings.Add([pscustomobject]@{ Code = 'LauncherPolicyRelaxation'; Token = $Matches[1].ToLowerInvariant() })
    }
    return $findings.ToArray()
}

function Get-SigningThirdPartyRoots {
    <#
        Folders holding vendor toolkit content. A PSADT layout is identified
        by its entry script; the manifest may declare further third-party
        scripts by relative path.
    #>
    param([Parameter(Mandatory)][string]$StageRoot, $ManifestData)

    $roots = @()
    if (Test-Path -LiteralPath $StageRoot) {
        $roots = @(Get-ChildItem -LiteralPath $StageRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Invoke-AppDeployToolkit.ps1' -or $_.Name -eq 'Deploy-Application.ps1' } |
            ForEach-Object { $_.DirectoryName.TrimEnd('\') })
    }
    $declared = @()
    $declaredValue = Get-SigningPropertyValue -InputObject $ManifestData -Name 'ThirdPartyScripts' -Default $null
    if ($null -ne $declaredValue) { $declared = @($declaredValue | ForEach-Object { [string]$_ }) }
    return [pscustomobject]@{ Roots = @($roots); Declared = @($declared) }
}

function Test-SigningThirdPartyFile {
    param(
        [Parameter(Mandatory)][string]$FullName,
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)]$ThirdParty
    )
    foreach ($root in @($ThirdParty.Roots)) {
        if ($FullName.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    foreach ($name in @($ThirdParty.Declared)) {
        if ($RelativePath -ieq $name) { return $true }
    }
    return $false
}

function Write-SigningScriptFile {
    <#
        Materializes script text under the stage root. Text that already
        carries a signature block only verifies when the original bytes are
        reproduced, so each candidate encoding is written and verified and the
        first one that keeps the signature intact wins. Unsigned text is
        written as UTF-8 without a BOM.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [string]$EncodingName = ''
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $candidates = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($EncodingName)) {
        $named = switch ($EncodingName) {
            'UTF8BOM'   { New-Object System.Text.UTF8Encoding($true) }
            'UTF16LE'   { New-Object System.Text.UnicodeEncoding($false, $true) }
            'UTF16BE'   { New-Object System.Text.UnicodeEncoding($true, $true) }
            'ASCII'     { New-Object System.Text.ASCIIEncoding }
            'ANSI'      { [System.Text.Encoding]::Default }
            default     { $utf8NoBom }
        }
        [void]$candidates.Add($named)
    }
    [void]$candidates.Add($utf8NoBom)

    if ($Text -match '# SIG # Begin signature block') {
        [void]$candidates.Add((New-Object System.Text.UTF8Encoding($true)))
        [void]$candidates.Add((New-Object System.Text.UnicodeEncoding($false, $true)))
        [void]$candidates.Add((New-Object System.Text.UnicodeEncoding($true, $true)))

        foreach ($encoding in $candidates) {
            [System.IO.File]::WriteAllText($Path, $Text, $encoding)
            if ((Test-ScriptSignature -Path $Path).SignatureIntact) { return }
        }
    }

    [System.IO.File]::WriteAllText($Path, $Text, $candidates[0])
}

function Get-SigningEncodingInfo {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return [pscustomobject]@{ Name = 'UTF8BOM'; HasBom = $true; Encoding = (New-Object System.Text.UTF8Encoding($true)) }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return [pscustomobject]@{ Name = 'UTF16LE'; HasBom = $true; Encoding = (New-Object System.Text.UnicodeEncoding($false, $true)) }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
        return [pscustomobject]@{ Name = 'UTF16BE'; HasBom = $true; Encoding = (New-Object System.Text.UnicodeEncoding($true, $true)) }
    }

    $nonAscii = $false
    foreach ($b in $Bytes) { if ($b -gt 0x7F) { $nonAscii = $true; break } }
    if (-not $nonAscii) {
        return [pscustomobject]@{ Name = 'ASCII'; HasBom = $false; Encoding = (New-Object System.Text.ASCIIEncoding) }
    }
    # No BOM and high bytes present: decode strictly as UTF-8 and fall back to
    # the ANSI code page, which is what Windows PowerShell 5.1 assumes.
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$strict.GetString($Bytes)
        return [pscustomobject]@{ Name = 'UTF8NoBOM'; HasBom = $false; Encoding = (New-Object System.Text.UTF8Encoding($false)) }
    }
    catch {
        return [pscustomobject]@{ Name = 'ANSI'; HasBom = $false; Encoding = [System.Text.Encoding]::Default }
    }
}

# ---------------------------------------------------------------------------
# Policy
# ---------------------------------------------------------------------------

function Get-SigningPolicy {
    <#
    .SYNOPSIS
        Returns the normalized script-signing policy.

    .DESCRIPTION
        Precedence: -Policy, then -Json, then the run snapshot's SigningPolicy
        block, then APP_PACKAGER_SIGNING, then all-off defaults. Unknown or
        malformed input never silently enables a switch.
    #>
    [CmdletBinding()]
    param(
        $Policy,
        [string]$Json,
        $Snapshot
    )

    $source = $null
    if ($null -ne $Policy) {
        $source = $Policy
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Json)) {
        $source = ConvertFrom-Json $Json
    }
    elseif ($null -ne $Snapshot) {
        $source = Get-SigningPropertyValue -InputObject $Snapshot -Name 'SigningPolicy' -Default $null
    }
    elseif (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_SIGNING)) {
        try { $source = ConvertFrom-Json $env:APP_PACKAGER_SIGNING }
        catch { throw "APP_PACKAGER_SIGNING is not valid JSON: $($_.Exception.Message)" }
    }

    $result = [ordered]@{}
    foreach ($key in $script:DefaultSigningPolicy.Keys) {
        $default = $script:DefaultSigningPolicy[$key]
        $value = Get-SigningPropertyValue -InputObject $source -Name $key -Default $default
        if ($default -is [bool]) { $value = ConvertTo-SigningBoolean -Value $value -Default $default }
        else { $value = [string]$value }
        $result[$key] = $value
    }

    $result['StoreLocation'] = switch -Regex ($result['StoreLocation']) {
        '^(?i)localmachine$' { 'LocalMachine' }
        '^(?i)currentuser$'  { 'CurrentUser' }
        default              { throw "Unsupported signing StoreLocation '$($result['StoreLocation'])'. Use CurrentUser or LocalMachine." }
    }
    if ([string]::IsNullOrWhiteSpace($result['HashAlgorithm'])) { $result['HashAlgorithm'] = 'SHA256' }
    $result['HashAlgorithm'] = $result['HashAlgorithm'].ToUpperInvariant()
    $result['CertificateThumbprint'] = ($result['CertificateThumbprint'] -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()

    $canonical = ($result.Keys | ForEach-Object { '{0}={1}' -f $_, $result[$_] }) -join ';'
    $result['PolicyDigest'] = Get-SigningSha256 -Bytes ([System.Text.Encoding]::UTF8.GetBytes($canonical))

    return [pscustomobject]$result
}

function Get-CodeSigningCertificateCandidates {
    <#
    .SYNOPSIS
        Lists code-signing certificates without touching the private key.

    .DESCRIPTION
        Key accessibility is probed by opening a key handle and reading
        provider metadata; no signature is produced, so no PIN dialog appears.
        A hardware or PIN-protected provider is reported through MayPrompt.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser', 'LocalMachine')]
        [string]$StoreLocation = 'CurrentUser'
    )

    $storePath = 'Cert:\{0}\My' -f $StoreLocation
    $certs = @()
    try { $certs = @(Get-ChildItem -Path $storePath -ErrorAction Stop) }
    catch { return @() }

    $now = Get-Date
    $results = New-Object System.Collections.ArrayList
    foreach ($cert in $certs) {
        if (-not (Test-SigningCertificateIsCodeSigning -Certificate $cert)) { continue }

        $keyInfo = Get-SigningPrivateKeyInfo -Certificate $cert
        $reasons = New-Object System.Collections.ArrayList
        if ($cert.NotBefore -gt $now) { [void]$reasons.Add('Certificate is not yet valid.') }
        if ($cert.NotAfter -lt $now) { [void]$reasons.Add('Certificate expired on {0:yyyy-MM-dd}.' -f $cert.NotAfter) }
        if (-not $keyInfo.HasPrivateKey) { [void]$reasons.Add('Certificate has no associated private key.') }
        elseif (-not $keyInfo.Accessible) { [void]$reasons.Add($keyInfo.Reason) }
        if ($keyInfo.MayPrompt) { [void]$reasons.Add($keyInfo.Reason) }

        $usable = ($cert.NotBefore -le $now) -and ($cert.NotAfter -ge $now) -and $keyInfo.HasPrivateKey -and $keyInfo.Accessible

        [void]$results.Add([pscustomobject]@{
            Thumbprint    = $cert.Thumbprint.ToUpperInvariant()
            Subject       = $cert.Subject
            Issuer        = $cert.Issuer
            NotBefore     = $cert.NotBefore
            NotAfter      = $cert.NotAfter
            HasPrivateKey = [bool]$keyInfo.HasPrivateKey
            KeyAccessible = [bool]$keyInfo.Accessible
            MayPrompt     = [bool]$keyInfo.MayPrompt
            KeyProvider   = $keyInfo.Provider
            StoreLocation = $StoreLocation
            Usable        = [bool]$usable
            Reason        = ($reasons -join ' ')
        })
    }

    return $results.ToArray()
}

function Resolve-SigningCertificate {
    <#
    .SYNOPSIS
        Returns the configured signing certificate, selected by thumbprint.

    .DESCRIPTION
        Selection is by thumbprint only. A missing, expired, keyless,
        inaccessible or duplicated selection throws
        SigningCertificateUnavailable with the precise reason; nothing falls
        back to the first certificate in the store.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Policy)

    $policy = Get-SigningPolicy -Policy $Policy
    $thumb = $policy.CertificateThumbprint
    if ([string]::IsNullOrWhiteSpace($thumb)) {
        throw "SigningCertificateUnavailable: no certificate selected. Choose a code-signing certificate thumbprint in Options."
    }

    $storePath = 'Cert:\{0}\My' -f $policy.StoreLocation
    $found = @()
    try {
        $found = @(Get-ChildItem -Path $storePath -ErrorAction Stop |
            Where-Object { $_.Thumbprint -and $_.Thumbprint.ToUpperInvariant() -eq $thumb })
    }
    catch {
        throw "SigningCertificateUnavailable: certificate store $storePath could not be read ($($_.Exception.Message))."
    }

    if ($found.Count -eq 0) {
        throw "SigningCertificateUnavailable: thumbprint $thumb was not found in $storePath."
    }
    if ($found.Count -gt 1) {
        throw "SigningCertificateUnavailable: ambiguous configuration, thumbprint $thumb resolves to $($found.Count) entries in $storePath."
    }

    $cert = $found[0]
    $now = Get-Date
    if ($cert.NotBefore -gt $now) {
        throw "SigningCertificateUnavailable: certificate $thumb is not valid before $($cert.NotBefore.ToString('yyyy-MM-dd'))."
    }
    if ($cert.NotAfter -lt $now) {
        throw "SigningCertificateUnavailable: certificate $thumb expired on $($cert.NotAfter.ToString('yyyy-MM-dd')). Select a renewed certificate."
    }
    if (-not (Test-SigningCertificateIsCodeSigning -Certificate $cert)) {
        throw "SigningCertificateUnavailable: certificate $thumb does not carry the code-signing extended key usage."
    }

    $keyInfo = Get-SigningPrivateKeyInfo -Certificate $cert
    if (-not $keyInfo.HasPrivateKey) {
        throw "SigningCertificateUnavailable: certificate $thumb has no associated private key in $storePath."
    }
    if (-not $keyInfo.Accessible) {
        throw "SigningCertificateUnavailable: the private key for $thumb is not accessible to this account. $($keyInfo.Reason)"
    }

    return $cert
}

function Test-SigningConfiguration {
    <#
    .SYNOPSIS
        Signs and verifies a temporary script to prove the configuration works.

    .DESCRIPTION
        The signature attempt runs in a background job with a timeout so that a
        PIN-protected or hardware-held key blocks the job rather than the
        caller. Expected failures are returned, not thrown.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Policy,
        [int]$TimeoutSeconds = 30
    )

    $policy = Get-SigningPolicy -Policy $Policy
    $result = [pscustomobject]@{
        Ok                = $false
        Thumbprint        = $policy.CertificateThumbprint
        Subject           = ''
        Issuer            = ''
        NotAfter          = $null
        StoreLocation     = $policy.StoreLocation
        HashAlgorithm     = $policy.HashAlgorithm
        TimestampServer   = $policy.TimestampServer
        TimestampRequired = [bool]$policy.TimestampRequired
        SignatureStatus   = 'NotAttempted'
        SignatureIntact   = $false
        TrustedOnThisHost = $false
        StatusMessage     = ''
        Timestamped       = $false
        TimestampVerified = $false
        MayPrompt         = $false
        TimedOut          = $false
        Reason            = ''
    }

    $cert = $null
    try { $cert = Resolve-SigningCertificate -Policy $policy }
    catch {
        $result.Reason = $_.Exception.Message
        return $result
    }

    $result.Subject = $cert.Subject
    $result.Issuer = $cert.Issuer
    $result.NotAfter = $cert.NotAfter
    $keyInfo = Get-SigningPrivateKeyInfo -Certificate $cert
    $result.MayPrompt = [bool]$keyInfo.MayPrompt

    $temp = New-SigningTempScriptPath
    [System.IO.File]::WriteAllText($temp, "Write-Output 'AppPackager signing test'`r`nexit 0", (New-Object System.Text.ASCIIEncoding))

    $job = $null
    try {
        $job = Start-Job -ScriptBlock {
            param($Path, $Thumbprint, $StoreLocation, $HashAlgorithm, $TimestampServer)
            $c = Get-ChildItem ('Cert:\{0}\My' -f $StoreLocation) | Where-Object { $_.Thumbprint -eq $Thumbprint }
            $signArgs = @{ FilePath = $Path; Certificate = $c[0]; HashAlgorithm = $HashAlgorithm }
            if (-not [string]::IsNullOrWhiteSpace($TimestampServer)) { $signArgs['TimestampServer'] = $TimestampServer }
            $s = Set-AuthenticodeSignature @signArgs
            [pscustomobject]@{ Status = [string]$s.Status; Message = [string]$s.StatusMessage }
        } -ArgumentList $temp, $cert.Thumbprint, $policy.StoreLocation, $policy.HashAlgorithm, $policy.TimestampServer

        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $completed) {
            $result.TimedOut = $true
            $result.Reason = "Signing did not complete within $TimeoutSeconds seconds; the private key may require interactive confirmation."
            return $result
        }

        $jobOutput = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
        if ($job.State -eq 'Failed' -or $jobOutput.Count -eq 0) {
            $reason = 'Signing failed.'
            if ($job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].Error.Count -gt 0) {
                $reason = [string]$job.ChildJobs[0].Error[0]
            }
            $result.Reason = $reason
            return $result
        }

        $verify = Test-ScriptSignature -Path $temp
        $result.SignatureStatus = $verify.Status
        $result.SignatureIntact = $verify.SignatureIntact
        $result.TrustedOnThisHost = $verify.TrustedOnThisHost
        $result.StatusMessage = $verify.Reason
        $result.Timestamped = $verify.Timestamped
        $result.TimestampVerified = $verify.TimestampVerified

        if ($policy.TimestampRequired -and -not $verify.TimestampVerified) {
            $result.Reason = 'A timestamp is required but no verifiable timestamp counter-signature was produced.'
            return $result
        }
        if (-not $verify.SignatureIntact) {
            $result.Reason = "Signature verification returned $($verify.Status). $($verify.Reason)"
            return $result
        }
        if (-not $verify.TrustedOnThisHost) {
            $result.Reason = 'Signature is intact; the signer chain is not trusted on this host. Endpoints need the publisher trust chain.'
        }

        $result.Ok = $true
        return $result
    }
    finally {
        if ($null -ne $job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

function Test-ScriptSignature {
    <#
    .SYNOPSIS
        Verifies the Authenticode signature on a file.

    .DESCRIPTION
        Timestamped reports that a timestamp counter-signature exists;
        TimestampVerified reports that the timestamping authority certificate
        was returned by the trust provider. The two differ when the timestamp
        chain is not trusted on the signing host.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Valid = $false; SignatureIntact = $false; TrustedOnThisHost = $false
            Status = 'NotFound'; Thumbprint = ''
            Timestamped = $false; TimestampVerified = $false
            Reason = "File not found: $Path"
        }
    }

    $sig = Get-AuthenticodeSignature -LiteralPath $Path
    $thumb = ''
    if ($null -ne $sig.SignerCertificate) { $thumb = $sig.SignerCertificate.Thumbprint.ToUpperInvariant() }

    $timestampVerified = ($null -ne $sig.TimeStamperCertificate)
    $timestamped = $timestampVerified
    if (-not $timestamped) {
        $timestamped = Test-SignedScriptHasCounterSignature -Path $Path
    }

    # A signature whose hash matches but whose chain is not trusted by this
    # host is intact: the endpoint decides trust, and the build host is not
    # required to trust the signer. A hash or content failure reports its
    # own status, so UnknownError with a signer present is a chain-policy
    # failure; the chain is rebuilt with trust and revocation ignored rather
    # than read from the localized status text.
    $trusted = ($sig.Status -eq 'Valid')
    $intact = $trusted
    if (-not $intact -and $null -ne $sig.SignerCertificate -and $sig.Status -eq 'UnknownError') {
        $intact = Test-SigningChainIgnoringTrust -Certificate $sig.SignerCertificate
    }

    return [pscustomobject]@{
        Valid             = $trusted
        SignatureIntact   = $intact
        TrustedOnThisHost = $trusted
        Status            = [string]$sig.Status
        Thumbprint        = $thumb
        Timestamped       = $timestamped
        TimestampVerified = $timestampVerified
        Reason            = [string]$sig.StatusMessage
    }
}

function Test-SigningChainIgnoringTrust {
    <#
        Builds the signer chain with an unknown authority allowed and no
        revocation lookup: a self-signed or privately issued signer passes,
        an expired signer without a timestamp or an explicitly distrusted
        one does not, which matches what the endpoint's own policy decides
        on top of publisher trust.
    #>
    param([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    try {
        $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
        $chain.ChainPolicy.VerificationFlags = [System.Security.Cryptography.X509Certificates.X509VerificationFlags]::AllowUnknownCertificateAuthority
        return [bool]$chain.Build($Certificate)
    }
    catch { return $false }
    finally {
        if ($chain -is [System.IDisposable]) { $chain.Dispose() }
    }
}

function Test-SignedScriptHasCounterSignature {
    <#
        The signature block of a PowerShell script is a Base64 PKCS#7 blob.
        A timestamp is an unsigned counter-signature attribute inside it, so it
        is present even when the timestamping chain is not trusted and
        Get-AuthenticodeSignature therefore leaves TimeStamperCertificate null.
    #>
    param([Parameter(Mandatory)][string]$Path)

    try {
        $text = [System.IO.File]::ReadAllText($Path)
        $match = [regex]::Match($text, '(?s)# SIG # Begin signature block(.*?)# SIG # End signature block')
        if (-not $match.Success) { return $false }
        $b64 = ($match.Groups[1].Value -split "`n" |
            ForEach-Object { ($_ -replace '^\s*#\s?', '').Trim() }) -join ''
        if ([string]::IsNullOrWhiteSpace($b64)) { return $false }
        $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
        $cms.Decode([Convert]::FromBase64String($b64))
        foreach ($signer in $cms.SignerInfos) {
            if ($signer.CounterSignerInfos.Count -gt 0) { return $true }
            foreach ($attr in $signer.UnsignedAttributes) {
                # RFC 3161 timestamp token attribute.
                if ($attr.Oid.Value -eq '1.3.6.1.4.1.311.3.3.1') { return $true }
            }
        }
        return $false
    }
    catch { return $false }
}

function Test-ScriptSignatureBytes {
    <#
    .SYNOPSIS
        Verifies a signed script held as bytes by materializing it unchanged.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][byte[]]$Bytes)

    $temp = New-SigningTempScriptPath
    try {
        [System.IO.File]::WriteAllBytes($temp, $Bytes)
        $result = Test-ScriptSignature -Path $temp
        return [pscustomobject]@{
            Valid             = $result.Valid
            SignatureIntact   = $result.SignatureIntact
            TrustedOnThisHost = $result.TrustedOnThisHost
            Status            = $result.Status
            Thumbprint        = $result.Thumbprint
            Timestamped       = $result.Timestamped
            TimestampVerified = $result.TimestampVerified
            Sha256            = (Get-SigningSha256 -Bytes $Bytes)
            Reason            = $result.Reason
        }
    }
    finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}

function Get-SignedScriptRepresentation {
    <#
    .SYNOPSIS
        Returns the byte, Base64 and text forms of a script plus its encoding.

    .DESCRIPTION
        Text is decoded with the file's own encoding so that writing it back
        with that same encoding reproduces the original bytes. A text transport
        that re-encodes with a different encoding breaks the signature.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $enc = Get-SigningEncodingInfo -Bytes $bytes
    $text = $enc.Encoding.GetString($bytes)
    if ($enc.HasBom -and $text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }

    return [pscustomobject]@{
        Path             = $Path
        Bytes            = $bytes
        Base64           = [Convert]::ToBase64String($bytes)
        Text             = $text
        Encoding         = $enc.Name
        HasBom           = $enc.HasBom
        Sha256           = (Get-SigningSha256 -Bytes $bytes)
        Length           = $bytes.Length
        ExceedsCMLimit   = ($bytes.Length -gt $script:ConfigMgrDetectionScriptMaxBytes)
        CMScriptMaxBytes = $script:ConfigMgrDetectionScriptMaxBytes
    }
}

function Test-SignedScriptRoundTrip {
    <#
    .SYNOPSIS
        Proves that a transport preserves a script signature on this host.

    .DESCRIPTION
        Base64 and File transports move bytes. The Text transport decodes and
        re-encodes with the file's own encoding; it reproduces the bytes only
        when the original encoding is recovered exactly, which is why a
        BOM-less file carrying non-ASCII content needs the byte transport.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('Base64', 'Text', 'File')][string]$Transport,
        [System.Text.Encoding]$TextEncoding
    )

    $original = [System.IO.File]::ReadAllBytes($Path)
    $rep = Get-SignedScriptRepresentation -Path $Path
    $temp = New-SigningTempScriptPath
    try {
        switch ($Transport) {
            'Base64' {
                [System.IO.File]::WriteAllBytes($temp, [Convert]::FromBase64String($rep.Base64))
            }
            'File' {
                Copy-Item -LiteralPath $Path -Destination $temp -Force
            }
            'Text' {
                $encoding = $TextEncoding
                if ($null -eq $encoding) { $encoding = (Get-SigningEncodingInfo -Bytes $original).Encoding }
                [System.IO.File]::WriteAllText($temp, $rep.Text, $encoding)
            }
        }

        $roundTripped = [System.IO.File]::ReadAllBytes($temp)
        $identical = ($roundTripped.Length -eq $original.Length)
        if ($identical) {
            for ($i = 0; $i -lt $original.Length; $i++) {
                if ($roundTripped[$i] -ne $original[$i]) { $identical = $false; break }
            }
        }
        $verify = Test-ScriptSignature -Path $temp

        return [pscustomobject]@{
            Transport       = $Transport
            Valid           = $verify.Valid
            SignatureIntact = $verify.SignatureIntact
            Status          = $verify.Status
            Thumbprint      = $verify.Thumbprint
            BytesIdentical  = $identical
            Encoding       = $rep.Encoding
            Sha256         = (Get-SigningSha256 -Bytes $roundTripped)
            Reason         = $verify.Reason
        }
    }
    finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}

# ---------------------------------------------------------------------------
# Signing
# ---------------------------------------------------------------------------

function Invoke-ScriptSigning {
    <#
    .SYNOPSIS
        Signs one script in place for a signing category and verifies the result.

    .DESCRIPTION
        -PreserveExisting keeps an unchanged third-party file that already
        carries a valid signature; AppPackager-generated files are always
        re-signed. A configured timestamp that does not verify while
        TimestampRequired is set throws rather than producing a weaker result.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Policy,
        [Parameter(Mandatory)][ValidateSet('Detection', 'Requirements', 'Deployment')][string]$Category,
        [switch]$PreserveExisting,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $policy = Get-SigningPolicy -Policy $Policy
    $signSwitch = [bool]$policy.("Sign$Category")
    $requireSwitch = [bool]$policy.("Require$Category")

    $result = [pscustomobject]@{
        Path              = $Path
        Category          = $Category
        Status            = 'NotRequested'
        Thumbprint        = ''
        Sha256            = ''
        Timestamped       = $false
        TimestampVerified = $false
        Reason            = ''
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $result.Status = 'NotApplicable'
        $result.Reason = "File not found: $Path"
        return $result
    }

    if (-not $signSwitch -and -not $requireSwitch) {
        $result.Sha256 = Get-SigningFileSha256 -Path $Path
        return $result
    }

    $existing = Test-ScriptSignature -Path $Path
    if ($existing.SignatureIntact -and ($PreserveExisting -or -not $signSwitch)) {
        $result.Status = 'ExistingSignatureValid'
        $result.Thumbprint = $existing.Thumbprint
        $result.Timestamped = $existing.Timestamped
        $result.TimestampVerified = $existing.TimestampVerified
        $result.Sha256 = Get-SigningFileSha256 -Path $Path
        $result.Reason = 'Existing valid signature preserved.'
        return $result
    }

    if (-not $signSwitch) {
        # Require without automatic signing: unsigned or altered input fails.
        $result.Status = 'Failed'
        $result.Sha256 = Get-SigningFileSha256 -Path $Path
        $result.Reason = "A valid signature is required for $Category but the file verifies as $($existing.Status). $($existing.Reason)"
        return $result
    }

    $cert = $Certificate
    if ($null -eq $cert) { $cert = Resolve-SigningCertificate -Policy $policy }

    $signArgs = @{
        FilePath      = $Path
        Certificate   = $cert
        HashAlgorithm = $policy.HashAlgorithm
        ErrorAction   = 'Stop'
    }
    $timestampConfigured = -not [string]::IsNullOrWhiteSpace($policy.TimestampServer)
    if ($timestampConfigured) { $signArgs['TimestampServer'] = $policy.TimestampServer }

    try {
        $null = Set-AuthenticodeSignature @signArgs
    }
    catch {
        if ($policy.TimestampRequired) {
            throw "Signing $Category script '$Path' failed and a timestamp is required: $($_.Exception.Message)"
        }
        $result.Status = 'Failed'
        $result.Reason = $_.Exception.Message
        $result.Sha256 = Get-SigningFileSha256 -Path $Path
        return $result
    }

    $verify = Test-ScriptSignature -Path $Path
    $result.Thumbprint = $verify.Thumbprint
    $result.Timestamped = $verify.Timestamped
    $result.TimestampVerified = $verify.TimestampVerified
    $result.Sha256 = Get-SigningFileSha256 -Path $Path

    if ($policy.TimestampRequired -and -not $verify.Timestamped) {
        throw "TimestampRequired is set but the $Category script '$Path' carries no timestamp counter-signature (server '$($policy.TimestampServer)')."
    }

    if ($verify.SignatureIntact) {
        $result.Status = 'SignedAndVerified'
        if (-not $verify.TrustedOnThisHost) {
            $result.Reason = 'Signature is intact; the signer chain is not trusted on this host, which says nothing about endpoint trust.'
        }
    }
    else {
        $result.Status = 'Failed'
        $result.Reason = "Signature verification returned $($verify.Status). $($verify.Reason)"
    }
    return $result
}

# ---------------------------------------------------------------------------
# Launchers
# ---------------------------------------------------------------------------

function New-DeploymentLauncherCommand {
    <#
    .SYNOPSIS
        Builds the deployment launcher command line and BAT body for a script.

    .DESCRIPTION
        Signed mode omits every execution-policy argument so the endpoint's
        configured policy governs; it never substitutes Unrestricted, an alias
        or an encoded command. Unsigned mode reproduces the existing strings
        byte for byte.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Script,
        [Parameter(Mandatory)][bool]$Signed,
        [ValidateSet('x64', 'x86')][string]$ScriptHost = 'x64',
        [string]$BatExitCode = '%ERRORLEVEL%'
    )

    $exe = if ($ScriptHost -eq 'x86') {
        '%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    }
    else { 'PowerShell.exe' }

    $switches = if ($Signed) { '-NoProfile -NonInteractive' } else { '-NonInteractive -ExecutionPolicy Bypass' }

    $commandLine = '{0} {1} -File "{2}"' -f $exe, $switches, $Script
    $batInvoke = '{0} {1} -File "%~dp0{2}"' -f $exe, $switches, $Script

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('@echo off')
    [void]$lines.Add($batInvoke)
    if ($BatExitCode -ne '%ERRORLEVEL%') {
        [void]$lines.Add(('if %ERRORLEVEL% EQU 0 exit /b {0}' -f $BatExitCode))
    }
    [void]$lines.Add('exit /b %ERRORLEVEL%')

    return [pscustomobject]@{
        Script      = $Script
        Signed      = $Signed
        ScriptHost  = $ScriptHost
        Executable  = $exe
        CommandLine = $commandLine
        BatInvoke   = $batInvoke
        BatBody     = ($lines -join "`r`n")
    }
}

function Test-DeploymentLauncherChain {
    <#
    .SYNOPSIS
        Inventories the staged execution chain and reports policy findings.

    .DESCRIPTION
        Scans every .bat/.cmd/.ps1/.psm1 under the stage root plus the resolved
        entry commands. In signed deployment mode any finding throws with the
        file and line; in unsigned mode findings are informational.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        $Manifest,
        [Parameter(Mandatory)]$Policy
    )

    $policy = Get-SigningPolicy -Policy $Policy
    $findings = New-Object System.Collections.ArrayList

    $addFinding = {
        param($Code, $File, $LineNumber, $Line, $Message)
        [void]$findings.Add([pscustomobject]@{
            Code       = $Code
            File       = $File
            LineNumber = $LineNumber
            Line       = $Line
            Message    = $Message
        })
    }

    $files = @()
    $thirdPartySkipped = 0
    if (Test-Path -LiteralPath $StageRoot) {
        # Vendor toolkit content keeps its own signatures and its own internal
        # launches; the chain that is checked here is what AppPackager
        # generates plus the commands the deployment type runs.
        $thirdParty = Get-SigningThirdPartyRoots -StageRoot $StageRoot -ManifestData $Manifest
        $rootFull = (Get-Item -LiteralPath $StageRoot).FullName.TrimEnd('\')
        foreach ($file in @(Get-ChildItem -LiteralPath $StageRoot -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^(?i)\.(bat|cmd|ps1|psm1)$' })) {
            $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\')
            if (Test-SigningThirdPartyFile -FullName $file.FullName -RelativePath $relative -ThirdParty $thirdParty) {
                $thirdPartySkipped++
                continue
            }
            $files += $file
        }
    }

    foreach ($file in $files) {
        $lines = @()
        try { $lines = @([System.IO.File]::ReadAllLines($file.FullName)) } catch { continue }
        $isBatch = ($file.Extension -match '^(?i)\.(bat|cmd)$')
        $inSignatureBlock = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            # Every code line is token-scanned: a launcher can reach PowerShell
            # as pwsh.exe, through a variable, or through a cmd start, so the
            # literal host name is not a reliable gate. Comment lines and the
            # Authenticode block are skipped because neither is a command.
            if ($line -match '^\s*#\s*SIG # Begin signature block') { $inSignatureBlock = $true; continue }
            if ($line -match '^\s*#\s*SIG # End signature block') { $inSignatureBlock = $false; continue }
            if ($inSignatureBlock) { continue }
            if ($isBatch) { if ($line -match '^\s*(?i:rem)\b|^\s*::') { continue } }
            elseif ($line -match '^\s*#') { continue }
            foreach ($hit in (Get-SigningTokenFindings -Text $line)) {
                $message = switch ($hit.Code) {
                    'LauncherExecutionPolicy'  { "Launcher passes an execution-policy argument ('$($hit.Token)')." }
                    'LauncherEncodedCommand'   { "Launcher passes an encoded command ('$($hit.Token)')." }
                    default                    { "Launcher relaxes the execution policy ('$($hit.Token)')." }
                }
                & $addFinding $hit.Code $file.FullName ($i + 1) $line $message
            }
        }
    }

    $commands = New-Object System.Collections.ArrayList
    foreach ($name in 'InstallCommandLine', 'UninstallCommandLine') {
        $value = Get-SigningPropertyValue -InputObject $Manifest -Name $name -Default ''
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { [void]$commands.Add([pscustomobject]@{ Source = $name; Text = [string]$value }) }
    }
    $dts = Get-SigningPropertyValue -InputObject $Manifest -Name 'DeploymentTypes' -Default $null
    if ($null -ne $dts) {
        $index = 0
        foreach ($dt in @($dts)) {
            $index++
            foreach ($name in 'InstallCommandLine', 'UninstallCommandLine', 'InstallCommand', 'UninstallCommand') {
                $value = Get-SigningPropertyValue -InputObject $dt -Name $name -Default ''
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    [void]$commands.Add([pscustomobject]@{ Source = ('DeploymentTypes[{0}].{1}' -f $index, $name); Text = [string]$value })
                }
            }
        }
    }

    foreach ($command in $commands) {
        foreach ($hit in (Get-SigningTokenFindings -Text $command.Text)) {
            $code = if ($hit.Code -eq 'LauncherPolicyRelaxation') { 'CustomCommandBypass' } else { $hit.Code }
            & $addFinding $code $command.Source 0 $command.Text "Resolved command carries '$($hit.Token)'."
        }
    }

    $result = [pscustomobject]@{
        Findings          = $findings.ToArray()
        FilesInspected    = $files.Count
        ThirdPartySkipped = $thirdPartySkipped
        CommandsInspected = $commands.Count
        BypassFree        = ($findings.Count -eq 0)
        Enforced          = ([bool]$policy.SignDeployment -or [bool]$policy.RequireDeployment)
    }

    # RequireDeployment is a publishing constraint in its own right: a strict
    # environment that verifies signatures without producing them must still
    # refuse a launcher that relaxes the endpoint policy.
    if ($result.Enforced -and $findings.Count -gt 0) {

        $detail = ($findings | ForEach-Object { '{0} ({1}:{2})' -f $_.Message, $_.File, $_.LineNumber }) -join '; '
        throw "Signed deployment mode refuses this build: $detail"
    }

    return $result
}

# ---------------------------------------------------------------------------
# Category orchestration
# ---------------------------------------------------------------------------

function Get-SigningDeploymentFileInventory {
    param([Parameter(Mandatory)][string]$StageRoot, $ManifestData)

    $owned = New-Object System.Collections.ArrayList
    $thirdParty = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $StageRoot)) {
        return [pscustomobject]@{ Owned = @(); ThirdParty = @() }
    }

    # A PSADT layout is vendor content; its files keep their own signatures
    # unless the caller has edited them.
    $thirdPartyRoots = Get-SigningThirdPartyRoots -StageRoot $StageRoot -ManifestData $ManifestData
    $rootFull = (Get-Item -LiteralPath $StageRoot).FullName.TrimEnd('\')

    foreach ($file in (Get-ChildItem -LiteralPath $StageRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^(?i)\.(ps1|psm1)$' })) {
        $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\')
        # scripts\ holds the detection and requirement categories; the
        # deployment category must not re-sign them under its own switch.
        if ($relative -like 'scripts\*') { continue }
        $isThirdParty = Test-SigningThirdPartyFile -FullName $file.FullName -RelativePath $relative -ThirdParty $thirdPartyRoots
        if ($isThirdParty) { [void]$thirdParty.Add([pscustomobject]@{ FullName = $file.FullName; RelativePath = $relative }) }
        else { [void]$owned.Add([pscustomobject]@{ FullName = $file.FullName; RelativePath = $relative }) }
    }

    return [pscustomobject]@{ Owned = $owned.ToArray(); ThirdParty = $thirdParty.ToArray() }
}

function Invoke-CategorySigning {
    <#
    .SYNOPSIS
        Signs every selected category of a staged build and returns the
        ScriptSigning manifest block.

    .DESCRIPTION
        Detection and requirement scripts are materialized under the stage root
        before signing and the manifest then carries the exact signed text.
        A Require switch whose category does not reach ExistingSignatureValid
        or SignedAndVerified throws, so no partially signed build publishes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)]$ManifestData,
        [Parameter(Mandatory)]$Policy
    )

    $policy = Get-SigningPolicy -Policy $Policy
    $cert = $null
    if ($policy.SignDetection -or $policy.SignRequirements -or $policy.SignDeployment) {
        $cert = Resolve-SigningCertificate -Policy $policy
    }

    $block = [ordered]@{
        PolicyDigest = $policy.PolicyDigest
        Detection    = [ordered]@{ Status = 'NotApplicable'; Thumbprint = ''; Sha256 = ''; Timestamped = $false; File = $null; Reason = '' }
        Requirements = [ordered]@{ Status = 'NotApplicable'; Items = @() }
        Deployment   = [ordered]@{ Status = 'NotApplicable'; Files = @(); LaunchersBypassFree = $true }
    }

    # --- Detection -------------------------------------------------------
    $detection = Get-SigningPropertyValue -InputObject $ManifestData -Name 'Detection' -Default $null
    $scriptText = Get-SigningPropertyValue -InputObject $detection -Name 'ScriptText' -Default ''
    if (-not [string]::IsNullOrWhiteSpace([string]$scriptText)) {
        $scriptsDir = Join-Path $StageRoot 'scripts'
        if (-not (Test-Path -LiteralPath $scriptsDir)) { New-Item -ItemType Directory -Path $scriptsDir -Force | Out-Null }
        $detectPath = Join-Path $scriptsDir 'detect.ps1'
        # Always rewritten from the manifest: a re-stage at the same vendor
        # version finds the previous stage's file in place, and reusing it
        # would publish the superseded detection script.
        Write-SigningScriptFile -Path $detectPath -Text ([string]$scriptText) `
            -EncodingName ([string](Get-SigningPropertyValue -InputObject $detection -Name 'ScriptEncoding' -Default ''))

        $signed = Invoke-ScriptSigning -Path $detectPath -Policy $policy -Category 'Detection' -Certificate $cert
        $representation = Get-SignedScriptRepresentation -Path $detectPath
        $block.Detection.Status = $signed.Status
        $block.Detection.Thumbprint = $signed.Thumbprint
        $block.Detection.Sha256 = $signed.Sha256
        $block.Detection.Timestamped = $signed.Timestamped
        $block.Detection.File = 'scripts\detect.ps1'
        $block.Detection.Reason = $signed.Reason
        if ($representation.ExceedsCMLimit) {
            throw "Detection script is $($representation.Length) bytes after signing; ConfigMgr refuses a script larger than $($representation.CMScriptMaxBytes) bytes."
        }
        Set-SigningManifestDetection -Detection $detection -Text $representation.Text -File 'scripts\detect.ps1'
    }

    # --- Requirements ----------------------------------------------------
    $requirements = @(Get-SigningRequirementScripts -ManifestData $ManifestData)
    if ($requirements.Count -gt 0) {
        $reqDir = Join-Path (Join-Path $StageRoot 'scripts') 'requirements'
        if (-not (Test-Path -LiteralPath $reqDir)) { New-Item -ItemType Directory -Path $reqDir -Force | Out-Null }
        $items = New-Object System.Collections.ArrayList
        $statuses = New-Object System.Collections.ArrayList
        foreach ($req in $requirements) {
            $file = Join-Path $reqDir ('{0}.ps1' -f $req.RuleId)
            Write-SigningScriptFile -Path $file -Text $req.ScriptText -EncodingName $req.ScriptEncoding
            $signed = Invoke-ScriptSigning -Path $file -Policy $policy -Category 'Requirements' -Certificate $cert
            [void]$statuses.Add($signed.Status)
            [void]$items.Add([pscustomobject]@{
                RuleId     = $req.RuleId
                File       = 'scripts\requirements\{0}.ps1' -f $req.RuleId
                Sha256     = $signed.Sha256
                Thumbprint = $signed.Thumbprint
                Status     = $signed.Status
                Reason     = $signed.Reason
            })
        }
        $block.Requirements.Items = $items.ToArray()
        $block.Requirements.Status = Get-SigningAggregateStatus -Statuses $statuses.ToArray()
    }

    # --- Deployment ------------------------------------------------------
    $inventory = Get-SigningDeploymentFileInventory -StageRoot $StageRoot -ManifestData $ManifestData
    if ($inventory.Owned.Count -gt 0 -or $inventory.ThirdParty.Count -gt 0) {
        $files = New-Object System.Collections.ArrayList
        $statuses = New-Object System.Collections.ArrayList
        foreach ($entry in $inventory.Owned) {
            $signed = Invoke-ScriptSigning -Path $entry.FullName -Policy $policy -Category 'Deployment' -Certificate $cert
            [void]$statuses.Add($signed.Status)
            [void]$files.Add([pscustomobject]@{
                RelativePath = $entry.RelativePath
                Sha256       = $signed.Sha256
                Thumbprint   = $signed.Thumbprint
                Status       = $signed.Status
                Owner        = 'AppPackager'
                Reason       = $signed.Reason
            })
        }
        foreach ($entry in $inventory.ThirdParty) {
            $signed = Invoke-ScriptSigning -Path $entry.FullName -Policy $policy -Category 'Deployment' -PreserveExisting -Certificate $cert
            [void]$statuses.Add($signed.Status)
            [void]$files.Add([pscustomobject]@{
                RelativePath = $entry.RelativePath
                Sha256       = $signed.Sha256
                Thumbprint   = $signed.Thumbprint
                Status       = $signed.Status
                Owner        = 'ThirdParty'
                Reason       = $signed.Reason
            })
        }
        $block.Deployment.Files = $files.ToArray()
        $block.Deployment.Status = Get-SigningAggregateStatus -Statuses $statuses.ToArray()
    }

    $chain = Test-DeploymentLauncherChain -StageRoot $StageRoot -Manifest $ManifestData -Policy $policy
    $block.Deployment.LaunchersBypassFree = $chain.BypassFree

    foreach ($category in 'Detection', 'Requirements', 'Deployment') {
        if (-not $policy.("Require$category")) { continue }
        $status = [string]$block[$category].Status
        if ($status -ne 'SignedAndVerified' -and $status -ne 'ExistingSignatureValid') {
            throw "Require$category is set but the $category category status is '$status'; refusing to publish."
        }
    }

    return $block
}

function Get-SigningAggregateStatus {
    param([string[]]$Statuses)

    if ($null -eq $Statuses -or $Statuses.Count -eq 0) { return 'NotApplicable' }
    if ($Statuses -contains 'Failed') { return 'Failed' }
    if ($Statuses -contains 'SignedAndVerified') { return 'SignedAndVerified' }
    if ($Statuses -contains 'ExistingSignatureValid') { return 'ExistingSignatureValid' }
    if ($Statuses -contains 'NotRequested') { return 'NotRequested' }
    return 'NotApplicable'
}

function Get-SigningRequirementScripts {
    param($ManifestData)

    $results = New-Object System.Collections.ArrayList
    $requirements = Get-SigningPropertyValue -InputObject $ManifestData -Name 'Requirements' -Default $null
    if ($null -eq $requirements) { return $results.ToArray() }

    $index = 0
    foreach ($rule in @($requirements)) {
        $index++
        $text = Get-SigningPropertyValue -InputObject $rule -Name 'ScriptText' -Default ''
        if ([string]::IsNullOrWhiteSpace([string]$text)) { continue }
        $ruleId = [string](Get-SigningPropertyValue -InputObject $rule -Name 'RuleId' -Default ('rule{0}' -f $index))
        $ruleId = $ruleId -replace '[^0-9A-Za-z._-]', '-'
        [void]$results.Add([pscustomobject]@{
            RuleId         = $ruleId
            ScriptText     = [string]$text
            ScriptEncoding = [string](Get-SigningPropertyValue -InputObject $rule -Name 'ScriptEncoding' -Default '')
        })
    }
    return $results.ToArray()
}

function Set-SigningManifestDetection {
    param($Detection, [string]$Text, [string]$File)

    if ($null -eq $Detection) { return }
    if ($Detection -is [System.Collections.IDictionary]) {
        $Detection['ScriptText'] = $Text
        $Detection['ScriptFile'] = $File
        return
    }
    if ($Detection.PSObject.Properties['ScriptText']) { $Detection.ScriptText = $Text }
    if ($Detection.PSObject.Properties['ScriptFile']) { $Detection.ScriptFile = $File }
    else { Add-Member -InputObject $Detection -NotePropertyName 'ScriptFile' -NotePropertyValue $File -Force }
}

function Get-ConfigMgrDetectionScriptMaxBytes {
    <#
    .SYNOPSIS
        Returns the ConfigMgr detection script size limit in bytes.
    #>
    [CmdletBinding()]
    param()
    return $script:ConfigMgrDetectionScriptMaxBytes
}

Export-ModuleMember -Function @(
    'Get-SigningPolicy',
    'Get-CodeSigningCertificateCandidates',
    'Resolve-SigningCertificate',
    'Test-SigningConfiguration',
    'Invoke-ScriptSigning',
    'Test-ScriptSignature',
    'Test-ScriptSignatureBytes',
    'Get-SignedScriptRepresentation',
    'Test-SignedScriptRoundTrip',
    'New-DeploymentLauncherCommand',
    'Test-DeploymentLauncherChain',
    'Invoke-CategorySigning',
    'Get-ConfigMgrDetectionScriptMaxBytes'
)
