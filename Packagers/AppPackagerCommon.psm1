<#
.SYNOPSIS
    Shared module for AppPackager packager scripts.

.DESCRIPTION
    Import this module at the top of every packager script to get:
      - TLS 1.2 enforcement
      - Structured logging (Write-Log, Initialize-Logging)
      - Download with retry (Invoke-DownloadWithRetry)
      - Admin check (Test-IsAdmin)
      - ConfigMgr site connection (Connect-CMSite)
      - Folder initialization (Initialize-Folder)
      - Network share access test (Test-NetworkShareAccess)
      - Content wrapper generation (Write-ContentWrappers, New-MsiWrapperContent)
      - ConfigMgr application creation (New-MECMApplicationFromManifest)
      - CM revision history cleanup (Remove-CMApplicationRevisionHistoryByCIId)

.EXAMPLE
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
    Initialize-Logging -LogPath $LogPath

    Write-Log "Starting packager..."
    Write-Log "Something went wrong" -Level ERROR
    Invoke-DownloadWithRetry -Url $url -OutFile $file
#>

# ---------------------------------------------------------------------------
# Shared core
# ---------------------------------------------------------------------------

# SuiteCommon (vendored at ..\Lib\SuiteCommon) supplies logging, settings,
# and the CM connection core. Guarded, not -Force: a -Force reimport of
# this module (background runspace init) must not reset SuiteCommon's
# attached logging state.
if (-not (Get-Module -Name SuiteCommon)) {
    try {
        Import-Module -Name (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking -ErrorAction Stop
    }
    catch {
        # A blocked file (Mark-of-the-Web from a zip download) fails this
        # import while the caller keeps running, so every module function
        # then surfaces as an unrelated unknown-command error mid-run.
        throw ("AppPackagerCommon could not load its shared core (Lib\SuiteCommon). " +
            "If these files came from a downloaded zip, clear the Mark-of-the-Web first: " +
            "Get-ChildItem <app folder> -Recurse | Unblock-File. Underlying error: {0}" -f $_.Exception.Message)
    }
}
# Packager scripts and operators set APP_PACKAGER_VERBOSE; SuiteCommon
# gates DEBUG on SUITE_VERBOSE. Bridge once at import.
if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_VERBOSE) -and [string]::IsNullOrWhiteSpace($env:SUITE_VERBOSE)) {
    $env:SUITE_VERBOSE = $env:APP_PACKAGER_VERBOSE
}

# The workbench (build model, finalization hook, build records) and the
# signing service load beside this module and are imported -Global so the
# 285 packager scripts reach their functions without a per-packager import.
# Same guard shape as SuiteCommon: a -Force reimport of this module must not
# reset their attached state.
function Test-IsAppPackagerTestHost {
    <#
    .SYNOPSIS
        True when this import is a unit-test host rather than a deployment run.
    #>
    if ($env:APP_PACKAGER_TEST_HOST -eq '1') { return $true }
    return [bool](Get-Module -Name Pester)
}

$script:WorkbenchModulesAvailable = @{}
foreach ($companion in @(
        @{ Name = 'AppPackagerWorkbench'; Role = 'build model and stage finalization' },
        @{ Name = 'AppPackagerSigning';   Role = 'script signing service' })) {
    $script:WorkbenchModulesAvailable[$companion.Name] = $false
    if (Get-Module -Name $companion.Name) {
        $script:WorkbenchModulesAvailable[$companion.Name] = $true
        continue
    }
    $manifest = Join-Path $PSScriptRoot ('{0}.psd1' -f $companion.Name)
    if (Test-Path -LiteralPath $manifest -PathType Leaf) {
        try {
            Import-Module -Name $manifest -Global -DisableNameChecking -ErrorAction Stop
            $script:WorkbenchModulesAvailable[$companion.Name] = $true
            continue
        }
        catch {
            throw ("AppPackagerCommon could not load {0} ({1}) from '{2}'. If these files came from a downloaded zip, clear the Mark-of-the-Web first: Get-ChildItem <app folder> -Recurse | Unblock-File. Underlying error: {3}" -f $companion.Name, $companion.Role, $manifest, $_.Exception.Message)
        }
    }
    # Only the unit-test host may run without them: the module files are part
    # of every shipped install, so their absence anywhere else is a broken
    # installation that must fail now rather than at stage time.
    if (-not (Test-IsAppPackagerTestHost)) {
        throw ("AppPackagerCommon requires {0}.psd1 ({1}) beside it in '{2}'. Reinstall AppPackager; a partial copy of the Packagers folder cannot stage or package." -f $companion.Name, $companion.Role, $PSScriptRoot)
    }
}

# ---------------------------------------------------------------------------
# Workbench adapters (signing policy, plan digest, launcher commands, timing)
# ---------------------------------------------------------------------------

function Resolve-CommonSigningPolicy {
    <#
    .SYNOPSIS
        Returns the effective script-signing policy, or all-off defaults.

    .DESCRIPTION
        Get-SigningPolicy lives in the signing module. A host without it (unit
        tests, a partial checkout) resolves to every switch off so unsigned
        output stays byte-identical to the pre-signing behavior.
    #>
    $defaults = [pscustomobject]@{
        SignDetection = $false; SignRequirements = $false; SignDeployment = $false
        RequireDetection = $false; RequireRequirements = $false; RequireDeployment = $false
        CertificateThumbprint = ''; StoreLocation = 'CurrentUser'
        TimestampServer = ''; TimestampRequired = $false; HashAlgorithm = 'SHA256'
    }
    if (-not (Get-Command -Name Get-SigningPolicy -ErrorAction SilentlyContinue)) { return $defaults }
    try {
        $policy = Get-SigningPolicy
        if ($policy) { return $policy }
    }
    catch {
        Write-Log ("Signing policy unavailable; treating every signing switch as off: {0}" -f $_.Exception.Message) -Level WARN
    }
    return $defaults
}

function Get-PolicyFlag {
    <#
    .SYNOPSIS
        Reads one boolean off a signing policy object without assuming shape.
    #>
    param($Policy, [Parameter(Mandatory)][string]$Name)
    if (-not $Policy) { return $false }
    if ($Policy -is [hashtable]) { return [bool]$Policy[$Name] }
    $property = $Policy.PSObject.Properties[$Name]
    if (-not $property) { return $false }
    return [bool]$property.Value
}

function ConvertTo-CanonicalManifestObject {
    <#
    .SYNOPSIS
        Deep copy of manifest data with every mapping's keys ordered by name.

    .DESCRIPTION
        The plan digest must not change when an unrelated caller inserts a key
        in a different order, so hashtables and objects are rewritten as
        ordered dictionaries sorted with the invariant ordinal comparer.
        Arrays keep their order: deployment-type and clause order is meaning.
    #>
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return $Value }

    if ($Value -is [System.Collections.IDictionary]) {
        $ordered = [ordered]@{}
        foreach ($key in @($Value.Keys | Sort-Object -CaseSensitive:$false)) {
            $ordered[[string]$key] = ConvertTo-CanonicalManifestObject -Value $Value[$key]
        }
        return $ordered
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        return @(foreach ($item in $Value) { ConvertTo-CanonicalManifestObject -Value $item })
    }
    if ($Value -is [pscustomobject]) {
        $ordered = [ordered]@{}
        foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) {
            $ordered[$property.Name] = ConvertTo-CanonicalManifestObject -Value $property.Value
        }
        return $ordered
    }
    return $Value
}

function Get-StageManifestPlanDigest {
    <#
    .SYNOPSIS
        SHA-256 over the canonical JSON of a manifest's resolved plan.

    .DESCRIPTION
        FileHashes covers the payload bytes; stage-manifest.json is excluded
        from its own hash list, so a metadata-only change (commands, detection,
        timing, signing status) would otherwise leave no integrity record.
        FileHashes, PlanDigest and StagedAt are excluded: the first is already
        covered, the second is the output, and the third changes on every run.
    #>
    param([Parameter(Mandatory)]$ManifestData)

    $copy = [ordered]@{}
    $names = if ($ManifestData -is [System.Collections.IDictionary]) { @($ManifestData.Keys) } else { @($ManifestData.PSObject.Properties.Name) }
    foreach ($name in $names) {
        if (@('FileHashes', 'PlanDigest', 'StagedAt') -contains [string]$name) { continue }
        $copy[[string]$name] = if ($ManifestData -is [System.Collections.IDictionary]) { $ManifestData[$name] } else { $ManifestData.$name }
    }
    $canonical = ConvertTo-CanonicalManifestObject -Value $copy
    $json = $canonical | ConvertTo-Json -Depth 12
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($json))
    }
    finally {
        $sha.Dispose()
    }
    return (($bytes | ForEach-Object { $_.ToString('X2') }) -join '')
}

function Get-DeploymentLauncherCommandLine {
    <#
    .SYNOPSIS
        The command line a .bat shim or fallback launcher uses to run a .ps1.

    .DESCRIPTION
        Unsigned mode reproduces the historical string byte for byte. Signed
        mode drops -ExecutionPolicy Bypass so the endpoint's configured policy
        governs, and adds -NoProfile because the detection/deployment chain
        must not inherit a profile. The signing module owns the authoritative
        strings; a returned command that does not name the requested script is
        rejected rather than written into deployed content.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptToken,
        [switch]$Signed,
        [ValidateSet('x64', 'x86')][string]$ScriptHost = 'x64',
        [string]$ExecutableName = 'PowerShell.exe'
    )

    $executable = if ($ScriptHost -eq 'x86') { '%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' } else { $ExecutableName }
    $fallback = if ($Signed) {
        '{0} -NoProfile -NonInteractive -File "{1}"' -f $executable, $ScriptToken
    }
    else {
        '{0} -NonInteractive -ExecutionPolicy Bypass -File "{1}"' -f $executable, $ScriptToken
    }

    # A caller that pins the executable spelling (the PSADT fallback keeps the
    # toolkit's own lowercase form) owns the whole string; the service builds
    # its own executable name.
    if ($ExecutableName -cne 'PowerShell.exe') { return $fallback }
    if (-not (Get-Command -Name New-DeploymentLauncherCommand -ErrorAction SilentlyContinue)) { return $fallback }
    try {
        $produced = [string]((New-DeploymentLauncherCommand -Script $ScriptToken -Signed ([bool]$Signed) -ScriptHost $ScriptHost).CommandLine)
    }
    catch {
        Write-Log ("Launcher command service unavailable; using the built-in launcher string: {0}" -f $_.Exception.Message) -Level WARN
        return $fallback
    }
    if ([string]::IsNullOrWhiteSpace($produced) -or $produced -notlike ("*" + $ScriptToken + "*")) {
        Write-Log "Launcher command service returned a command that does not name the requested script; using the built-in launcher string." -Level WARN
        return $fallback
    }
    return $produced
}

function Resolve-DeploymentTypeTiming {
    <#
    .SYNOPSIS
        Validates and resolves estimated/maximum runtime minutes.

    .DESCRIPTION
        Manifest Timing wins over the caller's parameters. Both values are
        whole minutes greater than zero; the estimate may not exceed the
        maximum. A maximum above the documented 12-hour monitoring ceiling is
        accepted but reported, because a maximum longer than a maintenance
        window prevents the deployment type from running at all.

    .OUTPUTS
        [pscustomobject] EstimatedRuntimeMins, MaximumRuntimeMins, Source.
    #>
    param(
        $Timing,
        [int]$DefaultEstimated = 15,
        [int]$DefaultMaximum = 30,
        [string]$Context = 'deployment type'
    )

    $estimated = $DefaultEstimated
    $maximum = $DefaultMaximum
    $source = 'parameters'

    $read = {
        param($Object, $Name)
        if (-not $Object) { return $null }
        if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
        $property = $Object.PSObject.Properties[$Name]
        if (-not $property) { return $null }
        return $property.Value
    }

    $manifestEstimated = & $read $Timing 'EstimatedMinutes'
    $manifestMaximum = & $read $Timing 'MaximumMinutes'
    if ($null -ne $manifestEstimated -and '' -ne [string]$manifestEstimated) { $estimated = [int]$manifestEstimated; $source = 'manifest' }
    if ($null -ne $manifestMaximum -and '' -ne [string]$manifestMaximum) { $maximum = [int]$manifestMaximum; $source = 'manifest' }

    foreach ($pair in @(@{ Name = 'estimated installation time'; Value = $estimated }, @{ Name = 'maximum allowed run time'; Value = $maximum })) {
        if ($pair.Value -lt 1 -or $pair.Value -gt 1440) {
            throw ("Timing for the $Context is out of range: {0} is {1} minutes; use a whole number of minutes from 1 to 1440." -f $pair.Name, $pair.Value)
        }
    }
    if ($estimated -gt $maximum) {
        throw ("Timing for the $Context is inconsistent: estimated installation time {0} minutes exceeds the maximum allowed run time {1} minutes." -f $estimated, $maximum)
    }
    if ($maximum -gt 720) {
        Write-Log ("Maximum allowed run time      : {0} minutes exceeds the 12-hour monitoring ceiling; a maintenance window shorter than this blocks the deployment type." -f $maximum) -Level WARN
    }

    return [pscustomobject]@{
        EstimatedRuntimeMins = $estimated
        MaximumRuntimeMins   = $maximum
        Source               = $source
    }
}

# ---------------------------------------------------------------------------
# Download with retry
# ---------------------------------------------------------------------------

function Get-StoredGitHubToken {
    <#
    .SYNOPSIS
        Returns the GitHub token that Options, Tools saved, or '' when there is none.

    .DESCRIPTION
        The token sits in AppPackager.preferences.json as DPAPI text for the current Windows user.
        Another account or another computer cannot read it, and the function then returns ''.
        The token is never written to a log.
    #>
    $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'AppPackager.preferences.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    try {
        $data = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $protected = [string]$data.GitHub.TokenProtected
        if ([string]::IsNullOrWhiteSpace($protected)) { return '' }
        $secure = ConvertTo-SecureString -String $protected -ErrorAction Stop
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    catch { return '' }
}

function Select-GitHubReleaseWithAsset {
    <#
    .SYNOPSIS
        The newest published release that carries an installer asset, read from the asset list of each release.
    .DESCRIPTION
        releases/latest names the newest release, but a release can ship without the installer that a packager needs
        (PowerShell 7.7.0 publishes no MSI). This reads the assets and never builds a file name from the version.
        A draft or a pre-release is not a candidate. The result holds the release, the asset, the version from the tag,
        the download URL of the asset and the versions of the newer releases that were passed over. Returns $null when no
        release carries the asset.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][AllowNull()][object[]]$Releases,
        [Parameter(Mandatory)][string]$AssetPattern
    )
    # Windows PowerShell 5.1 returns a JSON array as one object that holds the array; flatten any nesting.
    $flat = @(); foreach ($entry in @($Releases)) { $flat += $entry }
    $ordered = @($flat | Where-Object { $_ -and -not [bool]$_.draft -and -not [bool]$_.prerelease } |
            Sort-Object -Property @{ Expression = { $d = [datetime]::MinValue; if ([datetime]::TryParse([string]$_.published_at, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$d)) { $d } else { [datetime]::MinValue } }; Descending = $true })
    $skipped = New-Object System.Collections.Generic.List[string]
    foreach ($release in $ordered) {
        $asset = @($release.assets | Where-Object { $_ -and ([string]$_.name) -match $AssetPattern }) | Select-Object -First 1
        $version = ([string]$release.tag_name) -replace '^v', ''
        if ($asset) {
            return [pscustomobject]@{
                Release = $release; Asset = $asset; Version = $version; Tag = [string]$release.tag_name
                AssetName = [string]$asset.name; DownloadUrl = [string]$asset.browser_download_url; SkippedNewer = @($skipped)
            }
        }
        $skipped.Add($version)
    }
    return $null
}
Export-ModuleMember -Function Select-GitHubReleaseWithAsset

function Get-CurrentNewerNote {
    <#
    .SYNOPSIS
        The note for a product whose installed version is newer than the version of the catalog, when that is expected.
    .DESCRIPTION
        PowerShell 7 devices that were installed from the MSI take newer versions through Microsoft Update (the MSI property
        USE_MU), also when the release of the newer version publishes no MSI. For such a product a version newer than the
        catalog is compliant. Returns an empty string for any other packager.
    #>
    param([AllowEmptyString()][AllowNull()][string]$PackagerScript)
    if ([string]::IsNullOrWhiteSpace($PackagerScript)) { return '' }
    switch -Regex ([System.IO.Path]::GetFileName($PackagerScript)) {
        '^package-powershell7\.ps1$' { return 'Newer versions arrive through Microsoft Update, so a newer installed version is compliant.' }
        default { return '' }
    }
}
Export-ModuleMember -Function Get-CurrentNewerNote
function Assert-DownloadAvailable {
    <#
    .SYNOPSIS
        Stops with a clear message when the installer file is not published at the URL, before any download starts.
    .DESCRIPTION
        A version feed or a release tag can run ahead of the file (a release without the installer, a file that is not
        uploaded yet). A range request of the first byte answers without a download. A status other than 200 or 206 stops the run.
    #>
    param([Parameter(Mandatory)][string]$Url, [string]$What = 'The installer file')
    $code = ''
    try { $code = ([string]((curl.exe -L --silent --range 0-0 --output NUL --write-out '%{http_code}' --max-time 60 --max-redirs 10 $Url) -join '')).Trim() } catch { $code = 'error' }
    if ($code -notin @('200', '206')) {
        throw ("{0} is not published at {1} (HTTP status {2}). The version was announced before the file was published, or the file moved. Nothing was downloaded." -f $What, $Url, $(if ($code) { $code } else { 'none' }))
    }
    Write-Log ("Download check               : {0} answers HTTP {1}" -f $Url, $code)
}
Export-ModuleMember -Function Assert-DownloadAvailable

function Assert-InstallerSignature {
    <#
    .SYNOPSIS
        Checks the Authenticode signature of a downloaded installer before it is staged: the status must be Valid and the
        signer must be the expected publisher.
    .DESCRIPTION
        For a vendor that publishes no digest but signs the installer. The signer is matched against the CN of the signing
        certificate (a whole name, case-insensitive). With -RemoveOnFailure a file that fails is deleted, so that the next
        run downloads it again instead of reusing a bad file. The message names the file, the status or the signer, and says
        that nothing was staged.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSigner,
        [string]$What = '',
        [switch]$RemoveOnFailure
    )
    $name = if ($What) { $What } else { [System.IO.Path]::GetFileName($Path) }
    $fail = {
        param([string]$Message)
        if ($RemoveOnFailure) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        throw $Message
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { & $fail "The installer $name is missing, so its signature cannot be checked. Nothing was staged." }
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid') {
        & $fail ("The installer {0} has no valid Authenticode signature (status: {1}). It may be damaged or not from the vendor. Nothing was staged." -f $name, $signature.Status)
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $cn = ''
    if ($subject -match '(?:^|,\s*)CN=(?:"(?<q>[^"]+)"|(?<n>[^,]+))') { $cn = if ($Matches['q']) { $Matches['q'] } else { $Matches['n'] } }
    if ($cn.Trim() -ne $ExpectedSigner) {
        & $fail ("The installer {0} is signed by '{1}', not by '{2}'. Nothing was staged." -f $name, $cn.Trim(), $ExpectedSigner)
    }
    return $cn.Trim()
}

function Assert-InstallerDigest {
    <#
    .SYNOPSIS
        Compares the SHA-256 hash of a downloaded installer with the digest that the vendor publishes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Sha256,
        [string]$What = '',
        [switch]$RemoveOnFailure
    )
    $name = if ($What) { $What } else { [System.IO.Path]::GetFileName($Path) }
    $expected = ($Sha256 -replace '^sha256:', '').Trim().ToUpperInvariant()
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne $expected) {
        if ($RemoveOnFailure) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        throw ("The installer {0} does not match the SHA-256 digest of the vendor (expected {1}, got {2}). Nothing was staged." -f $name, $expected, $actual)
    }
    return $actual
}
Export-ModuleMember -Function Assert-InstallerSignature, Assert-InstallerDigest

function Select-GitHubReleaseByAsset {
    <#
    .SYNOPSIS
        The newest published release whose asset list passes the filter of the packager; the release object itself.
    .DESCRIPTION
        A packager that reads releases/latest stops when the newest release ships without its installer. This reads the list of
        releases and returns the newest published release (no draft, no pre-release) that has an asset which passes
        -AssetFilter. The filter is the test of the packager on one asset ($_ is the asset). When no release passes, the newest
        published release is returned, so the packager reports its own "asset not found" message. The release objects are
        returned as they are; the code of the packager that follows does not change. A passed-over newer release is named
        in the log.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][AllowNull()][object[]]$Releases,
        [Parameter(Mandatory)][scriptblock]$AssetFilter
    )
    $flat = @(); foreach ($entry in @($Releases)) { $flat += $entry }
    $ordered = @($flat | Where-Object { $_ -and -not [bool]$_.draft -and -not [bool]$_.prerelease } |
            Sort-Object -Property @{ Expression = { $d = [datetime]::MinValue; if ([datetime]::TryParse([string]$_.published_at, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$d)) { $d } else { [datetime]::MinValue } }; Descending = $true })
    if ($ordered.Count -eq 0) { return $null }
    $skipped = New-Object System.Collections.Generic.List[string]
    foreach ($release in $ordered) {
        $match = @($release.assets | Where-Object $AssetFilter) | Select-Object -First 1
        if ($match) {
            if ($skipped.Count -gt 0) { Write-Log ("Release {0} has no matching installer; using {1}." -f ($skipped -join ', '), [string]$release.tag_name) -Level WARN }
            return $release
        }
        $skipped.Add([string]$release.tag_name)
    }
    return $ordered[0]
}
Export-ModuleMember -Function Select-GitHubReleaseByAsset
function Select-GitHubReleaseJson {
    <#
    .SYNOPSIS
        The JSON text of the newest published release that matches a pattern; for packagers that read the release as text.
    .DESCRIPTION
        Same rule as Select-GitHubReleaseByAsset, for a packager that matches a regular expression against the JSON text of
        the release. The pattern is the pattern of the packager. When no release matches, the newest published release is
        returned, so the packager reports its own message.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Json, [Parameter(Mandatory)][string]$Pattern)
    $parsed = ConvertFrom-Json $Json
    $flat = @(); foreach ($entry in @($parsed)) { $flat += $entry }
    $ordered = @($flat | Where-Object { $_ -and -not [bool]$_.draft -and -not [bool]$_.prerelease } |
            Sort-Object -Property @{ Expression = { $d = [datetime]::MinValue; if ([datetime]::TryParse([string]$_.published_at, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$d)) { $d } else { [datetime]::MinValue } }; Descending = $true })
    if ($ordered.Count -eq 0) { return $Json }
    $skipped = New-Object System.Collections.Generic.List[string]
    foreach ($release in $ordered) {
        $text = ConvertTo-Json -InputObject $release -Depth 20 -Compress
        if ([regex]::IsMatch($text, $Pattern)) {
            if ($skipped.Count -gt 0) { Write-Log ("Release {0} has no matching installer; using {1}." -f ($skipped -join ', '), [string]$release.tag_name) -Level WARN }
            return $text
        }
        $skipped.Add([string]$release.tag_name)
    }
    return (ConvertTo-Json -InputObject $ordered[0] -Depth 20 -Compress)
}
Export-ModuleMember -Function Select-GitHubReleaseJson
function Get-GitHubApiCurlArgs {
    <#
    .SYNOPSIS
        Returns the curl.exe arguments that authenticate a GitHub REST API call.
    .DESCRIPTION
        The unauthenticated GitHub API allows 60 requests per hour per source
        address; a version sweep over the catalog issues more than that. The
        token comes from GITHUB_TOKEN, then GH_TOKEN, then the token that
        Options, Tools saved, then the GitHub CLI's stored sign-in (gh auth
        token) when gh.exe is on the PATH and signed in; with a token the call carries it as a bearer token and the limit
        is 5000 per hour. Returns an empty array when no token is available,
        so callers can splat the result unconditionally. The CLI lookup runs
        once per process.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the argument list for one call.')]
    param()

    $token = $env:GITHUB_TOKEN
    if ([string]::IsNullOrWhiteSpace($token)) { $token = $env:GH_TOKEN }
    if ([string]::IsNullOrWhiteSpace($token)) { $token = Get-StoredGitHubToken }
    if ([string]::IsNullOrWhiteSpace($token)) {
        if ($null -eq $script:GitHubCliToken) {
            $script:GitHubCliToken = ''
            $gh = Get-Command -Name 'gh.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($gh) {
                try {
                    $cliToken = (& $gh.Source auth token 2>$null | Select-Object -First 1)
                    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($cliToken)) { $script:GitHubCliToken = [string]$cliToken }
                }
                catch { $script:GitHubCliToken = '' }
            }
        }
        $token = $script:GitHubCliToken
    }
    if ([string]::IsNullOrWhiteSpace($token)) { return @() }
    return @('-H', ('Authorization: Bearer ' + $token.Trim()), '-H', 'X-GitHub-Api-Version: 2022-11-28')
}

function ConvertTo-SafeCurlUrl {
    <#
    .SYNOPSIS
        Returns a URL that is safe to pass to curl.exe as one argument from
        Windows PowerShell 5.1, and throws for any other text.

    .DESCRIPTION
        Windows PowerShell 5.1 does not escape a double quote inside an
        argument it passes to a native command. A quote in a URL taken from a
        vendor page ends the argument, and the rest of the text then acts as
        further curl options, such as an output path. A scheme other than
        http or https would also fall outside the downloads the packagers
        make, and text that starts with a dash would be read as an option.
        The function refuses a URL with a quote character, a control
        character, a trailing backslash, or a scheme other than http or
        https, and returns the URL without surrounding white space.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Url)

    $text = $Url.Trim()
    if ($text -notmatch '^(?i)https?://[^\s]') {
        throw ("The URL is not an http or https address: {0}" -f $text)
    }
    if ($text -match '["\x00-\x1F\x7F]' -or $text.EndsWith('\')) {
        throw ("The URL carries a quote, control character or trailing backslash and is not passed to curl.exe: {0}" -f ($text -replace '[\x00-\x1F\x7F]', '?'))
    }
    return $text
}

function Invoke-DownloadWithRetry {
    <#
    .SYNOPSIS
        Downloads a file through curl.exe, falling back to Invoke-WebRequest, with
        a single retry on failure.

    .DESCRIPTION
        curl.exe is primary (curl.exe -L --fail --silent --show-error -o <file> <url>):
        the in-box Schannel build trusts the Windows certificate store and
        negotiates modern TLS regardless of per-machine .NET registry state.

        When curl fails, the attempt falls back to Invoke-WebRequest, which
        discovers WinINET/system proxy settings curl does not read. The
        fallback sends default credentials to the system proxy (Kerberos/NTLM
        auth proxies) and suppresses the 5.1 progress bar that cripples
        download throughput. TLS 1.2 is forced at module import.

        Throws on final failure (both methods, all attempts).

        Does NOT wrap scraping/variable-capture calls or URL-resolution calls.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [string[]]$ExtraCurlArgs = @(),

        [int]$RetryCount = 1,

        [int]$RetryDelaySec = 5,

        [switch]$Quiet
    )

    $Url = ConvertTo-SafeCurlUrl $Url

    $maxAttempts = 1 + $RetryCount

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Retrying download (attempt {0} of {1}) after {2}s delay..." -f $attempt, $maxAttempts, $RetryDelaySec) -Level WARN -Quiet:$Quiet
            Start-Sleep -Seconds $RetryDelaySec
        }

        $allArgs = @('-L', '--fail', '--silent', '--show-error') + $ExtraCurlArgs + @('-o', $OutFile, $Url)
        & curl.exe @allArgs 2>$null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            return
        }

        Write-Log ("curl.exe failed (exit code {0}); falling back to Invoke-WebRequest." -f $exitCode) -Level WARN -Quiet:$Quiet

        # curl leaves a partial file when the transfer dies mid-stream; remove
        # it so a fallback/retry never passes integrity checks with torn content.
        if (Test-Path -LiteralPath $OutFile) {
            Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
        }

        $savedProgress = $ProgressPreference
        try {
            # Authenticated system proxies (the environments where curl fails)
            # reject anonymous CONNECT; hand the default proxy our credentials.
            [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

            # 5.1 redraws the progress bar per buffer read; silencing it is the
            # difference between KB/s and full line speed on large installers.
            $ProgressPreference = 'SilentlyContinue'

            Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop

            Write-Log "Download succeeded through Invoke-WebRequest fallback." -Quiet:$Quiet
            return
        }
        catch {
            Write-Log ("Invoke-WebRequest fallback failed: {0}" -f $_.Exception.Message) -Level WARN -Quiet:$Quiet
            if (Test-Path -LiteralPath $OutFile) {
                Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
            }
        }
        finally {
            $ProgressPreference = $savedProgress
        }

        if ($attempt -lt $maxAttempts) {
            Write-Log ("Download attempt {0} failed through both curl.exe and Invoke-WebRequest. Will retry." -f $attempt) -Level WARN -Quiet:$Quiet
        }
    }

    $msg = "Download failed after $maxAttempts attempt(s) via both curl.exe and Invoke-WebRequest: $Url"
    Write-Log $msg -Level ERROR -Quiet:$Quiet
    throw $msg
}

function Invoke-CachedDownload {
    <#
    .SYNOPSIS
        Downloads a fixed-name file once, then refreshes it only when the
        server reports a newer copy.

    .DESCRIPTION
        For vendor URLs that always serve the current release under one file
        name. The first call downloads; later calls send the cached file's
        timestamp as If-Modified-Since, so an unchanged release costs one
        round trip and a new release replaces the cache. The saved file keeps
        the server's Last-Modified time so the condition stays exact. A
        failed refresh keeps the cached file and logs a warning.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [switch]$Quiet,
        [string]$DownloadRoot = $env:APP_PACKAGER_DOWNLOAD_ROOT
    )

    # Per-profile staging points -DownloadRoot at a profile subfolder, which
    # would otherwise re-download the same vendor file once per profile. The
    # shared cache keeps one copy keyed by URL plus file name; the profile's
    # own path still receives the file, so callers see today's layout. Without
    # a known download root there is no shared cache and the file is fetched
    # into its own path exactly as before.
    $cacheRoot = Get-SharedDownloadCacheRootSafe -DownloadRoot $DownloadRoot
    if ($cacheRoot) {
        $cached = Join-Path $cacheRoot (Get-SharedDownloadCacheKey -Url $Url -FileName (Split-Path -Leaf $OutFile))
        Invoke-CachedDownloadToPath -Url $Url -OutFile $cached -Quiet:$Quiet
        Copy-Item -LiteralPath $cached -Destination $OutFile -Force -ErrorAction Stop
        # The If-Modified-Since condition reads the file's timestamp, so the
        # copy must carry the server's Last-Modified time, not the copy time.
        (Get-Item -LiteralPath $OutFile).LastWriteTimeUtc = (Get-Item -LiteralPath $cached).LastWriteTimeUtc
        return
    }
    Invoke-CachedDownloadToPath -Url $Url -OutFile $OutFile -Quiet:$Quiet
}

function Get-SharedDownloadCacheRootSafe {
    <#
    .SYNOPSIS
        The shared download cache folder, or $null when none is configured.
    #>
    param([AllowEmptyString()][AllowNull()][string]$DownloadRoot)

    if ([string]::IsNullOrWhiteSpace($DownloadRoot)) { return $null }
    if (-not (Get-Command -Name Get-SharedDownloadCacheRoot -ErrorAction SilentlyContinue)) { return $null }
    try {
        $root = [string](Get-SharedDownloadCacheRoot -DownloadRoot $DownloadRoot)
    }
    catch {
        Write-Log ("Shared download cache unavailable; caching per download root: {0}" -f $_.Exception.Message) -Level WARN
        return $null
    }
    if ([string]::IsNullOrWhiteSpace($root)) { return $null }
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop | Out-Null }
    return $root
}

function Get-SharedDownloadCacheKey {
    <#
    .SYNOPSIS
        Cache file name for a URL plus its download file name.

    .DESCRIPTION
        Two vendors can serve the same file name, and one vendor can serve
        different content under one name per channel, so the key carries a
        digest of the URL as well as the readable file name.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$FileName
    )
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Url))
    }
    finally {
        $sha.Dispose()
    }
    $digest = (($bytes[0..7] | ForEach-Object { $_.ToString('x2') }) -join '')
    return ('{0}-{1}' -f $digest, $FileName)
}

function Invoke-CachedDownloadToPath {
    <#
    .SYNOPSIS
        The conditional download itself, against one concrete path.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [switch]$Quiet
    )

    if (-not (Test-Path -LiteralPath $OutFile)) {
        Write-Log "Downloading                  : $Url" -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $Url -OutFile $OutFile -Quiet:$Quiet -ExtraCurlArgs @('-R')
        return
    }

    # The refresh lands in a side file: a 304 writes nothing, a failure removes
    # only the side file, and the cache is replaced only by a complete download.
    $refresh = $OutFile + '.refresh'
    if (Test-Path -LiteralPath $refresh) { Remove-Item -LiteralPath $refresh -Force -ErrorAction SilentlyContinue }
    Write-Log "Cached download present; refreshing only if newer." -Quiet:$Quiet
    try {
        Invoke-DownloadWithRetry -Url $Url -OutFile $refresh -Quiet:$Quiet -ExtraCurlArgs @('-R', '-z', $OutFile)
    }
    catch {
        Write-Log ("Cached download kept; refresh failed: {0}" -f $_.Exception.Message) -Level WARN -Quiet:$Quiet
        return
    }

    if (Test-Path -LiteralPath $refresh) {
        Move-Item -LiteralPath $refresh -Destination $OutFile -Force -ErrorAction Stop
        Write-Log "Cached download replaced by a newer release." -Quiet:$Quiet
    }
    else {
        Write-Log "Cached download is current." -Quiet:$Quiet
    }
}

# ---------------------------------------------------------------------------
# TLS 1.2 enforcement
# ---------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-ProfileContentScope {
    <#
    .SYNOPSIS
        The profile-scoped suffix for network content folders.

    .DESCRIPTION
        Two profiles of one application at one vendor version must never write to the
        same network folder. The scope comes from the run snapshot the caller inherits,
        so every path helper resolves the same value without a packager change. The
        virtual default profile has no suffix and keeps today's paths byte for byte.
    #>
    param([AllowNull()][string]$SnapshotPath = $env:APP_PACKAGER_RUN_SNAPSHOT)

    $empty = [pscustomobject]@{ ProfileId = 'default'; Name = ''; Suffix = '' }
    if ([string]::IsNullOrWhiteSpace($SnapshotPath) -or -not (Test-Path -LiteralPath "FileSystem::$SnapshotPath")) { return $empty }

    try { $snapshot = Get-Content -LiteralPath $SnapshotPath -Raw | ConvertFrom-Json }
    catch { return $empty }

    $profileId = ''
    if ($snapshot.PSObject.Properties['ProfileId']) { $profileId = [string]$snapshot.ProfileId }
    if ([string]::IsNullOrWhiteSpace($profileId) -or $profileId -eq 'default') { return $empty }

    $name = ''
    if ($snapshot.PSObject.Properties['Profile'] -and $snapshot.Profile -and
        $snapshot.Profile.PSObject.Properties['Name']) {
        $name = [string]$snapshot.Profile.Name
    }
    $suffix = ($name -replace '[^A-Za-z0-9._-]', '').TrimEnd([char]'.', [char]' ')
    if ([string]::IsNullOrWhiteSpace($suffix)) {
        $suffix = ($profileId -replace '[^A-Za-z0-9._-]', '')
        if ($suffix.Length -gt 8) { $suffix = $suffix.Substring(0, 8) }
    }
    return [pscustomobject]@{ ProfileId = $profileId; Name = $name; Suffix = $suffix }
}

function Get-ProfileScopedVersionFolder {
    <#
    .SYNOPSIS
        The version folder name for the current profile scope.
    #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [AllowEmptyString()][string]$ProfileSuffix = (Get-ProfileContentScope).Suffix
    )
    if ([string]::IsNullOrWhiteSpace($ProfileSuffix)) { return $Version }
    return ('{0}-{1}' -f $Version, $ProfileSuffix)
}

function ConvertTo-SafePathSegment {
    <#
    .SYNOPSIS
        Turns an identity string into one folder name, or '' when nothing
        usable remains.

    .DESCRIPTION
        Removes path separators, drive and wildcard characters and control
        characters, collapses whitespace, and trims trailing dots and spaces
        (Windows strips them, so the stored name would differ from the
        requested one). A result made only of dots is refused, and a reserved
        device name gets a trailing underscore, so the segment can never climb
        out of or alias the folder it is joined to.
    #>
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ($null -eq $Value) { return '' }
    $segment = (($Value -replace '[\\/:*?"<>|\x00-\x1F]', '') -replace '\s+', ' ').Trim()
    $segment = $segment.TrimEnd([char]'.', [char]' ')
    if ($segment -eq '' -or $segment -match '^\.+$') { return '' }
    if ($segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $segment += '_' }
    return $segment
}

function Assert-SafePathSegment {
    <#
    .SYNOPSIS
        Throws unless a value is one plain folder name.
    #>
    param(
        [AllowNull()][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -match '[\\/:*?"<>|\x00-\x1F]' -or $Value.Trim() -match '^\.+$') {
        throw "$Label '$Value' is not a usable folder name."
    }
}

function Assert-PathUnderRoot {
    <#
    .SYNOPSIS
        Throws unless a path resolves to a location beneath a root folder.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root
    )

    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd([char]'\', [char]'/') + '\'
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path '$fullPath' is outside the root folder '$fullRoot'."
    }
}

function Get-NetworkContentPath {
    <#
    .SYNOPSIS
        Creates and returns the network content folder for one application
        version in the configured share layout.

    .DESCRIPTION
        Nested (default): <FileServerPath>\Applications\<Vendor>\<App>\<Version>
        Flat:             <FileServerPath>\Applications\<Vendor>-<App>-<Version>

        Both layouts live under Applications\. Nested keeps an app's version
        folders adjacent (retention pruning deletes old version folders in
        place); Flat serves org conventions that mandate one folder per
        package. The layout choice must stay consistent across a site: mixing
        them leaves content split across two trees.
    #>
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder,
        [Parameter(Mandatory)][string]$Version,
        [ValidateSet('Nested', 'Flat')][string]$Layout = 'Nested',
        [AllowEmptyString()][string]$ProfileSuffix = (Get-ProfileContentScope).Suffix
    )

    Assert-SafePathSegment -Value $VendorFolder -Label 'Vendor folder'
    Assert-SafePathSegment -Value $AppFolder -Label 'Application folder'
    Assert-SafePathSegment -Value $Version -Label 'Version folder'

    # The version segment carries the profile scope in both layouts, so the stage, the
    # integrity check, the network sync and the deployment type all resolve one path.
    $versionFolder = Get-ProfileScopedVersionFolder -Version $Version -ProfileSuffix $ProfileSuffix

    if ($Layout -eq 'Flat') {
        $appsRoot    = Join-Path $FileServerPath 'Applications'
        $contentPath = Join-Path $appsRoot (('{0}-{1}-{2}' -f $VendorFolder, $AppFolder, $versionFolder).TrimEnd([char]'.', [char]' '))
        Initialize-Folder -Path $appsRoot
        Initialize-Folder -Path $contentPath
        return $contentPath
    }

    $appRoot     = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $contentPath = Join-Path $appRoot $versionFolder
    Initialize-Folder -Path $contentPath
    return $contentPath
}

# ---------------------------------------------------------------------------
# Environment & pre-flight checks
# ---------------------------------------------------------------------------

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Log "Admin check failed: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Get-AppPackagerRootPreferences {
    $prefsPath = Join-Path $PSScriptRoot '..\AppPackager.preferences.json'
    if (-not (Test-Path -LiteralPath $prefsPath)) { return $null }

    try {
        $raw = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        return $null
    }
}


function Resolve-CMProviderMachineName {
    param([string]$ProviderMachineName)

    if (-not [string]::IsNullOrWhiteSpace($ProviderMachineName)) {
        return $ProviderMachineName.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_CM_PROVIDER)) {
        return $env:APP_PACKAGER_CM_PROVIDER.Trim()
    }

    $prefs = Get-AppPackagerRootPreferences
    if ($prefs) {
        foreach ($propName in @('ProviderMachineName','ServerFQDN','ProviderServer')) {
            $prop = $prefs.PSObject.Properties[$propName]
            if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                return ([string]$prop.Value).Trim()
            }
        }
        if ($prefs.MECM) {
            foreach ($propName in @('ProviderMachineName','ServerFQDN','ProviderServer')) {
                $prop = $prefs.MECM.PSObject.Properties[$propName]
                if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                    return ([string]$prop.Value).Trim()
                }
            }
        }
    }

    return $null
}


function Connect-CMSite {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$ProviderMachineName = $null
    )
    # Resolution stays app-side: preferences, the AdminUI connect script,
    # and APP_PACKAGER_CM_PROVIDER feed Resolve-CMProviderMachineName; the
    # drive and session mechanics live in SuiteCommon. Site verification
    # stays off: packager runs never queried the site on connect.
    $provider = Resolve-CMProviderMachineName -ProviderMachineName $ProviderMachineName
    if ([string]::IsNullOrWhiteSpace($provider) -and
        -not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        Write-Log "Configuration Manager PSDrive '$SiteCode' is not available and no provider computer name is configured. Set Site server / SMS Provider in Options > Connections > ConfigMgr, copy the ProviderMachineName value from the AdminUI connect script, or set APP_PACKAGER_CM_PROVIDER." -Level ERROR
        return $false
    }
    return (SuiteCommon\Connect-CMSite -SiteCode $SiteCode -SMSProvider $provider -SkipSiteVerification)
}

function Initialize-Folder {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Log "Network path does not exist or is inaccessible: $Path" -Level ERROR
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Log "Network share is not writable: $Path ($($_.Exception.Message))" -Level ERROR
            return $false
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# ConfigMgr helpers
# ---------------------------------------------------------------------------

function Get-BurnBundleId {
    param([Parameter(Mandatory)][string]$Path)

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        if ($stream.Length -lt 0x40 -or $reader.ReadUInt16() -ne 0x5A4D) { return $null }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        if ($peOffset -lt 0 -or $peOffset + 24 -gt $stream.Length) { return $null }
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { return $null }
        $stream.Position = $peOffset + 6
        $sectionCount = $reader.ReadUInt16()
        $stream.Position = $peOffset + 20
        $optionalHeaderSize = $reader.ReadUInt16()
        $sectionTable = $peOffset + 24 + $optionalHeaderSize
        for ($i = 0; $i -lt $sectionCount; $i++) {
            $entry = $sectionTable + (40 * $i)
            if ($entry + 40 -gt $stream.Length) { return $null }
            $stream.Position = $entry
            $name = [System.Text.Encoding]::ASCII.GetString($reader.ReadBytes(8)).TrimEnd([char]0)
            if ($name -ne '.wixburn') { continue }
            $stream.Position = $entry + 20
            $rawOffset = $reader.ReadUInt32()
            if ($rawOffset + 24 -gt $stream.Length) { return $null }
            $stream.Position = $rawOffset
            if ($reader.ReadUInt32() -ne 0x00F14300) { return $null }
            [void]$reader.ReadUInt32()
            return ('{{{0}}}' -f ([guid]::new($reader.ReadBytes(16))).ToString().ToLowerInvariant())
        }
        return $null
    }
    finally {
        $stream.Dispose()
    }
}
function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)

    $installer = $null
    $db = $null
    $view = $null
    $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}

        foreach ($p in $wanted) {
            $sql  = "SELECT `Value` FROM `Property` WHERE `Property`='$p'"
            $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)

            if ($null -ne $record) {
                $val = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
                $map[$p] = $val
            }
            else {
                $map[$p] = $null
            }
        }

        return $map
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------------
# ARP (Add/Remove Programs) registry discovery
# ---------------------------------------------------------------------------

function Find-UninstallEntry {
    <#
    .SYNOPSIS
        Searches the ARP uninstall registry keys for a product by DisplayName.

    .DESCRIPTION
        Searches both native and WOW6432Node uninstall registry paths for entries
        matching the given DisplayName pattern. Returns the registry key path
        (relative, ready for New-CMDetectionClauseRegistryKeyValue), DisplayVersion,
        Publisher, and uninstall strings.

        Supports retry/polling for installers that register asynchronously.
    #>
    param(
        [Parameter(Mandatory)][string]$DisplayNamePattern,

        [string]$ExpectedVersion,

        [int]$MaxRetries = 1,

        [int]$RetryDelaySec = 0
    )

    $uninstallRoots = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $true },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $false }
    )

    $found = $null

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Registry poll attempt {0}/{1} - sleeping {2}s..." -f $attempt, $MaxRetries, $RetryDelaySec) -Level WARN
            Start-Sleep -Seconds $RetryDelaySec
        }

        $candidates = @()

        foreach ($root in $uninstallRoots) {
            $keys = Get-ChildItem -Path $root.Path -ErrorAction SilentlyContinue
            if (-not $keys) { continue }

            foreach ($k in $keys) {
                $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $dn = $props.DisplayName
                if ([string]::IsNullOrWhiteSpace($dn)) { continue }

                if ($dn -like $DisplayNamePattern) {
                    $regRelative = ($root.Path -replace '^HKLM:\\', '') + '\' + $k.PSChildName

                    $candidates += [pscustomobject]@{
                        RegistryKeyRelative  = $regRelative
                        DisplayName          = $dn
                        DisplayVersion       = $props.DisplayVersion
                        Publisher            = $props.Publisher
                        UninstallString      = $props.UninstallString
                        QuietUninstallString = $props.QuietUninstallString
                        Is64Bit              = $root.Is64Bit
                    }
                }
            }
        }

        if ($candidates.Count -gt 0) {
            if ($ExpectedVersion) {
                $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
                if ($match) { $found = $match; break }
            }

            $found = $candidates | Select-Object -First 1
            break
        }
    }

    return $found
}

# ---------------------------------------------------------------------------
# Stage manifest
# ---------------------------------------------------------------------------

function Get-StageFileHashes {
    <#
    .SYNOPSIS
        Computes SHA256 + size metadata for files under a stage folder.

    .DESCRIPTION
        Returns an ordered list of RelativePath/Sha256/Size records. Relative
        paths use backslashes so manifest data is stable across callers. The
        stage manifest itself is excluded by default because it is the recorder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [string[]]$Exclude = @('stage-manifest.json')
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        throw "Stage root not found: $Root"
    }

    $rootItem = Get-Item -LiteralPath $Root -ErrorAction Stop
    if (-not $rootItem.PSIsContainer) {
        throw "Stage root is not a directory: $Root"
    }

    $rootFull = $rootItem.FullName.TrimEnd('\', '/')
    $excludeSet = @{}
    foreach ($item in @($Exclude)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $normalized = ([string]$item).TrimStart('\', '/') -replace '/', '\'
        $excludeSet[$normalized.ToLowerInvariant()] = $true
    }

    $records = New-Object System.Collections.Generic.List[object]
    $files = @(Get-ChildItem -LiteralPath $rootFull -File -Recurse -Force -ErrorAction Stop | Sort-Object FullName)
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\', '/') -replace '/', '\'
        if ($excludeSet.ContainsKey($relative.ToLowerInvariant())) { continue }

        $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop
        $records.Add([pscustomobject]@{
            RelativePath = $relative
            Sha256       = $hash.Hash.ToUpperInvariant()
            Size         = [Int64]$file.Length
        })
    }

    return @($records.ToArray())
}

function Compare-StageFileHashes {
    <#
    .SYNOPSIS
        Compares a folder tree to expected stage-manifest file hashes.

    .DESCRIPTION
        Returns a result object instead of throwing so callers can decide
        whether to hard-fail or warn. Missing expected hash lists are treated as
        a soft-landing pass for pre-1.0.7 manifests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [AllowNull()][object[]]$Expected,
        [string[]]$Exclude = @('stage-manifest.json'),
        [switch]$AllowExtra
    )

    if ($null -eq $Expected) {
        return [pscustomobject]@{
            Pass          = $true
            Skipped       = $true
            Reason        = 'Stage manifest does not contain FileHashes.'
            Missing       = @()
            Mismatches    = @()
            Extra         = @()
            ExpectedCount = 0
            ActualCount   = 0
            Root          = $Root
        }
    }

    if (-not (Test-Path -LiteralPath $Root)) {
        return [pscustomobject]@{
            Pass          = $false
            Skipped       = $false
            Reason        = "Stage root not found: $Root"
            Missing       = @($Expected)
            Mismatches    = @()
            Extra         = @()
            ExpectedCount = @($Expected).Count
            ActualCount   = 0
            Root          = $Root
        }
    }

    $expectedMap = @{}
    foreach ($entry in @($Expected)) {
        if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.RelativePath)) { continue }
        $relative = ([string]$entry.RelativePath).TrimStart('\', '/') -replace '/', '\'
        $expectedMap[$relative.ToLowerInvariant()] = [pscustomobject]@{
            RelativePath = $relative
            Sha256       = ([string]$entry.Sha256).ToUpperInvariant()
            Size         = [Int64]$entry.Size
        }
    }

    $actual = @(Get-StageFileHashes -Root $Root -Exclude $Exclude)
    $actualMap = @{}
    foreach ($entry in $actual) {
        $actualMap[[string]$entry.RelativePath.ToLowerInvariant()] = $entry
    }

    $missing = New-Object System.Collections.Generic.List[object]
    $mismatches = New-Object System.Collections.Generic.List[object]
    foreach ($key in $expectedMap.Keys) {
        $expectedEntry = $expectedMap[$key]
        if (-not $actualMap.ContainsKey($key)) {
            $missing.Add($expectedEntry)
            continue
        }

        $actualEntry = $actualMap[$key]
        if ([Int64]$actualEntry.Size -ne [Int64]$expectedEntry.Size -or
            ([string]$actualEntry.Sha256).ToUpperInvariant() -ne ([string]$expectedEntry.Sha256).ToUpperInvariant()) {
            $mismatches.Add([pscustomobject]@{
                RelativePath   = $expectedEntry.RelativePath
                ExpectedSha256 = $expectedEntry.Sha256
                ActualSha256   = $actualEntry.Sha256
                ExpectedSize   = [Int64]$expectedEntry.Size
                ActualSize     = [Int64]$actualEntry.Size
            })
        }
    }

    $extra = New-Object System.Collections.Generic.List[object]
    if (-not $AllowExtra) {
        foreach ($key in $actualMap.Keys) {
            if (-not $expectedMap.ContainsKey($key)) {
                $extra.Add($actualMap[$key])
            }
        }
    }

    $missingArray = @($missing.ToArray())
    $mismatchArray = @($mismatches.ToArray())
    $extraArray = @($extra.ToArray())

    return [pscustomobject]@{
        Pass          = ($missingArray.Count -eq 0 -and $mismatchArray.Count -eq 0 -and $extraArray.Count -eq 0)
        Skipped       = $false
        Reason        = ''
        Missing       = $missingArray
        Mismatches    = $mismatchArray
        Extra         = $extraArray
        ExpectedCount = $expectedMap.Count
        ActualCount   = $actual.Count
        Root          = $Root
    }
}

function Sync-StagedContentToNetwork {
    <#
    .SYNOPSIS
        Copies staged content to the network share, refreshing stale files.

    .DESCRIPTION
        Enumerates the local stage recursively and mirrors it under
        NetworkContentPath, recreating intermediate directories. Presence alone
        is not proof of currency: a destination file that already exists is
        compared by SHA-256 and overwritten on mismatch. Re-staging the same
        SoftwareVersion with different wrapper content therefore leaves the
        share matching the stage manifest, which New-MECMApplicationFromManifest
        verifies through Compare-StageFileHashes.

        When the destination already holds files and this stage would add or
        change any of them, and the existing-application policy is not
        Overwrite, nothing is written: the copy is recorded and
        New-MECMApplicationFromManifest runs it only when no application
        exists or the application is replaced. -NoDefer copies regardless.

        When Manifest carries FileHashes, the expected hash for a relative path
        comes from that record; otherwise the local file is hashed. The stage
        manifest itself is never copied.

    .OUTPUTS
        Nothing unless -PassThru is supplied, in which case a summary object
        with Copied/Refreshed/Unchanged counts and the relative paths touched.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalContentPath,
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [AllowNull()][object]$Manifest,
        [string[]]$Exclude = @('stage-manifest.json'),
        [switch]$PassThru,
        [switch]$NoDefer
    )

    $localRoot = (Resolve-Path -LiteralPath $LocalContentPath -ErrorAction Stop).Path.TrimEnd('\', '/')

    $excludeSet = @{}
    foreach ($item in @($Exclude)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $normalized = ([string]$item).TrimStart('\', '/') -replace '/', '\'
        $excludeSet[$normalized.ToLowerInvariant()] = $true
    }

    $expectedMap = @{}
    if ($null -ne $Manifest -and $Manifest.PSObject.Properties.Name -contains 'FileHashes' -and $null -ne $Manifest.FileHashes) {
        foreach ($entry in @($Manifest.FileHashes)) {
            if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.RelativePath)) { continue }
            $relative = ([string]$entry.RelativePath).TrimStart('\', '/') -replace '/', '\'
            $expectedMap[$relative.ToLowerInvariant()] = ([string]$entry.Sha256).ToUpperInvariant()
        }
    }

    # A stage that was refused after the packager wrote files leaves a folder
    # whose files no longer match the manifest beside them. That folder is
    # refused here, before any of it reaches the share.
    if ($expectedMap.Count -gt 0) {
        $localCheck = Compare-StageFileHashes -Root $localRoot -Expected @($Manifest.FileHashes) -Exclude $Exclude
        if (-not $localCheck.Skipped -and -not $localCheck.Pass) {
            throw ("Staged content in '{0}' does not match its stage manifest: {1}. Re-run the Stage phase; nothing was copied." -f $localRoot, (Format-StageFileHashComparison -Comparison $localCheck))
        }
    }

    $script:DeferredContentSync = $null
    $destinationHadFiles = (Test-Path -LiteralPath $NetworkContentPath) -and
        $null -ne (Get-ChildItem -LiteralPath $NetworkContentPath -File -Recurse -Force -ErrorAction Stop | Select-Object -First 1)

    $plan = New-Object System.Collections.Generic.List[object]
    $localFiles = @(Get-ChildItem -LiteralPath $localRoot -File -Recurse -Force -ErrorAction Stop | Sort-Object FullName)
    foreach ($f in $localFiles) {
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\', '/') -replace '/', '\'
        if ($excludeSet.ContainsKey($relative.ToLowerInvariant())) { continue }

        $dest = Join-Path $NetworkContentPath $relative
        $action = 'Copy'
        if (Test-Path -LiteralPath $dest) {
            $expected = $expectedMap[$relative.ToLowerInvariant()]
            if ([string]::IsNullOrWhiteSpace($expected)) {
                $expected = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
            }
            $actual = (Get-FileHash -LiteralPath $dest -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
            $action = if ($actual -eq $expected) { 'Unchanged' } else { 'Refresh' }
        }
        $plan.Add([pscustomobject]@{ Source = $f.FullName; Relative = $relative; Destination = $dest; Action = $action })
    }

    # Content already on the share may belong to a published application.
    # Changing it before New-MECMApplicationFromManifest decides on that
    # application rewrites the source of an application that Skip or Fail
    # then leaves in place, and the next redistribution ships the new files.
    $changes = @($plan | Where-Object { $_.Action -ne 'Unchanged' })
    if (-not $NoDefer -and $destinationHadFiles -and $changes.Count -gt 0 -and
        (Resolve-OnExistingBehavior -Requested '').Behavior -ne 'Overwrite') {
        $script:DeferredContentSync = [pscustomobject]@{
            LocalContentPath   = $LocalContentPath
            NetworkContentPath = $NetworkContentPath
            Manifest           = $Manifest
            Exclude            = $Exclude
        }
        Write-Log ("Network content differs      : {0} file(s); the copy waits for the existing-application decision" -f $changes.Count) -Level WARN
        if ($PassThru) {
            return [pscustomobject]@{
                LocalContentPath   = $localRoot
                NetworkContentPath = $NetworkContentPath
                Deferred           = $true
                Copied             = @()
                Refreshed          = @()
                Unchanged          = @()
                CopiedCount        = 0
                RefreshedCount     = 0
                UnchangedCount     = 0
            }
        }
        return
    }

    # A file on the share that the manifest does not name fails the package
    # integrity verification that follows this copy. The copy never removes
    # files, so the run is refused here, before any file on the share changes.
    if ($destinationHadFiles -and $expectedMap.Count -gt 0) {
        $extra = @((Compare-StageFileHashes -Root $NetworkContentPath -Expected @($Manifest.FileHashes)).Extra)
        if ($extra.Count -gt 0) {
            throw ("Network content folder '{0}' holds {1} file(s) that the stage manifest does not name: {2}. Remove them or publish to an empty folder; nothing was copied." -f $NetworkContentPath, $extra.Count, (($extra | Select-Object -First 5 | ForEach-Object { [string]$_.RelativePath }) -join ', '))
        }
    }

    if (-not (Test-Path -LiteralPath $NetworkContentPath)) {
        New-Item -ItemType Directory -Path $NetworkContentPath -Force -ErrorAction Stop | Out-Null
    }

    $copied = New-Object System.Collections.Generic.List[string]
    $refreshed = New-Object System.Collections.Generic.List[string]
    $unchanged = New-Object System.Collections.Generic.List[string]

    foreach ($entry in $plan) {
        if ($entry.Action -eq 'Unchanged') {
            Write-Log "Already on network           : $($entry.Relative)"
            $unchanged.Add($entry.Relative)
            continue
        }

        $destDir = Split-Path -Parent $entry.Destination
        if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null
        }
        Copy-Item -LiteralPath $entry.Source -Destination $entry.Destination -Force -ErrorAction Stop
        if ($entry.Action -eq 'Copy') {
            Write-Log "Copied to network            : $($entry.Relative)"
            $copied.Add($entry.Relative)
        }
        else {
            Write-Log "Refreshed on network         : $($entry.Relative)"
            $refreshed.Add($entry.Relative)
        }
    }

    if ($PassThru) {
        return [pscustomobject]@{
            LocalContentPath   = $localRoot
            NetworkContentPath = $NetworkContentPath
            Deferred           = $false
            Copied             = @($copied.ToArray())
            Refreshed          = @($refreshed.ToArray())
            Unchanged          = @($unchanged.ToArray())
            CopiedCount        = $copied.Count
            RefreshedCount     = $refreshed.Count
            UnchangedCount     = $unchanged.Count
        }
    }
}

function Format-StageFileHashComparison {
    param([Parameter(Mandatory)]$Comparison)

    if ($Comparison.Pass) { return 'Stage/package file integrity verified.' }
    if ($Comparison.Skipped) { return [string]$Comparison.Reason }

    $parts = New-Object System.Collections.Generic.List[string]
    if ($Comparison.Missing.Count -gt 0) {
        $sample = @($Comparison.Missing | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("missing {0}: {1}" -f $Comparison.Missing.Count, $sample))
    }
    if ($Comparison.Mismatches.Count -gt 0) {
        $sample = @($Comparison.Mismatches | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("mismatched {0}: {1}" -f $Comparison.Mismatches.Count, $sample))
    }
    if ($Comparison.Extra.Count -gt 0) {
        $sample = @($Comparison.Extra | Select-Object -First 5 | ForEach-Object { $_.RelativePath }) -join ', '
        $parts.Add(("extra {0}: {1}" -f $Comparison.Extra.Count, $sample))
    }
    if ($parts.Count -eq 0 -and $Comparison.Reason) { $parts.Add([string]$Comparison.Reason) }
    return ($parts.ToArray() -join '; ')
}

# ---------------------------------------------------------------------------
# Installer icon extraction
# ---------------------------------------------------------------------------

# ConfigMgr's Set-CMApplication -IconLocationFile and Intune's win32LobApp
# largeIcon both accept .ico and .png; ConfigMgr stores at most 512x512, so a
# larger extraction is downscaled before it is handed over.
$script:IconMaxDimension = 512

# A 16- or 32-pixel-only icon resource is what generic installer stubs ship;
# real application icons carry a 48px or larger image. Distinguishing an
# NSIS/Inno default setup icon from a real one by content is not reliable,
# so the size floor plus the per-packager IconSource opt-in is the control.
$script:IconMinimumDimension = 32

function Initialize-IconNativeType {
    <#
    .SYNOPSIS
        Compiles the icon-resource P/Invoke helper once per process.

    .DESCRIPTION
        Add-Type registers the type in the AppDomain permanently; a second
        call in the same process throws on the duplicate name, so the type
        presence is the guard rather than a script-scope flag (a -Force
        module reimport resets script scope but not the AppDomain).
    #>
    if ('AppPackager.IconNative' -as [type]) { return }

    $memberDefinition = @'
    private const int RT_GROUP_ICON = 14;
    private const int RT_ICON = 3;
    private const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryExW(string lpFileName, IntPtr hFile, uint dwFlags);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FreeLibrary(IntPtr hModule);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr FindResourceW(IntPtr hModule, IntPtr lpName, IntPtr lpType);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LockResource(IntPtr hResData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);

    public delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool EnumResourceNamesW(IntPtr hModule, IntPtr lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);

    private static byte[] ResourceBytes(IntPtr hModule, IntPtr type, IntPtr name)
    {
        IntPtr info = FindResourceW(hModule, name, type);
        if (info == IntPtr.Zero) { return null; }
        uint size = SizeofResource(hModule, info);
        if (size == 0) { return null; }
        IntPtr data = LoadResource(hModule, info);
        if (data == IntPtr.Zero) { return null; }
        IntPtr ptr = LockResource(data);
        if (ptr == IntPtr.Zero) { return null; }
        byte[] buffer = new byte[size];
        Marshal.Copy(ptr, buffer, 0, (int)size);
        return buffer;
    }

    // Returns a single-image .ico file built from the largest RT_ICON in the
    // module's first icon group. Returns null when the file carries no icon
    // group. The reported pixel size is the true resource size: the group
    // directory is read directly rather than asking the shell for a scaled
    // handle, which would report whatever size was requested.
    public static byte[] ExtractLargestIcon(string path, out int pixelSize)
    {
        pixelSize = 0;
        IntPtr hModule = LoadLibraryExW(path, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == IntPtr.Zero) { return null; }
        try
        {
            IntPtr firstGroup = IntPtr.Zero;
            EnumResNameProc collect = delegate(IntPtr m, IntPtr t, IntPtr n, IntPtr l)
            {
                firstGroup = n;
                return false;
            };
            EnumResourceNamesW(hModule, (IntPtr)RT_GROUP_ICON, collect, IntPtr.Zero);
            if (firstGroup == IntPtr.Zero) { return null; }

            byte[] group = ResourceBytes(hModule, (IntPtr)RT_GROUP_ICON, firstGroup);
            if (group == null || group.Length < 6) { return null; }

            int count = BitConverter.ToUInt16(group, 4);
            if (count <= 0 || group.Length < 6 + (count * 14)) { return null; }

            int bestIndex = -1;
            int bestSize = -1;
            int bestBits = -1;
            for (int i = 0; i < count; i++)
            {
                int off = 6 + (i * 14);
                int w = group[off];
                int h = group[off + 1];
                // A zero width/height byte encodes 256 in the icon directory.
                if (w == 0) { w = 256; }
                if (h == 0) { h = 256; }
                int dim = Math.Max(w, h);
                int bits = BitConverter.ToUInt16(group, off + 6);
                if (dim > bestSize || (dim == bestSize && bits > bestBits))
                {
                    bestSize = dim;
                    bestBits = bits;
                    bestIndex = i;
                }
            }
            if (bestIndex < 0) { return null; }

            int bestOff = 6 + (bestIndex * 14);
            int iconId = BitConverter.ToUInt16(group, bestOff + 12);
            byte[] image = ResourceBytes(hModule, (IntPtr)RT_ICON, (IntPtr)iconId);
            if (image == null || image.Length == 0) { return null; }

            byte[] ico = new byte[22 + image.Length];
            ico[2] = 1;
            ico[4] = 1;
            ico[6] = group[bestOff];
            ico[7] = group[bestOff + 1];
            ico[8] = group[bestOff + 2];
            ico[10] = group[bestOff + 4];
            ico[11] = group[bestOff + 5];
            ico[12] = group[bestOff + 6];
            ico[13] = group[bestOff + 7];
            Buffer.BlockCopy(BitConverter.GetBytes(image.Length), 0, ico, 14, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(22), 0, ico, 18, 4);
            Buffer.BlockCopy(image, 0, ico, 22, image.Length);

            pixelSize = bestSize;
            return ico;
        }
        finally
        {
            FreeLibrary(hModule);
        }
    }
'@

    Add-Type -MemberDefinition $memberDefinition `
        -Name 'IconNative' `
        -Namespace 'AppPackager' `
        -ErrorAction Stop
}

function ConvertTo-SingleLargestIconFrame {
    <#
    .SYNOPSIS
        Rebuilds .ico bytes as a single-frame icon holding the largest frame.

    .DESCRIPTION
        .ico directory entries are 16 bytes; a zero width or height byte
        encodes 256. Input that is not a parseable icon directory is returned
        unchanged.
    #>
    param([Parameter(Mandatory)][byte[]]$Bytes)

    if ($Bytes.Length -lt 22) { return $Bytes }
    if ([BitConverter]::ToUInt16($Bytes, 0) -ne 0 -or [BitConverter]::ToUInt16($Bytes, 2) -ne 1) { return $Bytes }
    $count = [BitConverter]::ToUInt16($Bytes, 4)
    if ($count -le 1 -or $Bytes.Length -lt (6 + $count * 16)) { return $Bytes }

    $bestIndex = -1; $bestDim = -1; $bestBits = -1
    for ($i = 0; $i -lt $count; $i++) {
        $off = 6 + ($i * 16)
        $w = [int]$Bytes[$off]; if ($w -eq 0) { $w = 256 }
        $h = [int]$Bytes[$off + 1]; if ($h -eq 0) { $h = 256 }
        $dim = [Math]::Max($w, $h)
        $bits = [BitConverter]::ToUInt16($Bytes, $off + 6)
        if ($dim -gt $bestDim -or ($dim -eq $bestDim -and $bits -gt $bestBits)) {
            $bestDim = $dim; $bestBits = $bits; $bestIndex = $i
        }
    }
    if ($bestIndex -lt 0) { return $Bytes }

    $entryOff = 6 + ($bestIndex * 16)
    $imageSize = [BitConverter]::ToInt32($Bytes, $entryOff + 8)
    $imageOff = [BitConverter]::ToInt32($Bytes, $entryOff + 12)
    if ($imageSize -le 0 -or $imageOff -le 0 -or ($imageOff + $imageSize) -gt $Bytes.Length) { return $Bytes }

    $out = New-Object byte[] (22 + $imageSize)
    $out[2] = 1; $out[4] = 1
    [Array]::Copy($Bytes, $entryOff, $out, 6, 8)
    [Buffer]::BlockCopy([BitConverter]::GetBytes($imageSize), 0, $out, 14, 4)
    [Buffer]::BlockCopy([BitConverter]::GetBytes(22), 0, $out, 18, 4)
    [Array]::Copy($Bytes, $imageOff, $out, 22, $imageSize)
    return $out
}

function Get-IconBytesDimension {
    <#
    .SYNOPSIS
        Reads the largest image dimension out of .ico file bytes.

    .DESCRIPTION
        Returns 0 when the buffer is not an icon directory. A zero width or
        height byte in an ICONDIRENTRY encodes 256 pixels.
    #>
    param([Parameter(Mandatory)][AllowNull()][byte[]]$Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -lt 22) { return 0 }
    if ([BitConverter]::ToUInt16($Bytes, 0) -ne 0 -or [BitConverter]::ToUInt16($Bytes, 2) -ne 1) { return 0 }

    $count = [BitConverter]::ToUInt16($Bytes, 4)
    if ($count -le 0 -or $Bytes.Length -lt (6 + ($count * 16))) { return 0 }

    $best = 0
    for ($i = 0; $i -lt $count; $i++) {
        $off = 6 + ($i * 16)
        $w = [int]$Bytes[$off]; if ($w -eq 0) { $w = 256 }
        $h = [int]$Bytes[$off + 1]; if ($h -eq 0) { $h = 256 }
        $dim = [Math]::Max($w, $h)
        if ($dim -gt $best) { $best = $dim }
    }
    return $best
}

function Get-MsiIconBytes {
    <#
    .SYNOPSIS
        Reads an icon stream out of an MSI Icon table.

    .DESCRIPTION
        Prefers the icon named by the ARP ProductIcon property, falling back
        to the first row of the Icon table. Returns $null when the database
        has no Icon table or no rows. The COM objects are released explicitly
        because a live WindowsInstaller.Database handle keeps a write lock on
        the file for the rest of the process lifetime.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $installer = $null
    $database = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer -ErrorAction Stop
        $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))

        $wanted = ''
        try {
            $propView = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database,
                @("SELECT Value FROM Property WHERE Property = 'ARPPRODUCTICON'"))
            $null = $propView.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $propView, $null)
            $propRec = $propView.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $propView, $null)
            if ($propRec) {
                $wanted = [string]$propRec.GetType().InvokeMember('StringData', 'GetProperty', $null, $propRec, @(1))
            }
            $null = $propView.GetType().InvokeMember('Close', 'InvokeMethod', $null, $propView, $null)
        }
        catch {
            # An MSI without a Property table row for the ARP icon is normal;
            # the Icon table fallback still applies.
            $wanted = ''
        }

        $query = if ([string]::IsNullOrWhiteSpace($wanted)) {
            'SELECT Name, Data FROM Icon'
        }
        else {
            "SELECT Name, Data FROM Icon WHERE Name = '{0}'" -f ($wanted -replace "'", "''")
        }

        $view = $null
        try {
            $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @($query))
            $null = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
            $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        }
        catch {
            return $null
        }

        if (-not $record -and -not [string]::IsNullOrWhiteSpace($wanted)) {
            # A ProductIcon name that does not resolve (authoring drift) still
            # leaves the table's own rows usable.
            $null = $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
            $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @('SELECT Name, Data FROM Icon'))
            $null = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
            $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        }
        if (-not $record) { return $null }

        $size = [int]$record.GetType().InvokeMember('DataSize', 'GetProperty', $null, $record, @(2))
        if ($size -le 0) { return $null }
        $bytes = $record.GetType().InvokeMember('ReadStream', 'InvokeMethod', $null, $record, @(2, $size, 1))
        $null = $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)

        if ($bytes -is [string]) {
            # MsiReadStreamInteger (1) hands back a binary string; each char is
            # one byte of the stream.
            $buffer = New-Object byte[] $bytes.Length
            for ($i = 0; $i -lt $bytes.Length; $i++) { $buffer[$i] = [byte][char]$bytes[$i] }
            return $buffer
        }
        return [byte[]]$bytes
    }
    catch {
        Write-Log ("MSI icon read failed         : {0}" -f $_.Exception.Message) -Level DEBUG
        return $null
    }
    finally {
        foreach ($comObject in @($database, $installer)) {
            if ($comObject) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch { <# Releasing or disposing an object is cleanup. #> $null = $_ }
            }
        }
    }
}

function Get-InstallerIcon {
    <#
    .SYNOPSIS
        Extracts an application icon from an installer into a ConfigMgr-compatible file.

    .DESCRIPTION
        PE inputs (.exe, .dll) yield the largest image of the first icon group,
        read from the resource directory so the reported size is the real one.
        ExtractAssociatedIcon is the fallback and caps at 32 pixels. MSI inputs
        are read from the Icon table, preferring the ARPPRODUCTICON stream.

        OutputPath's extension selects the written format: .png converts through
        System.Drawing, anything else writes the icon bytes verbatim. Images
        wider than 512 pixels are downscaled, the ConfigMgr ceiling.

        Returns $null when no icon exists or the largest image is under
        MinimumSize.

    .OUTPUTS
        [pscustomobject] with Path, PixelSize, Format, and Source, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$OutputPath,
        [int]$MinimumSize = $script:IconMinimumDimension
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Installer not found: $Path"
    }

    $full = (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
    $extension = [System.IO.Path]::GetExtension($full).ToLowerInvariant()

    $bytes = $null
    $pixelSize = 0
    $source = ''

    if ($extension -eq '.msi') {
        $bytes = Get-MsiIconBytes -Path $full
        if ($null -ne $bytes) {
            $pixelSize = Get-IconBytesDimension -Bytes $bytes
            $source = 'MsiIconTable'
            if ($pixelSize -gt 0) {
                # An Icon table stream is a multi-frame .ico file whose first
                # directory entry is often the smallest frame; downstream
                # conversion reads entry 0, so the largest frame is promoted
                # to a single-frame icon here.
                $bytes = ConvertTo-SingleLargestIconFrame -Bytes $bytes
                $pixelSize = Get-IconBytesDimension -Bytes $bytes
            }
            if ($pixelSize -eq 0) {
                # An Icon table row may hold an EXE rather than an .ico; the
                # PE reader handles that shape.
                $temp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([guid]::NewGuid().ToString('n') + '.exe'))
                try {
                    [System.IO.File]::WriteAllBytes($temp, $bytes)
                    Initialize-IconNativeType
                    # Same transient-load retry as the direct PE path; the temp
                    # file was written a moment ago.
                    $peSize = 0
                    $peBytes = $null
                    foreach ($delayMs in 0, 400, 1200) {
                        if ($delayMs) { Start-Sleep -Milliseconds $delayMs }
                        $peBytes = [AppPackager.IconNative]::ExtractLargestIcon($temp, [ref]$peSize)
                        if ($peBytes) { break }
                    }
                    if ($peBytes) { $bytes = $peBytes; $pixelSize = $peSize; $source = 'MsiIconTablePe' }
                }
                finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
            }
        }
    }
    else {
        Initialize-IconNativeType
        # LoadLibraryExW fails transiently on a freshly written file (an
        # on-access scan still holds it) and the helper returns null with no
        # error, which would silently degrade to the 32px fallback on the
        # very first probe of a just-downloaded payload.
        $nativeSize = 0
        $bytes = $null
        foreach ($delayMs in 0, 400, 1200) {
            if ($delayMs) { Start-Sleep -Milliseconds $delayMs }
            $bytes = [AppPackager.IconNative]::ExtractLargestIcon($full, [ref]$nativeSize)
            if ($bytes) { break }
        }
        if ($bytes) {
            $pixelSize = $nativeSize
            $source = 'PeResource'
        }
        else {
            try {
                Add-Type -AssemblyName System.Drawing -ErrorAction Stop
                $associated = [System.Drawing.Icon]::ExtractAssociatedIcon($full)
                if ($associated) {
                    $stream = New-Object System.IO.MemoryStream
                    try {
                        $associated.Save($stream)
                        $bytes = $stream.ToArray()
                    }
                    finally { $stream.Dispose(); $associated.Dispose() }
                    $pixelSize = Get-IconBytesDimension -Bytes $bytes
                    $source = 'ExtractAssociatedIcon'
                }
            }
            catch {
                Write-Log ("Icon fallback failed         : {0}" -f $_.Exception.Message) -Level DEBUG
            }
        }
    }

    if ($null -eq $bytes -or $bytes.Length -eq 0) {
        Write-Log ("No icon resource found       : {0}" -f (Split-Path -Leaf $full)) -Level DEBUG
        return $null
    }

    if ($pixelSize -lt $MinimumSize) {
        Write-Log ("Icon rejected (too small)    : {0}px from {1} (floor {2}px)" -f $pixelSize, (Split-Path -Leaf $full), $MinimumSize) -Level WARN
        return $null
    }

    $outExtension = [System.IO.Path]::GetExtension($OutputPath).ToLowerInvariant()
    $outDirectory = Split-Path -Path $OutputPath -Parent
    if ($outDirectory -and -not (Test-Path -LiteralPath $outDirectory)) {
        New-Item -Path $outDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    if ($outExtension -eq '.png' -or $pixelSize -gt $script:IconMaxDimension) {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $target = [Math]::Min($pixelSize, $script:IconMaxDimension)
        $pngPath = if ($outExtension -eq '.png') { $OutputPath } else { [System.IO.Path]::ChangeExtension($OutputPath, '.png') }

        # A 256px icon frame is normally stored PNG-compressed rather than as a
        # DIB, and System.Drawing.Icon cannot decode that shape; the frame is
        # already the wanted file, so it is copied out byte for byte.
        $imageOffset = if ($bytes.Length -ge 22) { [BitConverter]::ToInt32($bytes, 18) } else { 0 }
        $isPngFrame = ($imageOffset -gt 0 -and $bytes.Length -gt ($imageOffset + 8) -and
            $bytes[$imageOffset] -eq 0x89 -and $bytes[$imageOffset + 1] -eq 0x50 -and
            $bytes[$imageOffset + 2] -eq 0x4E -and $bytes[$imageOffset + 3] -eq 0x47)

        $sourceBitmap = $null
        $bitmap = $null
        $icon = $null
        $stream = $null
        try {
            if ($isPngFrame) {
                $frame = New-Object byte[] ($bytes.Length - $imageOffset)
                [Array]::Copy($bytes, $imageOffset, $frame, 0, $frame.Length)
                if ($pixelSize -le $script:IconMaxDimension) {
                    [System.IO.File]::WriteAllBytes($pngPath, $frame)
                    $OutputPath = $pngPath
                    $pixelSize = $target
                    $frame = $null
                }
                else {
                    $stream = New-Object System.IO.MemoryStream (, $frame)
                    $sourceBitmap = [System.Drawing.Image]::FromStream($stream)
                }
            }
            else {
                $stream = New-Object System.IO.MemoryStream (, $bytes)
                $icon = New-Object System.Drawing.Icon ($stream, $target, $target)
                $sourceBitmap = $icon.ToBitmap()
            }

            if ($sourceBitmap) {
                $bitmap = New-Object System.Drawing.Bitmap ($sourceBitmap, $target, $target)
                $bitmap.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $OutputPath = $pngPath
                $pixelSize = $target
            }
        }
        finally {
            foreach ($disposable in @($bitmap, $sourceBitmap, $icon, $stream)) {
                if ($disposable) { try { $disposable.Dispose() } catch { <# Releasing or disposing an object is cleanup. #> $null = $_ } }
            }
        }
    }
    else {
        [System.IO.File]::WriteAllBytes($OutputPath, $bytes)
    }

    # Configuration Manager caps icon images at 256 KB. An .ico carrying every
    # frame can exceed that even when its largest frame is small, so an
    # oversized file is re-encoded as a single PNG, downscaling until it fits.
    $maxIconBytes = 256KB
    if ((Get-Item -LiteralPath $OutputPath).Length -gt $maxIconBytes) {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $pngPath = [System.IO.Path]::ChangeExtension($OutputPath, '.png')
        $fileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
        # Same PNG-compressed-frame trap as above: System.Drawing.Icon throws
        # on such frames, so they decode through Image.FromStream instead.
        $frameOffset = if ($fileBytes.Length -ge 22 -and $fileBytes[0] -eq 0 -and $fileBytes[2] -eq 1) { [BitConverter]::ToInt32($fileBytes, 18) } else { 0 }
        $hasPngFrame = ($frameOffset -gt 0 -and $fileBytes.Length -gt ($frameOffset + 8) -and
            $fileBytes[$frameOffset] -eq 0x89 -and $fileBytes[$frameOffset + 1] -eq 0x50)
        foreach ($target in @($pixelSize, 256, 128, 64) | Where-Object { $_ -le $pixelSize } | Sort-Object -Descending -Unique) {
            $stream = $null; $icon = $null; $sourceBitmap = $null; $bitmap = $null
            try {
                if ($hasPngFrame) {
                    $frame = New-Object byte[] ($fileBytes.Length - $frameOffset)
                    [Array]::Copy($fileBytes, $frameOffset, $frame, 0, $frame.Length)
                    $stream = New-Object System.IO.MemoryStream (, $frame)
                    $sourceBitmap = [System.Drawing.Image]::FromStream($stream)
                }
                elseif ($fileBytes[0] -eq 0x89 -and $fileBytes[1] -eq 0x50) {
                    $stream = New-Object System.IO.MemoryStream (, $fileBytes)
                    $sourceBitmap = [System.Drawing.Image]::FromStream($stream)
                }
                else {
                    $stream = New-Object System.IO.MemoryStream (, $fileBytes)
                    $icon = New-Object System.Drawing.Icon ($stream, $target, $target)
                    $sourceBitmap = $icon.ToBitmap()
                }
                $bitmap = New-Object System.Drawing.Bitmap ($sourceBitmap, $target, $target)
                $bitmap.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            }
            catch { continue }
            finally {
                foreach ($disposable in @($bitmap, $sourceBitmap, $icon, $stream)) {
                    if ($disposable) { try { $disposable.Dispose() } catch { <# Releasing or disposing an object is cleanup. #> $null = $_ } }
                }
            }
            if ((Get-Item -LiteralPath $pngPath).Length -le $maxIconBytes) {
                if ($pngPath -ne $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue }
                $OutputPath = $pngPath
                $pixelSize = $target
                break
            }
        }
        if ((Get-Item -LiteralPath $OutputPath).Length -gt $maxIconBytes) {
            Write-Log ("Icon rejected (too large)    : {0} exceeds the 256 KB image limit after re-encoding" -f (Split-Path -Leaf $OutputPath)) -Level WARN
            Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            return $null
        }
    }

    Write-Log ("Extracted installer icon     : {0} ({1}px, {2})" -f (Split-Path -Leaf $OutputPath), $pixelSize, $source)

    return [pscustomobject]@{
        Path      = $OutputPath
        PixelSize = $pixelSize
        Format    = [System.IO.Path]::GetExtension($OutputPath).TrimStart('.').ToLowerInvariant()
        Source    = $source
    }
}

function Get-PackagerIconSource {
    <#
    .SYNOPSIS
        Reads the IconSource header tag out of a packager script.

    .DESCRIPTION
        Returns 'Installer', 'External', or 'None'. An absent, unreadable, or
        unrecognized tag is 'None', which keeps every untagged packager on the
        pre-icon path.
    #>
    param([AllowNull()][string]$ScriptPath)

    if ([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        return 'None'
    }

    try {
        $head = Get-Content -LiteralPath $ScriptPath -TotalCount 40 -ErrorAction Stop
    }
    catch {
        return 'None'
    }

    foreach ($line in $head) {
        if ($line -match '^\s*(?:#\s*)?IconSource\s*:\s*(.+?)\s*$') {
            $value = $Matches[1]
            foreach ($known in @('Installer', 'External', 'None')) {
                if ($value -eq $known) { return $known }
            }
            return 'None'
        }
    }
    return 'None'
}

function Add-StageIcon {
    <#
    .SYNOPSIS
        Places app-icon.* in a stage folder, icon pack first, then per the
        packager's IconSource tag.

    .DESCRIPTION
        An icon pack entry Packagers\Icons\<packagername>.png|ico wins for
        every tag: pack icons are at least 256x256, extracted installer icons
        are often 64x64 or 128x128. Without a pack entry, Installer extracts
        from the staged installer named by the manifest's InstallerFile,
        External warns, and None is a no-op.

        Sets ManifestData['Icon'] to the icon's file name on success so the
        Package phase can find it inside the content folder. Never throws: an
        icon is decoration, and a failed extraction must not fail a stage.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [AllowNull()][string]$PackagerScriptPath,
        [string]$IconsDirectory = (Join-Path $PSScriptRoot 'Icons')
    )

    $iconSource = Get-PackagerIconSource -ScriptPath $PackagerScriptPath

    try {
        if (-not [string]::IsNullOrWhiteSpace($PackagerScriptPath)) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($PackagerScriptPath) -replace '^package-', ''
            $candidate = @(Get-ChildItem -LiteralPath $IconsDirectory -Filter "$baseName.*" -File -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -eq $baseName -and $_.Extension -match '^\.(ico|png)$' } | Sort-Object Extension | Select-Object -First 1)
            if ($candidate.Count -gt 0) {
                $destination = Join-Path $StageRoot ('app-icon' + $candidate[0].Extension.ToLowerInvariant())
                Copy-Item -LiteralPath $candidate[0].FullName -Destination $destination -Force -ErrorAction Stop
                $ManifestData['Icon'] = Split-Path -Leaf $destination
                Write-Log ("Staged icon pack icon        : {0}" -f $ManifestData['Icon'])
                return
            }
        }

        if ($iconSource -eq 'None') { return }
        if ($iconSource -eq 'External') {
            Write-Log ("IconSource External but no Packagers\Icons\{0}.ico|png found; continuing without an icon." -f $baseName) -Level WARN
            return
        }

        $installerFile = [string]$ManifestData['InstallerFile']
        if ([string]::IsNullOrWhiteSpace($installerFile)) {
            Write-Log 'IconSource Installer but the manifest declares no InstallerFile; continuing without an icon.' -Level WARN
            return
        }

        $installerPath = Join-Path $StageRoot $installerFile
        if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
            Write-Log ("IconSource Installer but {0} is not in the stage folder; continuing without an icon." -f $installerFile) -Level WARN
            return
        }

        $result = Get-InstallerIcon -Path $installerPath -OutputPath (Join-Path $StageRoot 'app-icon.ico')
        if ($null -eq $result) { return }

        $ManifestData['Icon'] = Split-Path -Leaf $result.Path
    }
    catch {
        Write-Log ("Icon staging failed          : {0}" -f $_.Exception.Message) -Level WARN
    }
}

function Set-CMApplicationIconFromManifest {
    <#
    .SYNOPSIS
        Applies a manifest's icon to an existing ConfigMgr application.

    .DESCRIPTION
        Warns and returns on any failure. A missing or rejected icon is not a
        packaging error, and Set-CMApplication reports success without the
        provider having accepted the write, so the failure surface here is
        advisory either way.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$NetworkContentPath
    )

    $iconName = if ($Manifest.PSObject.Properties.Name -contains 'Icon') { [string]$Manifest.Icon } else { '' }
    if ([string]::IsNullOrWhiteSpace($iconName)) { return }

    $iconPath = Join-Path $NetworkContentPath $iconName
    # This runs while the session sits on the CMSite provider drive, where a
    # bare UNC Test-Path resolves through the CM provider and fails.
    if (-not (Test-Path -LiteralPath ("FileSystem::{0}" -f $iconPath) -PathType Leaf)) {
        Write-Log ("Application icon not found   : {0}" -f $iconPath) -Level WARN
        return
    }

    try {
        Set-CMApplication -Name $AppName -IconLocationFile $iconPath -ErrorAction Stop
        Write-Log ("Application icon set         : {0}" -f $iconName)
    }
    catch {
        Write-Log ("Application icon not applied : {0}" -f $_.Exception.Message) -Level WARN
    }
}

function Get-IconMimeContent {
    <#
    .SYNOPSIS
        Builds a Graph mimeContent object for an icon file.

    .DESCRIPTION
        win32LobApp.largeIcon is a mimeContent: a content mime type plus the
        raw bytes, which Graph JSON carries base64-encoded. Returns $null when
        the file is missing or unreadable.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
    }
    catch {
        Write-Log ("Icon read failed             : {0}" -f $_.Exception.Message) -Level WARN
        return $null
    }
    if ($bytes.Length -eq 0) { return $null }

    $type = switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.png'  { 'image/png' }
        '.jpg'  { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        default { 'image/x-icon' }
    }

    return @{
        '@odata.type' = '#microsoft.graph.mimeContent'
        type          = $type
        value         = [Convert]::ToBase64String($bytes)
    }
}

function Write-StageManifest {
    <#
    .SYNOPSIS
        Writes a stage-manifest.json file.

    .DESCRIPTION
        Serializes ManifestData to JSON with schema metadata.

        Schema v2 adds optional fields for PSADT/deployment tool integration:
          InstallerType     "MSI" or "EXE"
          InstallArgs       Silent install arguments
          UninstallArgs     Silent uninstall arguments
          UninstallCommand  Full uninstall command (for EXE products)
          ProductCode       MSI ProductCode GUID (for MSI products)
          RunningProcess    Array of process names to close before install

        Schema v3 adds FileHashes, an ordered list of every staged payload and
        wrapper file with RelativePath, SHA256, and Size. stage-manifest.json is
        excluded from its own hash list.

        Icon names the extracted or copied app-icon.<ext> in the same folder
        when the packager declares an IconSource header tag; it is absent
        otherwise and is covered by FileHashes when present.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [AllowNull()][string]$PackagerScriptPath
    )

    $stageRoot = Split-Path -Path $Path -Parent
    $manifestName = Split-Path -Path $Path -Leaf

    # The icon lands before the hashes are computed so it is covered by stage
    # integrity like any other payload file. The packager script is located
    # from the call stack rather than a parameter, which keeps all 284
    # packagers' Write-StageManifest calls unchanged.
    if (-not $PSBoundParameters.ContainsKey('PackagerScriptPath') -or [string]::IsNullOrWhiteSpace($PackagerScriptPath)) {
        $callerFrame = @(Get-PSCallStack | Where-Object { $_.ScriptName -and $_.ScriptName -ne $PSCommandPath }) | Select-Object -First 1
        if ($callerFrame) { $PackagerScriptPath = [string]$callerFrame.ScriptName }
    }
    # A requested install mode rewrites arguments, uninstaller, detection and
    # wrappers before the icon lands and the hashes are computed.
    $requestedMode = Get-RequestedInstallMode
    if ($requestedMode) {
        [void](Set-StageManifestInstallMode -StageRoot $stageRoot -ManifestData $ManifestData -Mode $requestedMode)
    }

    Add-StageIcon -StageRoot $stageRoot -ManifestData $ManifestData -PackagerScriptPath $PackagerScriptPath

    # Operator command overrides resolve before finalization so the launcher
    # chain check inspects the commands the build actually ships.
    $stageOverrides = Get-RequestedCommandOverrides
    if ($stageOverrides) {
        $ManifestData['CommandOverrides'] = @{
            Install   = [string]$stageOverrides.Install
            Uninstall = [string]$stageOverrides.Uninstall
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$stageOverrides.Install)) {
            $ManifestData['InstallCommandLine'] = [string]$stageOverrides.Install
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$stageOverrides.Uninstall)) {
            $ManifestData['UninstallCommandLine'] = [string]$stageOverrides.Uninstall
        }
        Write-Log "Command overrides recorded in manifest."
    }

    # Profile customization, custom source files and script signing all land
    # here, after the icon and before hashing, so every byte the finalizer
    # writes is covered by FileHashes and nothing edits a signed file after.
    if (Get-Command -Name Invoke-StageFinalization -ErrorAction SilentlyContinue) {
        [void](Invoke-StageFinalization -StageRoot $stageRoot -ManifestData $ManifestData -PackagerScriptPath $PackagerScriptPath)
    }

    $fileHashes = Get-StageFileHashes -Root $stageRoot -Exclude @($manifestName)

    $ManifestData['SchemaVersion'] = 4
    $ManifestData['StagedAt'] = (Get-Date -Format 'o')
    $ManifestData['FileHashes'] = @($fileHashes)
    # FileHashes cannot cover the manifest that records them, so a separate
    # digest seals the resolved plan: commands, detection, timing, signing.
    $ManifestData['PlanDigest'] = Get-StageManifestPlanDigest -ManifestData $ManifestData

    $json = $ManifestData | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8 -ErrorAction Stop
    Write-Log "Wrote stage manifest         : $Path"
    Write-Log ("Recorded file hashes         : {0} file(s)" -f @($fileHashes).Count)

    $verification = Compare-StageFileHashes -Root $stageRoot -Expected $fileHashes -Exclude @($manifestName)
    if (-not $verification.Pass) {
        throw ("Stage integrity verification failed: {0}" -f (Format-StageFileHashComparison -Comparison $verification))
    }
    Write-Log ("Stage integrity verified     : {0} file(s)" -f $verification.ExpectedCount)

    # The build record is the workbench's durable provenance for this stage and
    # the only copy of the signed detection/requirement scripts that survives
    # content pruning.
    if (Get-Command -Name Write-BuildRecord -ErrorAction SilentlyContinue) {
        [void](Write-BuildRecord -StageRoot $stageRoot -Manifest $ManifestData)
    }
}

function Format-ApplicationTitleArchitecture {
    <#
    .SYNOPSIS
        Applies the architecture rule to a final application title.
    .DESCRIPTION
        A repeated tag such as (x64) (x64) becomes one. A trailing tag (x64), (x86), (arm64),
        (arm), (32-bit) or (64-bit) goes when -Append is false. Same rule as App Packager Desktop.
    #>
    param([Parameter(Mandatory)][string]$Title, [bool]$Append = $true)
    $tag = '\(\s*(?:x64|x86|arm64|arm|32-bit|64-bit)\s*\)'
    $result = [regex]::Replace($Title, '(' + $tag + ')(\s+\1)+', '$1', 'IgnoreCase')
    if (-not $Append) { $result = [regex]::Replace($result, '\s*' + $tag + '\s*$', '', 'IgnoreCase') }
    return $result.Trim()
}

function Get-PackagedApplicationName {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [ValidateSet('', 'Default', 'IncludeVersion', 'NoVersion')][string]$Mode = $env:APP_PACKAGER_TITLE_MODE
    )
    if (-not $Mode -or $Mode -eq 'Default') { return $AppName }
    # Remove only this release's exact version token, preserving product
    # years, architecture, language, and channel qualifiers.
    $pattern = '(?<![\w.])v?' + [regex]::Escape($Version) + '(?![\w.])'
    if ($Mode -eq 'IncludeVersion' -and [regex]::IsMatch($AppName, $pattern, 'IgnoreCase')) { return $AppName }
    $baseName = [regex]::Replace($AppName, $pattern, '', 'IgnoreCase')
    $baseName = $baseName -replace '\(\s*\)', '' -replace '\s+-\s+(?=\(|$)', ' ' -replace '\s{2,}', ' '
    $baseName = $baseName.Trim().TrimEnd('-').Trim()
    if (-not $baseName) { throw 'Application title cannot consist only of its version.' }
    if ($Mode -eq 'IncludeVersion') { return "$baseName - $Version" }
    return $baseName
}

function Write-PackagePreflight {
    param([Parameter(Mandatory)][string]$AppName, [Parameter(Mandatory)][string]$Version)
    # Only the isolated GUI probe process sets this flag. Stop at the
    # packager's authoritative manifest selection, before share/site writes.
    if ($env:APP_PACKAGER_PACKAGE_PREFLIGHT -eq '1') {
        $identity = @{ AppName = $AppName; Version = $Version } | ConvertTo-Json -Compress
        # Packagers assign Read-StageManifest to a variable: bypass the
        # success pipeline so the probe marker still reaches child stdout.
        [Console]::Out.WriteLine("[APP_PACKAGER_PREFLIGHT] $identity")
        exit 0
    }
}

function Get-FileHeaderBytes {
    <#
    .SYNOPSIS
        The first Count bytes of a file as a byte array, fewer when the
        file is shorter. Throws when the file cannot be opened.
    .DESCRIPTION
        Get-Content -Encoding Byte does not exist in PowerShell 7 and
        -AsByteStream does not exist in Windows PowerShell; a stream read
        works on both hosts.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateRange(1, 1048576)][int]$Count
    )
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $buffer = New-Object byte[] $Count
        $read = $stream.Read($buffer, 0, $Count)
        if ($read -lt $Count) { $buffer = $buffer[0..([Math]::Max($read, 1) - 1)]; if ($read -eq 0) { $buffer = [byte[]]@() } }
        return ,[byte[]]$buffer
    }
    finally { $stream.Dispose() }
}

function Read-StageManifest {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Stage manifest not found: $Path"
    }

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
    $manifest = $json | ConvertFrom-Json

    if (-not $manifest.SchemaVersion) {
        throw "Invalid stage manifest (missing SchemaVersion): $Path"
    }

    # A newer AppPackager can add fields this build would silently ignore, and
    # an ignored detection or signing field publishes the wrong thing.
    $schemaVersion = [int]$manifest.SchemaVersion
    if ($schemaVersion -gt 4) {
        throw "Stage manifest schema version $schemaVersion is newer than this AppPackager understands (maximum 4): $Path. Update AppPackager, or re-run the Stage phase with this version."
    }

    $hasFileHashes = ($manifest.PSObject.Properties.Name -contains 'FileHashes') -and $null -ne $manifest.FileHashes
    if (-not $hasFileHashes) {
        if ($schemaVersion -ge 3) {
            throw "Invalid stage manifest (SchemaVersion $($manifest.SchemaVersion) missing FileHashes): $Path"
        }
        Write-Log "Stage manifest has no file hashes; byte-level integrity verification skipped for this pre-1.0.7 manifest." -Level WARN
    }

    Write-Log "Read stage manifest          : $Path"
    $titleMode = [string]$env:APP_PACKAGER_TITLE_MODE
    if ($manifest.PSObject.Properties['TitleMode'] -and [string]$manifest.TitleMode) { $titleMode = [string]$manifest.TitleMode }
    if ($titleMode -and $titleMode -ne 'Default') {
        $manifest.AppName = Get-PackagedApplicationName -AppName $manifest.AppName -Version $manifest.SoftwareVersion -Mode $titleMode
    }
    if ($env:APP_PACKAGER_TITLE_APPEND_ARCH -eq '0') {
        $manifest.AppName = Format-ApplicationTitleArchitecture -Title ([string]$manifest.AppName) -Append $false
    }
    if ($env:APP_PACKAGER_PACKAGE_PREFLIGHT -eq '1') {
        Write-PackagePreflight -AppName $manifest.AppName -Version $manifest.SoftwareVersion
    }
    return $manifest
}

# ---------------------------------------------------------------------------
# ConfigMgr helpers (continued)
# ---------------------------------------------------------------------------

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][UInt32]$CI_ID,
        [UInt32]$KeepLatest = 1
    )

    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }

    $revs = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        try {
            Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
        }
        catch [System.Management.Automation.ItemNotFoundException] {
            # Removing one revision can retire others in the same chain, so a
            # revision enumerated a moment ago may already be gone. Cleanup is
            # housekeeping and must not fail a completed application.
            Write-Log ("Revision {0} already removed : CI_ID {1}" -f $rev, $CI_ID) -Level DEBUG
        }
    }
}

# ---------------------------------------------------------------------------
# Network path helpers
# ---------------------------------------------------------------------------

function Get-NetworkAppRoot {
    <#
    .SYNOPSIS
        Creates and returns the network content root for an application.

    .DESCRIPTION
        Builds the path <FileServerPath>\Applications\<VendorFolder>\<AppFolder>,
        creating each level if it does not exist. Returns the final path.
    #>
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder
    )

    # Windows strips trailing dots and spaces when creating a directory, but the site stores the
    # literal path and resolves it through \\?\UNC\, which does not normalize: an untrimmed segment
    # distributes as 0x80070003. Trim before the path reaches either side.
    $appsRoot   = Join-Path $FileServerPath "Applications"
    $vendorPath = Join-Path $appsRoot ($VendorFolder.TrimEnd([char]'.', [char]' '))
    $appPath    = Join-Path $vendorPath ($AppFolder.TrimEnd([char]'.', [char]' '))

    Initialize-Folder -Path $appsRoot
    Initialize-Folder -Path $vendorPath
    Initialize-Folder -Path $appPath

    return $appPath
}

# ---------------------------------------------------------------------------
# Content wrapper generation
# ---------------------------------------------------------------------------

function Write-ContentWrappers {
    <#
    .SYNOPSIS
        Creates install/uninstall .bat and .ps1 wrapper files in a content folder.

    .DESCRIPTION
        Writes four files to OutputPath: install.bat, install.ps1, uninstall.bat,
        uninstall.ps1. The .bat files are thin shims that call the corresponding
        .ps1. The .ps1 content is passed as strings by the caller.

        Overwrites existing wrapper files so regenerated staged content remains
        deterministic. All files are written with -Encoding ASCII to avoid BOM
        issues.

    .PARAMETER InstallBatExitCode
        Exit code expression for install.bat. Default: '%ERRORLEVEL%'.
        Use '3010' for products that always require reboot (for example VMware Tools).

    .PARAMETER UninstallBatExitCode
        Exit code expression for uninstall.bat. Default: '%ERRORLEVEL%'.
    #>
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$InstallPs1Content,
        [Parameter(Mandatory)][string]$UninstallPs1Content,
        [string]$InstallBatExitCode   = '%ERRORLEVEL%',
        [string]$UninstallBatExitCode = '%ERRORLEVEL%'
    )

    $InstallPs1Content   = Add-StartProcessFailFast $InstallPs1Content
    $UninstallPs1Content = Add-StartProcessFailFast $UninstallPs1Content

    $installBatPath   = Join-Path $OutputPath "install.bat"
    $installPs1Path   = Join-Path $OutputPath "install.ps1"
    $uninstallBatPath = Join-Path $OutputPath "uninstall.bat"
    $uninstallPs1Path = Join-Path $OutputPath "uninstall.ps1"

    # .bat wrapper template: @echo off, call PowerShell, propagate exit code
    # When exit code override is set (e.g. 3010), only apply on success --
    # real failures must propagate so ConfigMgr can detect them.
    # Signed deployment mode removes the execution-policy argument so the
    # endpoint's configured policy governs the launched script.
    $signed = Get-PolicyFlag -Policy (Resolve-CommonSigningPolicy) -Name 'SignDeployment'
    $installLaunch = Get-DeploymentLauncherCommandLine -ScriptToken '%~dp0install.ps1' -Signed:$signed
    $uninstallLaunch = Get-DeploymentLauncherCommandLine -ScriptToken '%~dp0uninstall.ps1' -Signed:$signed

    $installBat = if ($InstallBatExitCode -eq '%ERRORLEVEL%') {
        (@(
            '@echo off',
            $installLaunch,
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    } else {
        (@(
            '@echo off',
            $installLaunch,
            ('if %ERRORLEVEL% EQU 0 exit /b {0}' -f $InstallBatExitCode),
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    }

    $uninstallBat = if ($UninstallBatExitCode -eq '%ERRORLEVEL%') {
        (@(
            '@echo off',
            $uninstallLaunch,
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    } else {
        (@(
            '@echo off',
            $uninstallLaunch,
            ('if %ERRORLEVEL% EQU 0 exit /b {0}' -f $UninstallBatExitCode),
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
    }

    $files = @(
        @{ Path = $installBatPath;   Content = $installBat;          Label = 'install.bat' },
        @{ Path = $installPs1Path;   Content = $InstallPs1Content;   Label = 'install.ps1' },
        @{ Path = $uninstallBatPath; Content = $uninstallBat;        Label = 'uninstall.bat' },
        @{ Path = $uninstallPs1Path; Content = $UninstallPs1Content; Label = 'uninstall.ps1' }
    )

    foreach ($f in $files) {
        Set-Content -LiteralPath $f.Path -Value $f.Content -Encoding ASCII -Force -ErrorAction Stop
        Write-Log "Wrote wrapper                : $($f.Label)"
    }
}

function ConvertTo-SingleQuotedContent {
    <#
    .SYNOPSIS
        Escapes text for use inside a single-quoted PowerShell string literal
        in generated script content.

    .DESCRIPTION
        PowerShell ends a single-quoted string at an apostrophe and also at the
        typographic quotes U+2018, U+2019, U+201A and U+201B. Doubling only the
        apostrophe leaves the other four able to end the literal early, so
        text that came from a vendor page or a file name could add statements
        to the generated script.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    return [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($Text)
}

function Add-StartProcessFailFast {
    <#
    .SYNOPSIS
        Wraps each captured Start-Process call in wrapper script content so a
        target that cannot start ends the script with exit code 1.

    .DESCRIPTION
        Start-Process reports a target that cannot start as a statement-
        terminating error, which -ErrorAction does not turn into a script-
        terminating one. The script then continues, reads ExitCode from a null
        process, and exits 0, which reports success for an install or
        uninstall that never ran. Inside try/catch the failure is caught and
        the script exits 1.

        Only single-line assignments of the form $name = Start-Process ... are
        changed. A line that names -ErrorAction, ends in a line continuation,
        or carries a pipeline is left as written. Remove-StartProcessFailFast
        reverses the change.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content)

    # A backtick that escapes a character inside an argument is part of the
    # call; a backtick at the line end is a continuation and stays unmatched.
    # A line with a statement separator or a comment mark is left as written.
    $pattern = '(?m)^(?<indent>[ \t]*)(?<call>\$\w+[ \t]*=[ \t]*Start-Process\b(?:(?!-ErrorAction)(?:`[^\r\n]|[^\r\n`|;#]))*?)(?<end>[ \t]*)(?=\r?$)'
    return [regex]::Replace($Content, $pattern, '${indent}try { ${call} -ErrorAction Stop } catch { Write-Error $$_.Exception.Message; exit 1 }${end}')
}

function Remove-StartProcessFailFast {
    <#
    .SYNOPSIS
        Returns wrapper script content as it was before Add-StartProcessFailFast.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content)

    $pattern = '(?m)^(?<indent>[ \t]*)try \{ (?<call>\$\w+[ \t]*=[ \t]*Start-Process\b[^\r\n]*?) -ErrorAction Stop \} catch \{ Write-Error \$_\.Exception\.Message; exit 1 \}(?<end>[ \t]*)(?=\r?$)'
    return [regex]::Replace($Content, $pattern, '${indent}${call}${end}')
}

function New-MsiWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an MSI product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList (avoiding quoting issues). Returns a hashtable
        with Install and Uninstall keys.

        -ExtraInstallArgs: optional MSI properties (for example "APITOKEN=xxx",
        'ASSIGNMENTOPTIONS="--grant-easy-access"') appended to the install
        command line. Each entry becomes one Start-Process argument.

        -PostInstallKillProcesses: optional array of process names (no .exe)
        that the install.ps1 should Stop-Process after msiexec returns. Used
        to kill the installer-spawned GUI that runs under the SYSTEM context
        (TeamViewer, TeamViewer Host, anything else that helpfully launches
        a post-install splash).
    #>
    param(
        [Parameter(Mandatory)][string]$MsiFileName,
        [string[]]$ExtraInstallArgs = @(),
        [string[]]$PostInstallKillProcesses = @(),
        [string[]]$ExtraUninstallArgs = @()
    )

    # The filename lands inside a single-quoted literal in the generated
    # script; an unescaped quote character terminates the string early.
    $MsiFileName = ConvertTo-SingleQuotedContent $MsiFileName

    $installLines = @(
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$args = @(''/i'', "`"$msiPath`"", ''/qn'', ''/norestart'')'
    )
    foreach ($extra in $ExtraInstallArgs) {
        $escaped = ConvertTo-SingleQuotedContent $extra
        $installLines += ('$args += ''{0}''' -f $escaped)
    }
    $installLines += '$proc = Start-Process msiexec.exe -ArgumentList $args -Wait -PassThru -NoNewWindow'
    $installLines += '$exit = $proc.ExitCode'
    if ($PostInstallKillProcesses.Count -gt 0) {
        $procList = ($PostInstallKillProcesses | ForEach-Object { "'" + (ConvertTo-SingleQuotedContent $_) + "'" }) -join ', '
        $installLines += ('foreach ($pn in @({0})) {{ Get-Process -Name $pn -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }}' -f $procList)
    }
    $installLines += 'exit $exit'
    $install = $installLines -join "`r`n"

    $uninstallLines = @(
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$args = @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'')'
    )
    foreach ($extra in $ExtraUninstallArgs) {
        $uninstallLines += ('$args += ''{0}''' -f (ConvertTo-SingleQuotedContent $extra))
    }
    $uninstallLines += '$proc = Start-Process msiexec.exe -ArgumentList $args -Wait -PassThru -NoNewWindow'
    $uninstallLines += 'exit $proc.ExitCode'
    $uninstall = $uninstallLines -join "`r`n"

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

function New-MsixWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an MSIX/APPX
        package, using the Script deployment-type pattern (install.bat +
        install.ps1) so ConfigMgr treats it the same way as MSI / EXE packagers.

    .DESCRIPTION
        ConfigMgr also supports a native MSIX deployment type through
        Add-CMWindowsAppxDeploymentType, but the house rule is Script
        deployment for everything we can shoehorn that way (single code
        path, uniform logging, consistent detection authoring). These
        wrappers use Add-AppxProvisionedPackage for deployment-wide
        (per-device) installs. Per-user MSIX can layer on top of this
        pattern but isn't the default.

        -MsixFileName  : name of the .msix / .appx / .msixbundle in the
                         package content folder.
        -Provisioned   : when true (default), installs through
                         Add-AppxProvisionedPackage -Online so all users
                         (including new accounts) get the app. Uninstall
                         uses Remove-AppxProvisionedPackage.
                         When false, installs per-user through Add-AppxPackage.
        -SignatureSha1 : optional SHA1 thumbprint of the expected signing
                         certificate. When set, install.ps1 verifies the
                         MSIX is signed by that cert before invoking the
                         AppX cmdlet (catches tampered / wrong-package
                         downloads before deployment).
    #>
    param(
        [Parameter(Mandatory)][string]$MsixFileName,
        [bool]$Provisioned = $true,
        [string]$SignatureSha1 = ''
    )

    # The file name lands inside single-quoted literals in the generated scripts.
    $MsixFileName = ConvertTo-SingleQuotedContent $MsixFileName

    $sigCheckLines = @()
    if (-not [string]::IsNullOrWhiteSpace($SignatureSha1)) {
        $sigCheckLines = @(
            ('$expected = ''{0}''' -f (ConvertTo-SingleQuotedContent $SignatureSha1)),
            '$sig = Get-AuthenticodeSignature -LiteralPath $msixPath',
            'if ($sig.Status -ne ''Valid'') { Write-Error "MSIX signature not valid: $($sig.Status)"; exit 2 }',
            'if ($sig.SignerCertificate.Thumbprint -ne $expected) { Write-Error "MSIX signed by unexpected cert (got $($sig.SignerCertificate.Thumbprint), expected $expected)"; exit 3 }'
        )
    }

    if ($Provisioned) {
        $installBody = @(
            'try { Add-AppxProvisionedPackage -Online -PackagePath $msixPath -SkipLicense -ErrorAction Stop | Out-Null } catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }'
        )
        $uninstallBody = @(
            'Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem',
            ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
            'if (-not (Test-Path -LiteralPath $msixPath)) { exit 0 }',
            '# Provisioned-package removal needs the PackageName, which is stored',
            '# inside the MSIX AppxManifest. Import-Metadata pattern: read the',
            '# manifest during install and stash the name; for the uninstall path',
            '# re-read the manifest on the client.',
            '$zip = [System.IO.Compression.ZipFile]::OpenRead($msixPath)',
            'try {',
            '    $entry = $zip.GetEntry(''AppxManifest.xml'')',
            '    if (-not $entry) { Write-Error ''AppxManifest.xml not found in MSIX''; exit 4 }',
            '    $reader = New-Object System.IO.StreamReader($entry.Open())',
            '    try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }',
            '} finally { $zip.Dispose() }',
            '$identity = $manifest.Package.Identity',
            '$pkgFullName = ''{0}_{1}_{2}__{3}'' -f $identity.Name, $identity.Version, $identity.ProcessorArchitecture, $identity.PublisherId',
            '$prov = Get-AppxProvisionedPackage -Online -ErrorAction Stop | Where-Object { $_.PackageName -eq $pkgFullName -or $_.DisplayName -eq $identity.Name }',
            'try {',
            '    if ($prov) { Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null }',
            '    Get-AppxPackage -AllUsers -Name $identity.Name -ErrorAction Stop | ForEach-Object { Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction Stop }',
            '} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }',
            'exit 0'
        )
    } else {
        $installBody = @(
            'try { Add-AppxPackage -Path $msixPath -ErrorAction Stop } catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }'
        )
        $uninstallBody = @(
            'Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem',
            ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
            'if (-not (Test-Path -LiteralPath $msixPath)) { exit 0 }',
            '$zip = [System.IO.Compression.ZipFile]::OpenRead($msixPath)',
            'try {',
            '    $entry = $zip.GetEntry(''AppxManifest.xml'')',
            '    if (-not $entry) { Write-Error ''AppxManifest.xml not found in MSIX''; exit 4 }',
            '    $reader = New-Object System.IO.StreamReader($entry.Open())',
            '    try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }',
            '} finally { $zip.Dispose() }',
            '$name = $manifest.Package.Identity.Name',
            'try { Get-AppxPackage -Name $name -ErrorAction Stop | Remove-AppxPackage -ErrorAction Stop } catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }',
            'exit 0'
        )
    }

    $installLines = @(
        'Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem',
        ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
        'if (-not (Test-Path -LiteralPath $msixPath)) { Write-Error "MSIX not found: $msixPath"; exit 1 }'
    ) + $sigCheckLines + $installBody + @('exit 0')

    return @{
        Install   = ($installLines -join "`r`n")
        Uninstall = ($uninstallBody -join "`r`n")
    }
}

function New-ExeWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an EXE product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList for the installer EXE. Returns a hashtable
        with Install and Uninstall keys.

        For products where uninstall uses a different command (for example registry
        lookup, msiexec), the caller should build uninstall content directly
        and pass it to Write-ContentWrappers.

        -PostInstallKillProcesses: process names (no .exe) the installer
        launches on success. Start-Process -Wait also waits for those
        descendants, so the wrapper then waits for the installer alone and
        closes the launched instances, which a system-context install would
        otherwise leave running where no user can see them.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$InstallArgs,
        [Parameter(Mandatory)][string]$UninstallCommand,
        [string]$UninstallArgs = '',
        [string[]]$PostInstallKillProcesses = @()
    )

    # Filename and uninstall path land inside single-quoted literals in the
    # generated script; an unescaped quote character terminates the string
    # early. InstallArgs/UninstallArgs are excluded: they are caller-authored
    # PowerShell element lists interpolated into @() as-is.
    $InstallerFileName = ConvertTo-SingleQuotedContent $InstallerFileName
    $UninstallCommand  = ConvertTo-SingleQuotedContent $UninstallCommand

    if ($PostInstallKillProcesses.Count -gt 0) {
        $procList = ($PostInstallKillProcesses | ForEach-Object { "'" + (ConvertTo-SingleQuotedContent $_) + "'" }) -join ', '
        # Without -Wait, Windows PowerShell 5.1 leaves ExitCode null after the
        # process exits unless the handle is read while it runs.
        $install = (
            ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
            ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -PassThru -NoNewWindow' -f $InstallArgs),
            '$null = $proc.Handle',
            '$proc.WaitForExit()',
            '$exit = $proc.ExitCode',
            'Start-Sleep -Seconds 3',
            ('foreach ($pn in @({0})) {{ Get-Process -Name $pn -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }}' -f $procList),
            'exit $exit'
        ) -join "`r`n"
    }
    else {
        $install = (
            ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
            ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $InstallArgs),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }

    # The uninstall path may carry %VAR% segments (a per-user uninstaller
    # under %LOCALAPPDATA%); a single-quoted literal would hand them to
    # Start-Process verbatim, so the wrapper expands them at run time in the
    # context the deployment type runs under.
    $uninstallPathLine = ('$uninstallPath = [Environment]::ExpandEnvironmentVariables(''{0}'')' -f $UninstallCommand)
    if ($UninstallArgs -ne '') {
        $uninstall = (
            $uninstallPathLine,
            ('$proc = Start-Process -FilePath $uninstallPath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $UninstallArgs),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }
    else {
        $uninstall = (
            $uninstallPathLine,
            '$proc = Start-Process -FilePath $uninstallPath -Wait -PassThru -NoNewWindow',
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

function New-TargetSecretBlock {
    <#
    .SYNOPSIS
        Returns install-script lines that give each target its own random
        secret.

    .DESCRIPTION
        The generated lines create a 32 character random value on the target,
        store it computer-protected (DPAPI, LocalMachine) in a folder that only
        SYSTEM, Administrators and the installing account can read, and leave it
        in the named variable for the installer argument list. A later run reuses the stored value so
        the file always matches the product, but only from a folder that those
        principals own and alone can access. If the secret cannot be generated,
        stored or trusted, the script exits 1 before the installer starts.
    #>
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9-]{1,64}$')][string]$FolderName,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9.-]{1,64}$')][string]$FileName,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z][A-Za-z0-9]{0,31}$')][string]$VariableName,
        [string]$Label = 'product'
    )

    if ($Label -notmatch '^[A-Za-z0-9 ]{1,40}$') { throw "Label '$Label' is not usable." }

    $lines = @(
        'try {',
        '    Add-Type -AssemblyName System.Security -ErrorAction Stop',
        ('    $secretDir = Join-Path $env:ProgramData ''{0}''' -f $FolderName),
        ('    $secretFile = Join-Path $secretDir ''{0}''' -f $FileName),
        '    $trusted = @(''S-1-5-18'', ''S-1-5-32-544'', [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)',
        '    if (Test-Path -LiteralPath $secretDir) {',
        '        # A folder that already exists is used only when SYSTEM, Administrators or this account own it and',
        '        # are the only principals it grants access to; a standard user can create the folder first.',
        '        $acl = Get-Acl -LiteralPath $secretDir',
        '        $ownerSid = $acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value',
        '        if ($trusted -notcontains $ownerSid) { throw "the folder $secretDir is owned by $ownerSid" }',
        '        foreach ($rule in $acl.Access) {',
        '            $ruleSid = $rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value',
        '            if ($rule.AccessControlType -eq ''Allow'' -and $trusted -notcontains $ruleSid) { throw "the folder $secretDir grants access to $ruleSid" }',
        '        }',
        '    }',
        '    else {',
        '        New-Item -ItemType Directory -Path $secretDir -Force -ErrorAction Stop | Out-Null',
        '        $grants = @(''*S-1-5-18:(OI)(CI)F'', ''*S-1-5-32-544:(OI)(CI)F'')',
        '        if ($trusted[2] -ne ''S-1-5-18'') { $grants += (''*'' + $trusted[2] + '':(OI)(CI)F'') }',
        '        & icacls.exe $secretDir /inheritance:r /grant:r $grants | Out-Null',
        '        if ($LASTEXITCODE -ne 0) { Remove-Item -LiteralPath $secretDir -Recurse -Force -ErrorAction SilentlyContinue; throw "icacls failed with exit code $LASTEXITCODE" }',
        '    }',
        '    if (Test-Path -LiteralPath $secretFile) {',
        '        $plain = [System.Security.Cryptography.ProtectedData]::Unprotect([System.IO.File]::ReadAllBytes($secretFile), $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)',
        ('        ${0} = [System.Text.Encoding]::UTF8.GetString($plain)' -f $VariableName),
        '    }',
        '    else {',
        '        $alphabet = ''ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789''',
        '        $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()',
        '        $chars = New-Object char[] 32',
        '        $byte = New-Object byte[] 1',
        '        for ($index = 0; $index -lt $chars.Length; ) {',
        '            $rng.GetBytes($byte)',
        '            if ($byte[0] -lt (256 - (256 % $alphabet.Length))) { $chars[$index] = $alphabet[$byte[0] % $alphabet.Length]; $index++ }',
        '        }',
        '        $rng.Dispose()',
        ('        ${0} = -join $chars' -f $VariableName),
        ('        $protected = [System.Security.Cryptography.ProtectedData]::Protect([System.Text.Encoding]::UTF8.GetBytes(${0}), $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)' -f $VariableName),
        '        [System.IO.File]::WriteAllBytes($secretFile, $protected)',
        '    }',
        '}',
        'catch {',
        ('    Write-Error ("Could not create the {0} secret: " + $_.Exception.Message)' -f $Label),
        '    exit 1',
        '}'
    )
    return ($lines -join "`r`n")
}

function New-PostgreSqlSuperuserSecretBlock {
    <#
    .SYNOPSIS
        Install-script lines that give each target its own PostgreSQL
        superuser password in $pgSuperPassword.
    #>
    param([Parameter(Mandatory)][ValidatePattern('^\d{1,3}$')][string]$MajorVersion)

    return (New-TargetSecretBlock -FolderName 'PostgreSQL-Superuser-Secret' -FileName ('postgresql-{0}.dpapi' -f $MajorVersion) -VariableName 'pgSuperPassword' -Label 'PostgreSQL superuser')
}
# ---------------------------------------------------------------------------
# Supersedence: version rule, Exact guard, supersedence mode
# ---------------------------------------------------------------------------
# The rules are the rules of App Packager Desktop (Services/Supersedence). The version rule sets the
# operator of the version clause in the published detection: Minimum is GreaterEquals (this version or
# newer), Exact is IsEquals (this version only). The supersedence mode sets how the new ConfigMgr
# application relates to the older one of the same product line: Off, Replace, ReplaceAndUninstall.
# A mode is a token here ('Off', 'Replace', 'ReplaceAndUninstall'); the text of a mode is 'Replace and uninstall'.

$script:SupersedenceBlockedText = 'Replace and uninstall needs the version rule Exact. With Minimum, the detection of the old application still finds the new version after the uninstall, and the uninstall is reported as failed. Set the version rule to Exact, or set the supersedence mode to Replace or Off.'
$script:SupersedenceMarkerPrefix = 'APD-Line:'

function Get-SupersedenceProperty {
    # The value of a property, or $null when the object has none (safe under strict mode).
    param($InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }
    $p = $InputObject.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Get-SupersedenceMapEntry {
    # The key and value pairs of a hashtable or of a pscustomobject used as a map.
    param($Map)
    $entries = @()
    if ($null -eq $Map) { return $entries }
    if ($Map -is [System.Collections.IDictionary]) {
        foreach ($key in $Map.Keys) { $entries += [pscustomobject]@{ Key = [string]$key; Value = $Map[$key] } }
    }
    else {
        foreach ($p in $Map.PSObject.Properties) { $entries += [pscustomobject]@{ Key = [string]$p.Name; Value = $p.Value } }
    }
    return $entries
}

function ConvertTo-SupersedenceRule {
    # 'Minimum' or 'Exact' (any letter case), or $null for another word.
    param([AllowNull()][string]$Text)
    switch (([string]$Text).Trim().ToLowerInvariant()) {
        'minimum' { return 'Minimum' }
        'exact' { return 'Exact' }
        default { return $null }
    }
}

function ConvertTo-SupersedenceMode {
    # 'Off', 'Replace' or 'ReplaceAndUninstall', or $null for another word.
    param([AllowNull()][string]$Text)
    switch (([string]$Text).Trim().ToLowerInvariant().Replace('_', '-').Replace(' ', '-')) {
        'off' { return 'Off' }
        'replace' { return 'Replace' }
        { $_ -in @('replace-uninstall', 'replaceanduninstall', 'replace-and-uninstall') } { return 'ReplaceAndUninstall' }
        default { return $null }
    }
}

function Get-SupersedenceModeText {
    param([Parameter(Mandatory)][string]$Mode)
    switch ($Mode) {
        'Replace' { return 'Replace' }
        'ReplaceAndUninstall' { return 'Replace and uninstall' }
        default { return 'Off' }
    }
}

function New-SupersedenceSettings {
    # The shipped settings: Minimum, Minimum and Off.
    [CmdletBinding()]
    param()
    return [pscustomobject]@{
        ConfigMgrVersionRule = 'Minimum'
        IntuneVersionRule    = 'Minimum'
        Mode                 = 'Off'
        TitleOverrides       = @{}
        InstalledValues      = @{}
        AcceptedExactBlocks  = @()
    }
}

function Get-SupersedenceOverrideKey {
    # The key of a per-title override: the application identifier; $null for a title that has none.
    param([AllowNull()][string]$ApplicationId)
    if ([string]::IsNullOrWhiteSpace($ApplicationId)) { return $null }
    return $ApplicationId.Trim()
}

function Get-SupersedenceBlockedCombination {
    # The refusal text when the mode is ReplaceAndUninstall and the ConfigMgr rule is Minimum; $null otherwise.
    param([Parameter(Mandatory)][string]$ConfigMgrRule, [Parameter(Mandatory)][string]$Mode)
    if ($Mode -eq 'ReplaceAndUninstall' -and $ConfigMgrRule -eq 'Minimum') { return $script:SupersedenceBlockedText }
    return $null
}

function Resolve-Supersedence {
    <#
    .SYNOPSIS
        Resolves the version rules and the supersedence mode of one publish.
    .DESCRIPTION
        Each setting resolves by itself in the order command line, per-title override, site default. The
        blocked combination is checked after both are resolved. A word that is not allowed throws.
        CommandLine is an object or a hashtable with VersionRule and Mode (the words as typed), or $null.
    #>
    param(
        [Parameter(Mandatory)]$Settings,
        [AllowNull()][string]$ApplicationId,
        $CommandLine = $null
    )

    $title = $null
    $key = Get-SupersedenceOverrideKey -ApplicationId $ApplicationId
    if ($key) {
        foreach ($entry in @(Get-SupersedenceMapEntry -Map (Get-SupersedenceProperty $Settings 'TitleOverrides'))) {
            if ($entry.Key.Trim() -eq $key) { $title = $entry.Value; break }
        }
    }
    $cliRule = [string](Get-SupersedenceProperty $CommandLine 'VersionRule')
    $cliMode = [string](Get-SupersedenceProperty $CommandLine 'Mode')
    $titleRule = [string](Get-SupersedenceProperty $title 'VersionRule')
    $titleMode = [string](Get-SupersedenceProperty $title 'Mode')

    $rule = {
        param($cli, $over, $site, $name)
        if (-not [string]::IsNullOrWhiteSpace($cli)) { $where = 'the command line option -VersionRule'; $text = $cli; $source = 'CommandLine' }
        elseif (-not [string]::IsNullOrWhiteSpace($over)) { $where = "the override of this title ($name)"; $text = $over; $source = 'TitleOverride' }
        else { $where = "the site default ($name)"; $text = $site; $source = 'SiteDefault' }
        $value = ConvertTo-SupersedenceRule -Text $text
        if (-not $value) { throw "The version rule '$text' in $where is not Minimum or Exact." }
        return [pscustomobject]@{ Value = $value; Source = $source }
    }
    # The per-title override sets one version rule for the title; it applies to both targets.
    $configMgrRule = & $rule $cliRule $titleRule ([string](Get-SupersedenceProperty $Settings 'ConfigMgrVersionRule')) 'ConfigMgr version rule'
    $intuneRule = & $rule $cliRule $titleRule ([string](Get-SupersedenceProperty $Settings 'IntuneVersionRule')) 'Intune version rule'

    if (-not [string]::IsNullOrWhiteSpace($cliMode)) { $where = 'the command line option -Supersedence'; $text = $cliMode; $source = 'CommandLine' }
    elseif (-not [string]::IsNullOrWhiteSpace($titleMode)) { $where = 'the override of this title'; $text = $titleMode; $source = 'TitleOverride' }
    else { $where = 'the site default'; $text = [string](Get-SupersedenceProperty $Settings 'Mode'); $source = 'SiteDefault' }
    $modeValue = ConvertTo-SupersedenceMode -Text $text
    if (-not $modeValue) { throw "The supersedence mode '$text' in $where is not Off, Replace or Replace and uninstall." }
    $mode = [pscustomobject]@{ Value = $modeValue; Source = $source }

    return [pscustomobject]@{
        ConfigMgrRule = $configMgrRule
        IntuneRule    = $intuneRule
        Mode          = $mode
        Blocked       = (Get-SupersedenceBlockedCombination -ConfigMgrRule $configMgrRule.Value -Mode $mode.Value)
    }
}

# ---- the version rule on a detection ----

function Get-SupersedenceClauseDescription {
    # The words of a clause for a log line or a message.
    param([Parameter(Mandatory)]$Clause)
    if ([string](Get-SupersedenceProperty $Clause 'Type') -eq 'File') {
        return ('file {0}\{1}' -f (Get-SupersedenceProperty $Clause 'FilePath'), (Get-SupersedenceProperty $Clause 'FileName'))
    }
    return ('registry {0}\{1}' -f (Get-SupersedenceProperty $Clause 'RegistryKeyRelative'), (Get-SupersedenceProperty $Clause 'ValueName'))
}

function Test-SupersedenceVersionClause {
    # True for the clause that the version rule may change: a File version clause or a registry version clause
    # whose operator is GreaterEquals and whose expected value is the packaged version.
    param([Parameter(Mandatory)]$Clause, [AllowNull()][string]$PackagedVersion)
    if ([string]::IsNullOrWhiteSpace($PackagedVersion)) { return $false }
    $expected = [string](Get-SupersedenceProperty $Clause 'ExpectedValue')
    if ($expected.Trim() -ne $PackagedVersion.Trim()) { return $false }
    $type = [string](Get-SupersedenceProperty $Clause 'Type')
    $operator = [string](Get-SupersedenceProperty $Clause 'Operator')
    $propertyType = [string](Get-SupersedenceProperty $Clause 'PropertyType')
    if ($type -eq 'File') {
        if ([string]::IsNullOrWhiteSpace($propertyType)) { $propertyType = 'Version' }
        if ([string]::IsNullOrWhiteSpace($operator)) { $operator = 'GreaterEquals' }
        return ($propertyType -eq 'Version' -and $operator -eq 'GreaterEquals')
    }
    if ($type -eq 'RegistryKeyValue') {
        $isVersion = ($propertyType -eq 'Version') -or (([string](Get-SupersedenceProperty $Clause 'ValueName')).IndexOf('Version', [StringComparison]::OrdinalIgnoreCase) -ge 0)
        return ($isVersion -and $operator -eq 'GreaterEquals')
    }
    return $false
}

function Copy-SupersedenceObject {
    # A deep copy of a detection (objects, arrays and values); the original is never changed.
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string] -or $InputObject.GetType().IsValueType) { return $InputObject }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [System.Collections.IDictionary]) {
        return @(foreach ($item in $InputObject) { Copy-SupersedenceObject -InputObject $item })
    }
    $copy = [ordered]@{}
    foreach ($p in $InputObject.PSObject.Properties) { $copy[$p.Name] = Copy-SupersedenceObject -InputObject $p.Value }
    return [pscustomobject]$copy
}

function Set-SupersedenceOperator {
    param([Parameter(Mandatory)]$Clause, [Parameter(Mandatory)][string]$Operator)
    if ($Clause.PSObject.Properties['Operator']) { $Clause.Operator = $Operator }
    else { Add-Member -InputObject $Clause -NotePropertyName Operator -NotePropertyValue $Operator }
}

function Invoke-SupersedenceWalk {
    # Writes IsEquals to each version clause that the guard allows, and records one outcome per clause.
    param($Detection, [string]$Version, [scriptblock]$Guard, [System.Collections.ArrayList]$Outcomes)
    if ($null -eq $Detection) { return }
    if ([string](Get-SupersedenceProperty $Detection 'Type') -eq 'Compound') {
        foreach ($clause in @(Get-SupersedenceProperty $Detection 'Clauses')) { Invoke-SupersedenceWalk -Detection $clause -Version $Version -Guard $Guard -Outcomes $Outcomes }
        return
    }
    if (-not (Test-SupersedenceVersionClause -Clause $Detection -PackagedVersion $Version)) { return }
    $verdict = $null
    if ($Guard) { $verdict = & $Guard $Detection }
    if (-not $verdict) { $verdict = [pscustomobject]@{ Result = 'ExactBlocked'; Detail = 'No installed value is recorded for this version.' } }
    if ($verdict.Result -eq 'ExactApplied') { Set-SupersedenceOperator -Clause $Detection -Operator 'IsEquals' }
    [void]$Outcomes.Add([pscustomobject]@{ Clause = (Get-SupersedenceClauseDescription -Clause $Detection); Result = $verdict.Result; Detail = $verdict.Detail })
}

function Set-DetectionVersionRule {
    <#
    .SYNOPSIS
        A copy of a detection with the version rule written.
    .DESCRIPTION
        With the rule Minimum the copy equals the original and there are no outcomes. The guard (a script block
        that takes a clause and returns an object with Result and Detail) is asked once for each clause that would change.
    #>
    param(
        [Parameter(Mandatory)]$Detection,
        [AllowNull()][string]$PackagedVersion,
        [Parameter(Mandatory)][ValidateSet('Minimum', 'Exact')][string]$Rule,
        [AllowNull()][scriptblock]$Guard
    )
    $copy = Copy-SupersedenceObject -InputObject $Detection
    $outcomes = New-Object System.Collections.ArrayList
    if ($Rule -eq 'Exact') { Invoke-SupersedenceWalk -Detection $copy -Version $PackagedVersion -Guard $Guard -Outcomes $outcomes }
    return [pscustomobject]@{ Detection = $copy; Outcomes = @($outcomes) }
}

function Set-ManifestVersionRule {
    <#
    .SYNOPSIS
        A copy of a manifest with the version rule written into its detection and into the detection of each deployment type.
    .DESCRIPTION
        The manifest is not changed. With the rule Minimum the original object comes back and there are no outcomes.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)][ValidateSet('Minimum', 'Exact')][string]$Rule,
        [AllowNull()][scriptblock]$Guard
    )
    if ($Rule -ne 'Exact') { return [pscustomobject]@{ Manifest = $Manifest; Outcomes = @() } }
    $outcomes = New-Object System.Collections.ArrayList
    $version = [string](Get-SupersedenceProperty $Manifest 'SoftwareVersion')
    $copy = $Manifest.PSObject.Copy()
    $detection = Copy-SupersedenceObject -InputObject (Get-SupersedenceProperty $Manifest 'Detection')
    Invoke-SupersedenceWalk -Detection $detection -Version $version -Guard $Guard -Outcomes $outcomes
    if ($copy.PSObject.Properties['Detection']) { $copy.Detection = $detection } elseif ($null -ne $detection) { Add-Member -InputObject $copy -NotePropertyName Detection -NotePropertyValue $detection }
    $types = Get-SupersedenceProperty $Manifest 'DeploymentTypes'
    if ($types) {
        $newTypes = @()
        foreach ($type in @($types)) {
            $typeCopy = $type.PSObject.Copy()
            $own = Get-SupersedenceProperty $type 'Detection'
            if ($own) {
                $ownCopy = Copy-SupersedenceObject -InputObject $own
                Invoke-SupersedenceWalk -Detection $ownCopy -Version $version -Guard $Guard -Outcomes $outcomes
                $typeCopy.Detection = $ownCopy
            }
            $newTypes += $typeCopy
        }
        $copy.DeploymentTypes = $newTypes
    }
    return [pscustomobject]@{ Manifest = $copy; Outcomes = @($outcomes) }
}

# ---- the Exact guard ----

function Get-SupersedenceRecordKey {
    # The key of a recorded installed value or of an accepted block: the title and the version.
    param([Parameter(Mandatory)][string]$TitleKey, [Parameter(Mandatory)][string]$Version)
    return ('{0}|{1}' -f $TitleKey.Trim(), $Version.Trim())
}

function Test-SupersedenceSameVersionText {
    # The text compare of the specification: + and - become a dot, and no zero part is ignored.
    param([AllowNull()][string]$A, [AllowNull()][string]$B)
    $normalize = { param($v) ([string]$v).Trim().Replace('+', '.').Replace('-', '.') }
    return ((& $normalize $A) -eq (& $normalize $B))
}

function Test-SupersedenceMsiProductClause {
    # True for the registry clause on the Windows Installer product of an MSI: Uninstall\{ProductCode}, DisplayVersion.
    param([Parameter(Mandatory)]$Manifest, [Parameter(Mandatory)]$Clause)
    $productCode = [string](Get-SupersedenceProperty $Manifest 'ProductCode')
    return (([string](Get-SupersedenceProperty $Manifest 'InstallerType')) -eq 'MSI' -and
        ([string](Get-SupersedenceProperty $Clause 'Type')) -eq 'RegistryKeyValue' -and
        ([string](Get-SupersedenceProperty $Clause 'ValueName')) -eq 'DisplayVersion' -and
        -not [string]::IsNullOrWhiteSpace($productCode) -and
        ([string](Get-SupersedenceProperty $Clause 'RegistryKeyRelative')).IndexOf($productCode, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Get-ExactGuardVerdict {
    <#
    .SYNOPSIS
        The verdict of the Exact guard for one clause.
    .DESCRIPTION
        IsEquals hides nothing: a clause that says 3.7.7 does not find a product that reports 3.7.7.1234. So Exact
        applies to a clause only when a value that was read from the installed product equals the expected value:
        the ProductVersion of an MSI for the clause on the Windows Installer product, and for any other clause the
        value that a test install recorded for the title and the version. Result is ExactApplied, ExactBlocked or
        ExactAcceptedBlock.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)]$Clause,
        [AllowNull()][string]$TitleKey,
        [AllowNull()][scriptblock]$MsiProductVersion,
        $Recorded,
        [string[]]$Accepted
    )
    $expected = [string](Get-SupersedenceProperty $Clause 'ExpectedValue')
    $version = [string](Get-SupersedenceProperty $Manifest 'SoftwareVersion')
    $key = $null
    if (-not [string]::IsNullOrWhiteSpace($TitleKey)) { $key = Get-SupersedenceRecordKey -TitleKey $TitleKey -Version $version }
    $installed = $null
    $source = ''
    if (Test-SupersedenceMsiProductClause -Manifest $Manifest -Clause $Clause) {
        if ($MsiProductVersion) { $installed = & $MsiProductVersion $Manifest }
        $source = 'the ProductVersion of the MSI'
    }
    else {
        if ($key) {
            foreach ($entry in @(Get-SupersedenceMapEntry -Map $Recorded)) { if ($entry.Key -eq $key) { $installed = [string]$entry.Value } }
        }
        $source = 'the value that a test install recorded'
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$installed) -and (Test-SupersedenceSameVersionText -A $installed -B $expected)) {
        return [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null }
    }
    $describe = Get-SupersedenceClauseDescription -Clause $Clause
    if ([string]::IsNullOrWhiteSpace([string]$installed)) {
        $reason = "No installed value exists for $describe (version $version). Run the test install of the title to record it."
    }
    else {
        $reason = "The installed value '$installed' ($source) differs from the expected value '$expected' of $describe."
    }
    $isAccepted = $false
    if ($key) { foreach ($a in @($Accepted)) { if ($a -eq $key) { $isAccepted = $true } } }
    return [pscustomobject]@{ Result = $(if ($isAccepted) { 'ExactAcceptedBlock' } else { 'ExactBlocked' }); Detail = $reason }
}

# ---- the identity marker ----

function Get-SupersedenceIdentityMarker {
    # The identity record that a ConfigMgr publish writes into the keywords of the application: the title
    # identity (the application identifier), the profile and the architecture. $null for a title with no identifier.
    param([Parameter(Mandatory)]$Manifest, [AllowNull()][string]$FallbackApplicationId)
    $applicationId = [string](Get-SupersedenceProperty $Manifest 'ApplicationId')
    if ([string]::IsNullOrWhiteSpace($applicationId)) { $applicationId = $FallbackApplicationId }
    if ([string]::IsNullOrWhiteSpace($applicationId)) { return $null }
    $profile = [string](Get-SupersedenceProperty $Manifest 'ProfileId')
    if ([string]::IsNullOrWhiteSpace($profile)) { $profile = 'default' }
    $architecture = [string](Get-SupersedenceProperty $Manifest 'Architecture')
    if ([string]::IsNullOrWhiteSpace($architecture)) { $architecture = 'x64' }
    return ('{0}{1}|{2}|{3}' -f $script:SupersedenceMarkerPrefix, $applicationId.Trim(), $profile.Trim(), $architecture.Trim())
}

# ---- the planner ----

function Get-SupersedencePlan {
    <#
    .SYNOPSIS
        What a publish does with the version rule: the manifest that each target publishes, the identity marker,
        the lines for the log and the refusal when the row cannot be published.
    .DESCRIPTION
        A refused plan creates nothing. MsiProductVersion is a script block that takes the manifest and returns
        the ProductVersion of its staged MSI, or $null.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)]$Settings,
        $CommandLine = $null,
        [bool]$PublishesToConfigMgr = $true,
        [bool]$PublishesToIntune = $false,
        [AllowNull()][scriptblock]$MsiProductVersion = $null,
        [AllowNull()][string]$FallbackApplicationId = $null
    )
    # A grid stage of the default profile runs without a snapshot, so its manifest has no application identifier; the run knows it.
    $applicationId = [string](Get-SupersedenceProperty $Manifest 'ApplicationId')
    if ([string]::IsNullOrWhiteSpace($applicationId)) { $applicationId = $FallbackApplicationId }
    $log = New-Object System.Collections.Generic.List[string]
    $refused = { param($text) [pscustomobject]@{ ConfigMgrManifest = $Manifest; IntuneManifest = $Manifest; Marker = $null; Refusal = $text; Log = @($log); Mode = 'Off'; ConfigMgrRule = 'Minimum' } }

    try { $resolved = Resolve-Supersedence -Settings $Settings -ApplicationId $applicationId -CommandLine $CommandLine }
    catch { return (& $refused $_.Exception.Message) }

    $where = { param($source) switch ($source) { 'CommandLine' { 'command line' } 'TitleOverride' { 'title override' } default { 'site default' } } }
    if ($PublishesToConfigMgr) {
        $log.Add(('Version rule (ConfigMgr)    : {0} ({1})' -f $resolved.ConfigMgrRule.Value, (& $where $resolved.ConfigMgrRule.Source)))
        $log.Add(('Supersedence mode           : {0} ({1})' -f (Get-SupersedenceModeText -Mode $resolved.Mode.Value), (& $where $resolved.Mode.Source)))
        if ($resolved.Blocked) { return (& $refused $resolved.Blocked) }
    }
    if ($PublishesToIntune) {
        $log.Add(('Version rule (Intune)       : {0} ({1})' -f $resolved.IntuneRule.Value, (& $where $resolved.IntuneRule.Source)))
    }

    # A title with no identifier has no key: nothing is recorded for it and no block can be accepted.
    $titleKey = Get-SupersedenceOverrideKey -ApplicationId $applicationId
    $recorded = Get-SupersedenceProperty $Settings 'InstalledValues'
    $accepted = @(Get-SupersedenceProperty $Settings 'AcceptedExactBlocks' | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
    $guard = { param($clause) Get-ExactGuardVerdict -Manifest $Manifest -Clause $clause -TitleKey $titleKey -MsiProductVersion $MsiProductVersion -Recorded $recorded -Accepted $accepted }.GetNewClosure()
    $state = @{ AcceptedBlock = $false; Refusal = '' }

    $forTarget = {
        param($target, $rule)
        $applied = Set-ManifestVersionRule -Manifest $Manifest -Rule $rule -Guard $guard
        foreach ($outcome in $applied.Outcomes) {
            switch ($outcome.Result) {
                'ExactApplied' { $log.Add(('Version rule ({0})' -f $target).PadRight(28) + (': Exact written to {0}' -f $outcome.Clause)) }
                'ExactAcceptedBlock' {
                    $state.AcceptedBlock = $true
                    $log.Add(('Version rule ({0})' -f $target).PadRight(28) + (': Exact not applied to {0} (block accepted for this version). {1}' -f $outcome.Clause, $outcome.Detail))
                }
                default {
                    if ($state.Refusal.Length -gt 0) { $state.Refusal += ' ' }
                    $state.Refusal += ('The version rule Exact cannot be applied for {0}: {1}' -f $target, $outcome.Detail)
                }
            }
        }
        return $applied.Manifest
    }.GetNewClosure()

    $configMgrManifest = $Manifest
    $intuneManifest = $Manifest
    if ($PublishesToConfigMgr) { $configMgrManifest = & $forTarget 'ConfigMgr' $resolved.ConfigMgrRule.Value }
    if ($PublishesToIntune) { $intuneManifest = & $forTarget 'Intune' $resolved.IntuneRule.Value }
    if ($state.Refusal.Length -gt 0) {
        return (& $refused ($state.Refusal + ' Set the version rule of this title to Minimum, or accept the block for this version.'))
    }

    # Exact guard rule 5: a title whose Exact was not applied cannot be uninstalled by supersedence.
    if ($PublishesToConfigMgr -and $state.AcceptedBlock -and $resolved.Mode.Value -eq 'ReplaceAndUninstall') {
        return (& $refused 'The version rule Exact was not applied to this title (the block is accepted), so Replace and uninstall is blocked for it. Set the supersedence mode of the title to Replace or Off.')
    }

    $marker = $null
    if ($PublishesToConfigMgr) { $marker = Get-SupersedenceIdentityMarker -Manifest $Manifest -FallbackApplicationId $applicationId }
    return [pscustomobject]@{
        ConfigMgrManifest = $configMgrManifest
        IntuneManifest    = $intuneManifest
        Marker            = $marker
        Refusal           = $null
        Log               = @($log)
        Mode              = $resolved.Mode.Value
        ConfigMgrRule     = $resolved.ConfigMgrRule.Value
    }
}

# ---- recording an installed value and accepting a block ----

function Test-SupersedenceValidValue {
    # True for the text of an installed value: starts with a digit, letters, digits and . + _ - only, at most 64 characters.
    param([AllowNull()][string]$Text)
    return ($null -ne $Text -and $Text -cmatch '^[0-9][0-9A-Za-z.+_-]{0,63}\z')
}

function Test-SupersedenceValidVersionText {
    # True for the version of a title: 1 to 64 characters, no control character, no "|", no space at an end.
    param([AllowNull()][string]$Text)
    if ($null -eq $Text -or $Text.Length -lt 1 -or $Text.Length -gt 64 -or $Text -ne $Text.Trim()) { return $false }
    return ($Text -cnotmatch '[\p{Cc}|]')
}

function Get-SupersedenceStoredKey {
    param([string]$TitleKey, [string]$Version)
    if ([string]::IsNullOrWhiteSpace($TitleKey)) { throw 'The title has no identifier, so nothing can be recorded for it.' }
    if (-not (Test-SupersedenceValidVersionText -Text $Version)) { throw "The version '$Version' is empty, longer than 64 characters, or holds a control character, a space at an end, or '|'." }
    return (Get-SupersedenceRecordKey -TitleKey $TitleKey -Version $Version)
}

function Set-SupersedenceInstalledValue {
    <#
    .SYNOPSIS
        Stores the installed value of a title and a version in the settings; returns the key.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$TitleKey, [Parameter(Mandatory)][string]$Version, [Parameter(Mandatory)][string]$InstalledValue)
    $key = Get-SupersedenceStoredKey -TitleKey $TitleKey -Version $Version
    if (-not (Test-SupersedenceValidValue -Text $InstalledValue)) {
        throw "The installed value '$InstalledValue' is not a version (it starts with a digit and holds letters, digits and . + _ - only)."
    }
    # Rebuild the map so that two keys never differ only in letter case.
    $values = [ordered]@{}
    foreach ($entry in @(Get-SupersedenceMapEntry -Map (Get-SupersedenceProperty $Settings 'InstalledValues'))) {
        $existing = @($values.Keys | Where-Object { $_ -ieq $entry.Key })
        foreach ($k in $existing) { $values.Remove($k) }
        $values[$entry.Key] = $entry.Value
    }
    foreach ($k in @($values.Keys | Where-Object { $_ -ieq $key })) { $values.Remove($k) }
    $values[$key] = $InstalledValue
    $stored = [pscustomobject]$values
    if ($Settings.PSObject.Properties['InstalledValues']) { $Settings.InstalledValues = $stored } else { Add-Member -InputObject $Settings -NotePropertyName InstalledValues -NotePropertyValue $stored }
    return $key
}

function Add-SupersedenceAcceptedBlock {
    # Accepts the block of the Exact guard for a title and a version; the clause is then written as Minimum. Returns the key.
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$TitleKey, [Parameter(Mandatory)][string]$Version)
    $key = Get-SupersedenceStoredKey -TitleKey $TitleKey -Version $Version
    $list = @(Get-SupersedenceProperty $Settings 'AcceptedExactBlocks' | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
    if (-not ($list | Where-Object { $_ -ieq $key })) { $list += $key }
    if ($Settings.PSObject.Properties['AcceptedExactBlocks']) { $Settings.AcceptedExactBlocks = $list } else { Add-Member -InputObject $Settings -NotePropertyName AcceptedExactBlocks -NotePropertyValue $list }
    return $key
}


# ---- the settings as the publish reads them ----

function Get-SupersedenceSettings {
    <#
    .SYNOPSIS
        The Supersedence section of the preferences file, with the shipped values for what is missing.
    .DESCRIPTION
        The preferences file sits beside the Packagers folder. A file that is missing or cannot be read gives
        the shipped settings (Minimum, Minimum, Off). A word that is not allowed stays as written, so the
        resolver refuses the row with the place of the word.
    #>
    [CmdletBinding()]
    param([string]$PreferencesPath = (Join-Path $PSScriptRoot '..\AppPackager.preferences.json'))
    $settings = New-SupersedenceSettings
    try {
        if (-not (Test-Path -LiteralPath $PreferencesPath)) { return $settings }
        $raw = Get-Content -LiteralPath $PreferencesPath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $settings }
        $section = ($raw | ConvertFrom-Json -ErrorAction Stop).Supersedence
    }
    catch { return $settings }
    if ($null -eq $section) { return $settings }
    foreach ($name in 'ConfigMgrVersionRule', 'IntuneVersionRule', 'Mode') {
        $value = Get-SupersedenceProperty $section $name
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { $settings.$name = [string]$value }
    }
    $overrides = Get-SupersedenceProperty $section 'TitleOverrides'
    if ($overrides) { $settings.TitleOverrides = $overrides }
    $values = Get-SupersedenceProperty $section 'InstalledValues'
    if ($values) { $settings.InstalledValues = $values }
    $accepted = Get-SupersedenceProperty $section 'AcceptedExactBlocks'
    if ($null -ne $accepted) { $settings.AcceptedExactBlocks = @($accepted | ForEach-Object { [string]$_ }) }
    return $settings
}

function Get-SupersedenceCommandLine {
    # The version rule and the mode that the command line gave. They outrank the other levels. $null when none was given.
    [CmdletBinding()]
    param()
    $rule = [string]$env:APP_PACKAGER_VERSION_RULE
    $mode = [string]$env:APP_PACKAGER_SUPERSEDENCE
    if ([string]::IsNullOrWhiteSpace($rule) -and [string]::IsNullOrWhiteSpace($mode)) { return $null }
    return [pscustomobject]@{ VersionRule = $rule; Mode = $mode }
}

function Get-SupersedenceMsiProductVersion {
    # The ProductVersion of the staged MSI of a manifest; $null when the installer is not an MSI or cannot be read.
    param([Parameter(Mandatory)]$Manifest, [Parameter(Mandatory)][string]$ContentPath)
    if ([string](Get-SupersedenceProperty $Manifest 'InstallerType') -ne 'MSI') { return $null }
    try {
        $path = Join-Path $ContentPath ([string](Get-SupersedenceProperty $Manifest 'InstallerFile'))
        $map = Get-MsiPropertyMap -MsiPath $path
        $version = [string]$map['ProductVersion']
        if ([string]::IsNullOrWhiteSpace($version)) { return $null }
        return $version
    }
    catch { return $null }
}

# ---- the lookup of the old application and the pairing of deployment types ----
# A line application is [pscustomobject]@{ CiId; Title; SoftwareVersion; IsExpired; DeploymentTypes = @([pscustomobject]@{ CiId; Name }) }.

function ConvertTo-SupersedenceVersionText {
    # A version string that [version] can parse, or the input unchanged when it is not one. A build or pre-release
    # suffix after + or -, a trailing channel word, and a fifth number are dropped; a bare number reads as N.0.
    param([AllowNull()][string]$Value)
    $text = [regex]::Replace(([string]$Value).Trim(), '[+-].*$', '')
    $text = [regex]::Replace($text, '[A-Za-z]+$', '')
    $text = [regex]::Replace($text, '^((\d+\.){3}\d+)\..*$', '$1')
    return [regex]::Replace($text, '^(\d+)$', '$1.0')
}

function ConvertTo-SupersedenceVersion {
    # The [version] of a text, or $null when it cannot be read.
    param([AllowNull()][string]$Value)
    $parsed = $null
    if ([version]::TryParse((ConvertTo-SupersedenceVersionText -Value $Value), [ref]$parsed)) { return $parsed }
    return $null
}

function Compare-SupersedenceVersion {
    # -1, 0 or 1 over all four parts of both sides. Either side unparsable compares as equal.
    param([AllowNull()][string]$A, [AllowNull()][string]$B)
    $va = ConvertTo-SupersedenceVersion -Value $A
    $vb = ConvertTo-SupersedenceVersion -Value $B
    if ($null -eq $va -or $null -eq $vb) { return 0 }
    $partsA = @($va.Major, $va.Minor, [Math]::Max($va.Build, 0), [Math]::Max($va.Revision, 0))
    $partsB = @($vb.Major, $vb.Minor, [Math]::Max($vb.Build, 0), [Math]::Max($vb.Revision, 0))
    for ($i = 0; $i -lt 4; $i++) {
        if ($partsA[$i] -lt $partsB[$i]) { return -1 }
        if ($partsA[$i] -gt $partsB[$i]) { return 1 }
    }
    return 0
}

function Select-SupersedenceOldApplication {
    <#
    .SYNOPSIS
        The newest application of the line whose title is not the new title, that is not retired and whose
        version is lower than the packaged version.
    .DESCRIPTION
        A version that cannot be read is skipped and named. Returns Old (or $null), Notes and Summary.
    #>
    param([AllowEmptyCollection()][object[]]$Applications = @(), [Parameter(Mandatory)][string]$NewTitle, [Parameter(Mandatory)][string]$NewVersion)
    $notes = New-Object System.Collections.Generic.List[string]
    $candidates = @()
    foreach ($app in @($Applications)) {
        if ([string]::Equals([string]$app.Title, $NewTitle, [StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($app.IsExpired) { $notes.Add("'$($app.Title)' is retired and was skipped."); continue }
        if ($null -eq (ConvertTo-SupersedenceVersion -Value $app.SoftwareVersion)) {
            $notes.Add("The version '$($app.SoftwareVersion)' of '$($app.Title)' cannot be read, so it was skipped."); continue
        }
        if ((Compare-SupersedenceVersion -A $app.SoftwareVersion -B $NewVersion) -ge 0) {
            # A line that already holds a higher or an equal version under another title: the new application is not newer than it.
            $notes.Add("'$($app.Title)' $($app.SoftwareVersion) is not older than $NewVersion and is not superseded.")
            continue
        }
        $candidates += $app
    }
    $result = { param($old, $summaryNotes)
        $summary = if ($null -eq $old) { 'No older version found' + $(if ($summaryNotes.Count -gt 0) { ': ' + ($summaryNotes -join ' ') } else { '' }) } else { "Old application: $($old.Title) $($old.SoftwareVersion)" }
        [pscustomobject]@{ Old = $old; Notes = @($summaryNotes); Summary = $summary }
    }
    if ($null -eq (ConvertTo-SupersedenceVersion -Value $NewVersion)) { return (& $result $null ([string[]]@("The packaged version '$NewVersion' cannot be read."))) }
    if ($candidates.Count -eq 0) { return (& $result $null $notes) }
    $ordered = @($candidates | Sort-Object -Descending -Property @{ Expression = { ConvertTo-SupersedenceVersion -Value $_.SoftwareVersion } })
    foreach ($other in ($ordered | Select-Object -Skip 1)) { $notes.Add("'$($other.Title)' $($other.SoftwareVersion) also carries the marker and is not superseded by this publish.") }
    return (& $result $ordered[0] $notes)
}

function Get-SupersedenceDeploymentTypePair {
    <#
    .SYNOPSIS
        Pairs the deployment types of the new application with those of the old one.
    .DESCRIPTION
        One and one pair directly. With more than one, the pair is made by the suffix of the name that staging writes
        ('<Title> - <Suffix>'). A type with no partner gets a warning and no relation. One architecture is never linked to another.
    #>
    param([Parameter(Mandatory)]$NewApplication, [Parameter(Mandatory)]$OldApplication)
    $pairs = @()
    $warnings = New-Object System.Collections.Generic.List[string]
    $news = @($NewApplication.DeploymentTypes)
    $olds = @($OldApplication.DeploymentTypes)
    if ($olds.Count -eq 0) {
        $warnings.Add("'$($OldApplication.Title)' has no deployment type, so nothing is superseded.")
        return [pscustomobject]@{ Pairs = @(); Warnings = @($warnings) }
    }
    if ($news.Count -eq 1 -and $olds.Count -eq 1) {
        return [pscustomobject]@{ Pairs = @([pscustomobject]@{ New = $news[0]; Old = $olds[0] }); Warnings = @() }
    }
    $suffix = { param($name, $title) if ($name.StartsWith($title + ' - ', [StringComparison]::OrdinalIgnoreCase)) { $name.Substring($title.Length + 3).Trim() } else { $name.Trim() } }
    $unmatched = New-Object System.Collections.ArrayList
    foreach ($o in $olds) { [void]$unmatched.Add($o) }
    foreach ($dt in $news) {
        $wanted = & $suffix ([string]$dt.Name) ([string]$NewApplication.Title)
        $partner = $null
        foreach ($o in $unmatched) { if ([string]::Equals((& $suffix ([string]$o.Name) ([string]$OldApplication.Title)), $wanted, [StringComparison]::OrdinalIgnoreCase)) { $partner = $o; break } }
        if ($null -eq $partner) { $warnings.Add("The deployment type '$($dt.Name)' has no partner in '$($OldApplication.Title)' and supersedes nothing."); continue }
        $unmatched.Remove($partner)
        $pairs += [pscustomobject]@{ New = $dt; Old = $partner }
    }
    foreach ($rest in $unmatched) { $warnings.Add("The deployment type '$($rest.Name)' of '$($OldApplication.Title)' has no partner and is not superseded.") }
    return [pscustomobject]@{ Pairs = @($pairs); Warnings = @($warnings) }
}

function Get-SupersedenceLineApplication {
    <#
    .SYNOPSIS
        Reads the applications of the site whose stored XML holds the identity marker, with their deployment types.
    .DESCRIPTION
        The new application is always listed, with its deployment types, even when it holds no marker. It is found
        by its name: every save makes a new revision with a new CI_ID. Needs the ConfigMgr site drive.
    #>
    param([Parameter(Mandatory)][string]$Marker, [AllowEmptyString()][string]$NewApplicationName = '')
    $needle = '>' + $Marker.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') + '<'
    $found = @()
    foreach ($candidate in @(Get-CMApplication -ErrorAction Stop)) {
        $xml = [string]$candidate.SDMPackageXML
        if ([string]$candidate.LocalizedDisplayName -ne $NewApplicationName -and (-not $xml -or -not $xml.Contains($needle))) { continue }
        $types = @()
        foreach ($dt in @(Get-CMDeploymentType -ApplicationName ([string]$candidate.LocalizedDisplayName) -DisableWildcardHandling -ErrorAction Stop)) {
            $types += [pscustomobject]@{ CiId = [string]$dt.CI_ID; Name = [string]$dt.LocalizedDisplayName }
        }
        $found += [pscustomobject]@{
            CiId = [string]$candidate.CI_ID; Title = [string]$candidate.LocalizedDisplayName; SoftwareVersion = [string]$candidate.SoftwareVersion
            IsExpired = [bool]$candidate.IsExpired; DeploymentTypes = $types
        }
    }
    return $found
}

# ---- the supersedence step ----

function Test-SupersedenceRelationExists {
    # True when the deployment type of the new application already supersedes the old deployment type.
    param([string]$ApplicationName, [string]$DeploymentTypeName, [string]$OldDeploymentTypeName)
    try {
        $dt = Get-CMDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName -DisableWildcardHandling -ErrorAction Stop
        foreach ($relation in @(Get-CMDeploymentTypeSupersedence -InputObject $dt -ErrorAction Stop)) {
            if ((($relation | Out-String)).Contains($OldDeploymentTypeName)) { return $true }
        }
    }
    catch { $null = $_ }
    return $false
}

function Invoke-SupersedenceStep {
    <#
    .SYNOPSIS
        The supersedence step after a publish: the lookup of the older application of the line (by the identity
        marker) and one relation for each pair of deployment types.
    .DESCRIPTION
        A failure never undoes the application: the result says NotApplied and why. Status is NotRun, NoOlderVersion,
        Applied or NotApplied. The relations are read first, and one that exists is not written again. Names and not
        identifiers go to the cmdlet: every save of an application makes a new revision with a new CI_ID. The
        uninstall option of an existing relation is not changed. Needs the ConfigMgr site drive.
    #>
    param(
        [Parameter(Mandatory)][string]$NewApplicationName,
        [Parameter(Mandatory)][string]$NewVersion,
        [AllowNull()][string]$Marker,
        [Parameter(Mandatory)][string]$Mode,
        [switch]$RequirePublished
    )
    $result = { param($status, $message, $warnings = @()) [pscustomobject]@{ Status = $status; Message = $message; Warnings = @($warnings); Failure = $(if ($status -eq 'NotApplied') { $message } else { $null }) } }
    if ($Mode -eq 'Off' -or [string]::IsNullOrWhiteSpace($Marker)) { return (& $result 'NotRun' '') }
    try {
        try { $applications = @(Get-SupersedenceLineApplication -Marker $Marker -NewApplicationName $NewApplicationName) }
        catch { return (& $result 'NotApplied' "the lookup of the older version failed: $($_.Exception.Message)") }
        if ($RequirePublished) {
            # The action without a publish: the application of this title and version must be on the site, or nothing is applied.
            $published = $applications | Where-Object { [string]::Equals($_.Title, $NewApplicationName, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
            if (-not $published) { return (& $result 'NotApplied' "The application '$NewApplicationName' is not on the site. Publish it first.") }
            if (-not (Test-SupersedenceSameVersionText -A $published.SoftwareVersion -B $NewVersion)) {
                return (& $result 'NotApplied' "The site holds '$NewApplicationName' at version $($published.SoftwareVersion), not $NewVersion. Publish this version first.")
            }
        }
        $chosen = Select-SupersedenceOldApplication -Applications $applications -NewTitle $NewApplicationName -NewVersion $NewVersion
        if ($null -eq $chosen.Old) { return (& $result 'NoOlderVersion' $chosen.Summary $chosen.Notes) }
        $current = $applications | Where-Object { [string]::Equals($_.Title, $NewApplicationName, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
        if (-not $current) { return (& $result 'NotApplied' 'the new application was not found on the site after it was created.' $chosen.Notes) }
        $pairing = Get-SupersedenceDeploymentTypePair -NewApplication $current -OldApplication $chosen.Old
        $warnings = @($chosen.Notes) + @($pairing.Warnings)
        if (@($pairing.Pairs).Count -eq 0) { return (& $result 'NoOlderVersion' "No deployment types to pair with $($chosen.Old.Title)." $warnings) }

        $uninstall = ($Mode -eq 'ReplaceAndUninstall')
        $failures = @()
        foreach ($pair in $pairing.Pairs) {
            $label = '{0} supersedes {1}' -f $pair.New.Name, $pair.Old.Name
            if (Test-SupersedenceRelationExists -ApplicationName $current.Title -DeploymentTypeName $pair.New.Name -OldDeploymentTypeName $pair.Old.Name) {
                Write-Log "Supersedence already exists  : $label"
                continue
            }
            try {
                Set-CMApplicationSupersedence -Name $current.Title -CurrentDeploymentTypeName $pair.New.Name -SupersededApplicationName $chosen.Old.Title -OldDeploymentTypeName $pair.Old.Name -IsUninstall:$uninstall -DisableWildcardHandling -ErrorAction Stop | Out-Null
                Write-Log "Supersedence added           : $label"
            }
            catch {
                $writeError = $_.Exception.Message
                # A relation that is there after a failed write counts as done; no text of the provider decides this.
                if (Test-SupersedenceRelationExists -ApplicationName $current.Title -DeploymentTypeName $pair.New.Name -OldDeploymentTypeName $pair.Old.Name) {
                    Write-Log "Supersedence already exists  : $label"
                }
                else { $failures += ('{0}: {1}' -f $label, $writeError) }
            }
        }
        if ($failures.Count -gt 0) { return (& $result 'NotApplied' $failures[0] $warnings) }
        $text = "$($current.Title) supersedes $($chosen.Old.Title) $($chosen.Old.SoftwareVersion)"
        if ($uninstall) { $text += ' and uninstalls it' }
        return (& $result 'Applied' $text $warnings)
    }
    catch {
        # The application is complete; whatever goes wrong here is reported and never fails the run.
        return (& $result 'NotApplied' $_.Exception.Message)
    }
}

function Get-SupersedenceFallbackApplicationId {
    # A manifest that was staged without a run snapshot has no application identifier; the packager script that is running
    # gives it: catalog:<script name>. $null when no packager script is on the call stack.
    $frame = @(Get-PSCallStack | Where-Object { $_.ScriptName -and $_.ScriptName -match '[\\/]package-[^\\/]+\.ps1$' }) | Select-Object -First 1
    if (-not $frame) { return $null }
    return ('catalog:' + [System.IO.Path]::GetFileNameWithoutExtension([string]$frame.ScriptName))
}

function Invoke-SupersedenceStepAndLog {
    # Runs the supersedence step and writes its result to the log. A failure is a warning, never an error.
    param([Parameter(Mandatory)][string]$ApplicationName, [Parameter(Mandatory)][string]$Version, [Parameter(Mandatory)]$Plan)
    $step = Invoke-SupersedenceStep -NewApplicationName $ApplicationName -NewVersion $Version -Marker $Plan.Marker -Mode $Plan.Mode
    if ($step.Status -eq 'NotApplied') { Write-Log ('Published, supersedence not applied: {0}' -f $step.Message) -Level WARN }
    elseif ($step.Status -ne 'NotRun') { Write-Log ('Supersedence                 : {0}' -f $step.Message) }
    foreach ($warning in $step.Warnings) { Write-Log ('Supersedence warning         : {0}' -f $warning) -Level WARN }
    return $step
}

function Invoke-SupersedenceApply {
    <#
    .SYNOPSIS
        Applies the supersedence of one staged application to the ConfigMgr site, without a publish. This is the
        engine entry point of the Apply Supersedence action.
    .DESCRIPTION
        Reads the stage manifest, resolves the title as the publish resolves it, plans the supersedence with the
        same planner as a publish, and runs the supersedence step for the application that is already on the site.
        The step reads the relations first, so a repeat adds only the missing ones. Every refusal comes before the
        first call to the site. Returns Ok, Message, Status (NotRun, NoOlderVersion, Applied or NotApplied),
        RowStatus (the text for the grid), Title and Log (the lines for the log). Never throws. Settings and
        CommandLine default to the preferences file and the command line variables of the process.
    #>
    param(
        [AllowEmptyString()][string]$ManifestPath = '',
        [AllowEmptyString()][string]$SiteCode = '',
        [AllowEmptyString()][string]$ProviderMachineName = '',
        [AllowNull()][string]$FallbackApplicationId = $null,
        [AllowEmptyString()][string]$TitleMode = '',
        $Settings = $null,
        $CommandLine = $null
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $outcome = {
        param([bool]$ok, [string]$message, [string]$status, [string]$title = '')
        $rowStatus = if (-not $ok) { 'Supersedence not applied' } elseif ($status -eq 'Applied') { 'Supersedence applied' } else { 'No older version' }
        [pscustomobject]@{ Ok = $ok; Message = $message; Status = $status; RowStatus = $rowStatus; Title = $title; Log = @($lines) }
    }
    try {
        if ([string]::IsNullOrWhiteSpace($ManifestPath) -or -not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
            return (& $outcome $false 'No stage manifest was found for this application. Stage and publish it first.' 'NotApplied')
        }
        # The title mode reaches Read-StageManifest and Get-PackagedApplicationName through the environment, as it does in a publish.
        $savedTitleMode = $env:APP_PACKAGER_TITLE_MODE
        try {
            $env:APP_PACKAGER_TITLE_MODE = $TitleMode
            try { $manifest = Read-StageManifest -Path $ManifestPath }
            catch { return (& $outcome $false ('The stage manifest {0} could not be read: {1}' -f $ManifestPath, $_.Exception.Message) 'NotApplied') }
            if ([string]::IsNullOrWhiteSpace([string]$manifest.AppName)) {
                return (& $outcome $false 'The stage manifest has no application name. Stage the application again.' 'NotApplied')
            }
            try { $title = Get-PackagedApplicationName -AppName ([string]$manifest.AppName) -Version ([string]$manifest.SoftwareVersion) }
            catch { return (& $outcome $false ('The application title could not be resolved: {0}' -f $_.Exception.Message) 'NotApplied') }
        }
        finally {
            if ($null -eq $savedTitleMode) { Remove-Item -LiteralPath 'Env:\APP_PACKAGER_TITLE_MODE' -ErrorAction SilentlyContinue }
            else { $env:APP_PACKAGER_TITLE_MODE = $savedTitleMode }
        }
        if ([string]::IsNullOrWhiteSpace($SiteCode)) {
            return (& $outcome $false 'Supersedence needs a ConfigMgr site. Set the site code in Options.' 'NotApplied' $title)
        }

        if ($null -eq $Settings) { $Settings = Get-SupersedenceSettings }
        if ($null -eq $CommandLine) { $CommandLine = Get-SupersedenceCommandLine }
        $contentPath = Split-Path -Parent $ManifestPath
        $plan = Get-SupersedencePlan -Manifest $manifest -Settings $Settings -CommandLine $CommandLine `
            -PublishesToConfigMgr $true -PublishesToIntune $false `
            -MsiProductVersion { param($m) Get-SupersedenceMsiProductVersion -Manifest $m -ContentPath $contentPath }.GetNewClosure() `
            -FallbackApplicationId $FallbackApplicationId
        foreach ($planLine in $plan.Log) { $lines.Add([string]$planLine) }
        if ($plan.Refusal) { return (& $outcome $false ([string]$plan.Refusal) 'NotApplied' $title) }
        if ($plan.Mode -eq 'Off') {
            return (& $outcome $false 'Supersedence is off for this title. Turn it on in Options or in the override of the title.' 'NotApplied' $title)
        }
        if ([string]::IsNullOrWhiteSpace([string]$plan.Marker)) {
            return (& $outcome $false 'This application has no identity marker, so its line cannot be found.' 'NotApplied' $title)
        }

        $savedLocation = Get-Location
        try {
            if (-not (Connect-CMSite -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName)) {
                return (& $outcome $false ('The connection to the site {0} failed. Check the site code and the site server in Options, Connections, ConfigMgr.' -f $SiteCode) 'NotApplied' $title)
            }
            $step = Invoke-SupersedenceStep -NewApplicationName $title -NewVersion ([string]$manifest.SoftwareVersion) -Marker ([string]$plan.Marker) -Mode ([string]$plan.Mode) -RequirePublished
        }
        finally {
            # The site drive is left behind, so a later local path would resolve through the CM provider.
            try { Set-Location -LiteralPath $savedLocation.Path -ErrorAction Stop } catch { $null = $_ }
        }
        foreach ($warning in $step.Warnings) { $lines.Add(('Supersedence warning         : {0}' -f $warning)) }
        return (& $outcome ($step.Status -ne 'NotApplied') ([string]$step.Message) ([string]$step.Status) $title)
    }
    catch {
        return (& $outcome $false $_.Exception.Message 'NotApplied')
    }
}

# ---- the identity marker on the site ----

function Set-CMApplicationIdentityMarker {
    <#
    .SYNOPSIS
        Writes the identity marker into the keywords of the application.
    .DESCRIPTION
        A keyword of an administrator is never rewritten: when the application has keywords of its own, or when the
        keywords cannot be read, nothing is written and the function returns the reason. An earlier marker is
        replaced. Returns $null on success, or the reason text. Needs the ConfigMgr site drive.
    #>
    param([Parameter(Mandatory)][string]$ApplicationName, [Parameter(Mandatory)][string]$Marker)
    try {
        $app = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
        $tags = @()
        $digest = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($app.SDMPackageXML, $true)
        foreach ($info in $digest.DisplayInfo) { foreach ($tag in $info.Tags) { $tags += [string]$tag } }
        # The cmdlet takes the keyword as one value, so the marker is written only when the application has no keywords of its own.
        $others = @($tags | Where-Object { $_ -and -not $_.StartsWith($script:SupersedenceMarkerPrefix) })
        if ($others.Count -gt 0) { throw 'the application has keywords of its own, which a write would replace' }
        Set-CMApplication -InputObject $app -Keyword $Marker -ErrorAction Stop
        Write-Log "Identity marker              : $Marker"
        return $null
    }
    catch {
        Write-Log "Identity marker not written  : $($_.Exception.Message)" -Level WARN
        return $_.Exception.Message
    }
}

# ---- the older-application override ("Use this as the old application") ----
# The lookup of the old application takes only applications that carry the identity marker of the line. An application that an
# earlier release published, or one that an administrator made by hand, has no marker, so the lookup never chooses it. The
# administrator links it once: the program shows the title and the version, asks, and then writes the marker of the line
# into the keywords of that application. The link is the marker on the site; nothing is kept in the preferences file, so the next
# publish of the line, from this PC or another, finds the application through the same lookup.

function Get-SupersedenceApplicationKeyword {
    # The keywords that the application holds, read from its stored XML. Needs the ConfigMgr console assemblies.
    param([Parameter(Mandatory)]$Application)
    $tags = @()
    $digest = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString([string]$Application.SDMPackageXML, $true)
    foreach ($info in $digest.DisplayInfo) { foreach ($tag in $info.Tags) { $tags += [string]$tag } }
    return @($tags | Where-Object { $_ })
}

function Get-SupersedenceOldApplicationCandidate {
    <#
    .SYNOPSIS
        The applications of the site that can be linked as the old application of a line, with the reason for each one that cannot.
    .DESCRIPTION
        Applications is a list of objects with Title, SoftwareVersion, IsExpired and Keywords ($null when the keywords could
        not be read). An application with the title of the new one is never listed: the lookup never returns it. State is
        Linkable, Linked (it already holds the marker of this line), OtherLine (it holds the marker of another line),
        OwnKeywords (a write would replace the keywords of an administrator), KeywordsUnreadable, Retired, VersionUnreadable
        or NotOlder (the lookup would not choose it). Only Linkable can be linked. CanClear is true for an application whose
        keywords are the marker of this line and nothing else. The linkable ones come first, the newest version first.
    #>
    param(
        [AllowEmptyCollection()][object[]]$Applications = @(),
        [Parameter(Mandatory)][string]$Marker,
        [Parameter(Mandatory)][string]$NewTitle,
        [Parameter(Mandatory)][string]$NewVersion
    )
    $row = { param($app, $state, $reason, $canLink, $canClear)
        [pscustomobject]@{ Title = [string]$app.Title; SoftwareVersion = [string]$app.SoftwareVersion; State = $state; Reason = $reason; CanLink = $canLink; CanClear = $canClear }
    }
    $rows = @()
    foreach ($app in @($Applications)) {
        if ([string]::Equals([string]$app.Title, $NewTitle, [StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($null -eq $app.Keywords) { $rows += (& $row $app 'KeywordsUnreadable' 'The keywords could not be read, so nothing is written.' $false $false); continue }
        $keywords = @($app.Keywords | ForEach-Object { [string]$_ } | Where-Object { $_ })
        $mine = @($keywords | Where-Object { $_ -ceq $Marker })
        $markers = @($keywords | Where-Object { $_.StartsWith($script:SupersedenceMarkerPrefix) })
        $own = @($keywords | Where-Object { -not $_.StartsWith($script:SupersedenceMarkerPrefix) })
        if ($mine.Count -gt 0) {
            $clear = ($own.Count -eq 0 -and $markers.Count -eq 1)
            $rows += (& $row $app 'Linked' 'It already holds the marker of this line.' $false $clear); continue
        }
        if ($markers.Count -gt 0) { $rows += (& $row $app 'OtherLine' 'It holds the marker of another line.' $false $false); continue }
        if ($own.Count -gt 0) { $rows += (& $row $app 'OwnKeywords' 'It has keywords of its own, which a write would replace.' $false $false); continue }
        if ($app.IsExpired) { $rows += (& $row $app 'Retired' 'It is retired.' $false $false); continue }
        if ($null -eq (ConvertTo-SupersedenceVersion -Value $app.SoftwareVersion)) {
            $rows += (& $row $app 'VersionUnreadable' ("The version '{0}' cannot be read." -f $app.SoftwareVersion) $false $false); continue
        }
        if ($null -eq (ConvertTo-SupersedenceVersion -Value $NewVersion) -or (Compare-SupersedenceVersion -A $app.SoftwareVersion -B $NewVersion) -ge 0) {
            $rows += (& $row $app 'NotOlder' ("It is not older than {0}, so the lookup would not choose it." -f $NewVersion) $false $false); continue
        }
        $rows += (& $row $app 'Linkable' '' $true $false)
    }
    $linkable = @($rows | Where-Object { $_.CanLink } | Sort-Object -Descending -Property @{ Expression = { ConvertTo-SupersedenceVersion -Value $_.SoftwareVersion } })
    $others = @($rows | Where-Object { -not $_.CanLink } | Sort-Object -Property Title)
    return @($linkable) + @($others)
}

function Get-SupersedenceOldApplicationCandidateFromSite {
    <#
    .SYNOPSIS
        Reads the applications of the site whose name contains the text, and classifies them for the line.
    .DESCRIPTION
        The text limits the read: a site holds many applications, and a list of all of them is not a choice. The
        wildcard characters of the text are literal. Needs the ConfigMgr site drive.
    #>
    param(
        [Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$NewTitle, [Parameter(Mandatory)][string]$NewVersion,
        [Parameter(Mandatory)][string]$NameContains
    )
    if ([string]::IsNullOrWhiteSpace($NameContains)) { throw 'Type part of the application name to search for.' }
    $records = @()
    foreach ($app in @(Get-CMApplication -Name ('*' + [System.Management.Automation.WildcardPattern]::Escape($NameContains.Trim()) + '*') -ErrorAction Stop)) {
        $keywords = $null
        try { $keywords = @(Get-SupersedenceApplicationKeyword -Application $app) } catch { $keywords = $null }
        $records += [pscustomobject]@{ Title = [string]$app.LocalizedDisplayName; SoftwareVersion = [string]$app.SoftwareVersion; IsExpired = [bool]$app.IsExpired; Keywords = $keywords }
    }
    return @(Get-SupersedenceOldApplicationCandidate -Applications $records -Marker $Marker -NewTitle $NewTitle -NewVersion $NewVersion)
}

function Get-SupersedenceLinkConfirmationText {
    # What the administrator reads before the marker is written: the title and the version of the application first.
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Version, [Parameter(Mandatory)][string]$NewTitle)
    return ("Use '{0}' version {1} as the old application of '{2}'?`n`nThe program writes the identity marker of this product line into the keywords of '{0}'. The next publish of '{2}' with a supersedence mode other than Off then finds it and supersedes it. Its deployment types, detection and deployments are not changed." -f $Title, $Version, $NewTitle)
}

function Get-SupersedenceClearConfirmationText {
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Version)
    return ("Remove the identity marker from '{0}' version {1}?`n`nThe next publish of this line no longer finds it as an older application, so nothing supersedes it. If a publish of this program marked it, link it again with Use as Old Application when you want it back." -f $Title, $Version)
}

function Set-SupersedenceOldApplication {
    <#
    .SYNOPSIS
        Links an application as the old application of a line: writes the marker of the line into its keywords.
    .DESCRIPTION
        Nothing is written without -Confirmed. The keywords are read again first: an application that holds keywords of
        its own, or the marker of another line, is refused, and one that already holds the marker needs no write.
        Returns Ok and Message. Needs the ConfigMgr site drive.
    #>
    param([Parameter(Mandatory)][string]$ApplicationName, [Parameter(Mandatory)][string]$Marker, [switch]$Confirmed)
    $outcome = { param($ok, $message) [pscustomobject]@{ Ok = $ok; Message = $message } }
    if (-not $Confirmed) { return (& $outcome $false 'The link is written only after the administrator confirms it.') }
    try {
        $app = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
        if (-not $app) { return (& $outcome $false "The application '$ApplicationName' is not on the site.") }
        $keywords = @(Get-SupersedenceApplicationKeyword -Application $app)
    }
    catch { return (& $outcome $false ("The keywords of '{0}' could not be read: {1}" -f $ApplicationName, $_.Exception.Message)) }
    if (@($keywords | Where-Object { $_ -ceq $Marker }).Count -gt 0) { return (& $outcome $true "'$ApplicationName' already holds the marker of this line.") }
    if (@($keywords | Where-Object { $_.StartsWith($script:SupersedenceMarkerPrefix) }).Count -gt 0) { return (& $outcome $false "'$ApplicationName' holds the marker of another line, so it is not changed.") }
    if ($keywords.Count -gt 0) { return (& $outcome $false "'$ApplicationName' has keywords of its own, which a write would replace, so it is not changed.") }
    $reason = Set-CMApplicationIdentityMarker -ApplicationName $ApplicationName -Marker $Marker
    if ($reason) { return (& $outcome $false ("The marker was not written: {0}" -f $reason)) }
    return (& $outcome $true "'$ApplicationName' is the old application of this line.")
}

function Clear-SupersedenceOldApplication {
    <#
    .SYNOPSIS
        Takes the marker of a line off an application, so the lookup no longer finds it.
    .DESCRIPTION
        Nothing is written without -Confirmed. Only an application whose keywords are the marker of this line and nothing else
        is changed. Returns Ok and Message. Needs the ConfigMgr site drive.
    #>
    param([Parameter(Mandatory)][string]$ApplicationName, [Parameter(Mandatory)][string]$Marker, [switch]$Confirmed)
    $outcome = { param($ok, $message) [pscustomobject]@{ Ok = $ok; Message = $message } }
    if (-not $Confirmed) { return (& $outcome $false 'The marker is removed only after the administrator confirms it.') }
    try {
        $app = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
        if (-not $app) { return (& $outcome $false "The application '$ApplicationName' is not on the site.") }
        $keywords = @(Get-SupersedenceApplicationKeyword -Application $app)
        if ($keywords.Count -ne 1 -or $keywords[0] -cne $Marker) { return (& $outcome $false "The keywords of '$ApplicationName' are not the marker of this line alone, so they are not changed.") }
        Set-CMApplication -InputObject $app -Keyword '' -ErrorAction Stop
        Write-Log "Identity marker removed      : $ApplicationName"
        return (& $outcome $true "The marker is removed from '$ApplicationName'.")
    }
    catch { return (& $outcome $false ("The marker was not removed: {0}" -f $_.Exception.Message)) }
}

# ---- the chain check ----

function Test-SupersedenceMinimumOperator {
    param([string]$Operator)
    return ($Operator -in @('GreaterEquals', 'GreaterThanOrEquals', 'GreaterThanOrEqual'))
}

function Test-SupersedenceExactOperator {
    param([string]$Operator)
    return ($Operator -in @('Equals', 'IsEquals'))
}

function Get-SupersedenceXmlOperator {
    # The operators of the version expressions in one XML text; reads the XML text inside an argument when the elements hold none.
    param([AllowNull()][string]$Xml, [int]$Depth = 0)
    $operators = @()
    if ([string]::IsNullOrWhiteSpace($Xml) -or $Depth -gt 2) { return $operators }
    try { $document = New-Object System.Xml.XmlDocument; $document.LoadXml($Xml) }
    catch [System.Xml.XmlException] { return $operators }
    # The site stores the detection twice (checked on ConfigMgr 5.00.9146): as rule elements of the deployment type, and as XML text
    # inside an argument of the installer. The elements come first; the text is read as a document of its own only when the
    # elements hold no version clause, so one clause is never counted twice.
    foreach ($expression in $document.SelectNodes('//*[local-name()="Expression"]')) {
        $operatorNode = $expression.SelectSingleNode('*[local-name()="Operator"]')
        if ($null -eq $operatorNode -or [string]::IsNullOrWhiteSpace($operatorNode.InnerText)) { continue }
        $isVersion = $false
        foreach ($operand in $expression.SelectNodes('*[local-name()="Operands"]//*[local-name()="SettingReference" or local-name()="ConstantValue"]')) {
            if ($operand.GetAttribute('DataType') -ieq 'Version') { $isVersion = $true }
        }
        if ($isVersion) { $operators += $operatorNode.InnerText.Trim() }
    }
    if ($operators.Count -eq 0) {
        $seen = @{}
        foreach ($holder in $document.SelectNodes('//*[not(*)][contains(., "<EnhancedDetectionMethod")]')) {
            if ($seen.ContainsKey($holder.InnerText)) { continue }
            $seen[$holder.InnerText] = $true
            $operators += @(Get-SupersedenceXmlOperator -Xml $holder.InnerText -Depth ($Depth + 1))
        }
    }
    return $operators
}

function Get-DetectionVersionOperator {
    # The operators of the version clauses of one deployment type, read from its stored XML: every expression whose
    # operand is a setting reference or a constant of the data type Version. A document that cannot be read gives no clause.
    param([AllowNull()][string]$DeploymentTypeXml)
    return @(Get-SupersedenceXmlOperator -Xml $DeploymentTypeXml)
}

function Get-SupersedenceMinimumClause {
    # The version clauses of the new detection that still say "this version or newer" after the rule was written.
    param([Parameter(Mandatory)]$Manifest)
    $found = New-Object System.Collections.Generic.List[string]
    $walk = $null
    $walk = {
        param($d)
        if ($null -eq $d) { return }
        $type = [string](Get-SupersedenceProperty $d 'Type')
        if ($type -eq 'Compound') { foreach ($c in @(Get-SupersedenceProperty $d 'Clauses')) { & $walk $c }; return }
        $propertyType = [string](Get-SupersedenceProperty $d 'PropertyType')
        if ($type -eq 'File') { $isVersion = ([string]::IsNullOrWhiteSpace($propertyType) -or $propertyType -eq 'Version') }
        else { $isVersion = ($type -eq 'RegistryKeyValue' -and ($propertyType -eq 'Version' -or ([string](Get-SupersedenceProperty $d 'ValueName')).IndexOf('Version', [StringComparison]::OrdinalIgnoreCase) -ge 0)) }
        if (-not $isVersion) { return }
        $op = [string](Get-SupersedenceProperty $d 'Operator')
        if ([string]::IsNullOrWhiteSpace($op)) { $op = $(if ($type -eq 'File') { 'GreaterEquals' } else { 'IsEquals' }) }
        if ($op -eq 'GreaterEquals') { $found.Add((Get-SupersedenceClauseDescription -Clause $d)) }
    }
    & $walk (Get-SupersedenceProperty $Manifest 'Detection')
    foreach ($type in @(Get-SupersedenceProperty $Manifest 'DeploymentTypes')) {
        $own = Get-SupersedenceProperty $type 'Detection'
        if ($own) { & $walk $own }
    }
    return @($found)
}

function Get-SupersedenceChainWarning {
    <#
    .SYNOPSIS
        The chain warnings: the detection of the old application against the new one. They warn and never block.
    .DESCRIPTION
        OldDetection is [pscustomobject]@{ Title; Clauses = @([pscustomobject]@{ DeploymentType; Operator }) }.
    #>
    param(
        [Parameter(Mandatory)][string]$NewTitle,
        [Parameter(Mandatory)][ValidateSet('Minimum', 'Exact')][string]$NewRule,
        [Parameter(Mandatory)][string]$Mode,
        [Parameter(Mandatory)]$NewManifest,
        [Parameter(Mandatory)]$OldDetection
    )
    $warnings = @()
    $oldMinimum = @($OldDetection.Clauses | Where-Object { Test-SupersedenceMinimumOperator -Operator $_.Operator })
    $oldExact = @($OldDetection.Clauses | Where-Object { Test-SupersedenceExactOperator -Operator $_.Operator })
    if ($Mode -eq 'ReplaceAndUninstall' -and $NewRule -eq 'Exact') {
        foreach ($clause in $oldMinimum) {
            $warnings += "'$($OldDetection.Title)' ($($clause.DeploymentType)) has a version clause with GreaterEquals, so its uninstall will still be reported as failed after the uninstall. Rewrite its clause to Exact first."
        }
    }
    if ($NewRule -eq 'Minimum' -and $oldExact.Count -gt 0) {
        $warnings += "'$($OldDetection.Title)' has an Exact version clause and '$NewTitle' is published with Minimum, so the chain is inconsistent."
    }
    if ($NewRule -eq 'Exact' -and $oldMinimum.Count -gt 0 -and $Mode -ne 'ReplaceAndUninstall') {
        $warnings += "'$($OldDetection.Title)' has a GreaterEquals version clause and '$NewTitle' is published with Exact, so the chain is inconsistent."
    }
    if ($Mode -eq 'ReplaceAndUninstall') {
        foreach ($clause in @(Get-SupersedenceMinimumClause -Manifest $NewManifest)) {
            $warnings += "'$NewTitle' keeps a GreaterEquals version clause ($clause), so the next uninstall through supersedence will be reported as failed."
        }
    }
    return $warnings
}

function Get-SupersedenceDeploymentTypeXml {
    # The stored XML of each deployment type of an application, from the XML of the application itself (it holds the
    # application and every deployment type; Get-CMDeploymentType returns the XML of the first type for every type).
    # Returns objects with Name and Xml. Needs the ConfigMgr site drive.
    param([Parameter(Mandatory)][string]$ApplicationName)
    $app = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
    $document = New-Object System.Xml.XmlDocument
    $document.LoadXml([string]$app.SDMPackageXML)
    $items = @()
    foreach ($node in $document.SelectNodes('/*/*[local-name()="DeploymentType"]')) {
        $title = $node.SelectSingleNode('*[local-name()="Title"]')
        $items += [pscustomobject]@{ Name = $(if ($title) { [string]$title.InnerText } else { '' }); Xml = $node.OuterXml }
    }
    return $items
}

function Invoke-SupersedenceChainCheck {
    # The chain check before a publish: the detection of the old application is read from the site and compared with the
    # new one. It returns warnings and never blocks; a failure to read is itself a warning.
    param(
        [Parameter(Mandatory)][string]$NewTitle, [Parameter(Mandatory)][string]$NewVersion, [Parameter(Mandatory)][string]$Marker,
        [Parameter(Mandatory)][string]$NewRule, [Parameter(Mandatory)][string]$Mode, [Parameter(Mandatory)]$NewManifest
    )
    try {
        $chosen = Select-SupersedenceOldApplication -Applications @(Get-SupersedenceLineApplication -Marker $Marker -NewApplicationName $NewTitle) -NewTitle $NewTitle -NewVersion $NewVersion
        if ($null -eq $chosen.Old) { return @() }
        $clauses = @()
        foreach ($item in @(Get-SupersedenceDeploymentTypeXml -ApplicationName $chosen.Old.Title)) {
            foreach ($op in @(Get-DetectionVersionOperator -DeploymentTypeXml $item.Xml)) { $clauses += [pscustomobject]@{ DeploymentType = $item.Name; Operator = $op } }
        }
        return @(Get-SupersedenceChainWarning -NewTitle $NewTitle -NewRule $NewRule -Mode $Mode -NewManifest $NewManifest -OldDetection ([pscustomobject]@{ Title = $chosen.Old.Title; Clauses = $clauses }))
    }
    catch { return @("The chain check failed: $($_.Exception.Message)") }
}


# ---- the optional rewrite of the version clause of an old application to Exact ----
# The reader reads the one-clause detections that the site stores (checked on ConfigMgr 5.00.9146: the detection method is XML
# text, and its settings and its rule give the clause). The rewrite removes the clause by its logical name and adds the same
# clause with the other operator. Nothing changes on the site until apply runs with a confirmation.

$script:RewriteConfirmationText = 'ConfigMgr evaluates the detection again on the devices. A device that has a newer version than the one in the clause, because it updated itself or got it in another way, then reads "not installed".'

function Get-SupersedenceRewriteConfirmationText { return $script:RewriteConfirmationText }

function Read-SupersedenceStoredDetection {
    # The single version clause of a deployment type, from its stored XML. Returns Clause (or $null) and NotRewritableReason.
    param([AllowNull()][string]$DeploymentTypeXml)
    $fail = { param($reason) [pscustomobject]@{ Clause = $null; NotRewritableReason = $reason } }
    if ([string]::IsNullOrWhiteSpace($DeploymentTypeXml)) { return (& $fail 'the detection could not be read') }
    try { $outer = New-Object System.Xml.XmlDocument; $outer.LoadXml($DeploymentTypeXml) }
    catch [System.Xml.XmlException] { return (& $fail 'the detection could not be read') }
    $holder = $outer.SelectNodes('//*[not(*)][contains(., "<EnhancedDetectionMethod")]') | Select-Object -First 1
    if ($null -eq $holder) { return (& $fail 'the deployment type has no enhanced detection method') }
    try { $method = New-Object System.Xml.XmlDocument; $method.LoadXml($holder.InnerText) }
    catch [System.Xml.XmlException] { return (& $fail 'the detection method could not be read') }

    $settings = @($method.SelectNodes('//*[local-name()="File" or local-name()="SimpleSetting"]'))
    $expressions = @($method.SelectNodes('//*[local-name()="Expression"]'))
    if ($expressions.Count -ne 1 -or $settings.Count -ne 1) { return (& $fail 'the detection has more than one clause or a connector, which is not rewritten') }

    $expression = $expressions[0]
    $operatorNode = $expression.SelectSingleNode('*[local-name()="Operator"]')
    $operator = $(if ($operatorNode) { $operatorNode.InnerText.Trim() } else { '' })
    $reference = $expression.SelectSingleNode('.//*[local-name()="SettingReference"]')
    $constant = $expression.SelectSingleNode('.//*[local-name()="ConstantValue"]')
    if ($null -eq $reference -or $null -eq $constant -or $reference.GetAttribute('DataType') -ine 'Version') { return (& $fail 'the clause is not a version clause') }
    $setting = $settings[0]
    $logicalName = $setting.GetAttribute('LogicalName')
    if ([string]::IsNullOrEmpty($logicalName) -or $logicalName -cne $reference.GetAttribute('SettingLogicalName')) { return (& $fail 'the clause does not match its setting') }
    $expected = $constant.GetAttribute('Value')
    $childText = { param($node, $name) $n = $node.SelectSingleNode("*[local-name()='$name']"); if ($n) { $n.InnerText } else { '' } }

    if ($setting.LocalName -eq 'File') {
        return [pscustomobject]@{ NotRewritableReason = $null; Clause = [pscustomobject]@{
                Kind = 'File'; LogicalName = $logicalName; Operator = $operator; ExpectedValue = $expected
                Path = (& $childText $setting 'Path'); FileName = (& $childText $setting 'Filter'); Is64Bit = ($setting.GetAttribute('Is64Bit') -ieq 'true')
                Hive = ''; Key = ''; ValueName = '' } }
    }
    $source = $setting.SelectSingleNode('.//*[local-name()="RegistryDiscoverySource"]')
    if ($null -eq $source) { return (& $fail 'the clause is not a file or a registry value') }
    return [pscustomobject]@{ NotRewritableReason = $null; Clause = [pscustomobject]@{
            Kind = 'Registry'; LogicalName = $logicalName; Operator = $operator; ExpectedValue = $expected
            Path = ''; FileName = ''; Is64Bit = ($source.GetAttribute('Is64Bit') -ieq 'true')
            Hive = $source.GetAttribute('Hive'); Key = (& $childText $source 'Key'); ValueName = (& $childText $source 'ValueName') } }
}

function ConvertTo-SupersedenceCmdletOperator {
    param([string]$StoredOperator)
    if ($StoredOperator -ieq 'Equals') { return 'IsEquals' }
    return $StoredOperator
}

function Test-SupersedenceHiveSupported {
    # True for the hives that the rewrite writes (HKLM and HKCU).
    param([string]$Hive)
    return ($Hive -ieq 'HKEY_LOCAL_MACHINE' -or $Hive -ieq 'HKEY_CURRENT_USER')
}

function Get-SupersedenceOldOnes {
    # The old applications of the line: every readable, not retired application except the newest version (and except the
    # applications that hold the newest version under another title: 1.0 and 1.0.0 are the same version).
    param([AllowEmptyCollection()][object[]]$Line = @())
    $readable = @($Line | Where-Object { $null -ne (ConvertTo-SupersedenceVersion -Value $_.SoftwareVersion) -and -not $_.IsExpired })
    if ($readable.Count -eq 0) { return @() }
    $newest = @($readable | Sort-Object -Descending -Property @{ Expression = { ConvertTo-SupersedenceVersion -Value $_.SoftwareVersion } })[0]
    return @($readable | Where-Object { (Compare-SupersedenceVersion -A $_.SoftwareVersion -B $newest.SoftwareVersion) -lt 0 } |
            Sort-Object -Descending -Property @{ Expression = { ConvertTo-SupersedenceVersion -Value $_.SoftwareVersion } })
}

function Get-SupersedenceRewriteKey {
    # The key of a deployment type in the plan: the application and the deployment type, so that two applications that hold a
    # type of the same name never share a detection.
    param([string]$Application, [string]$DeploymentType)
    return ($Application + [char]1 + $DeploymentType)
}

function Get-SupersedenceRewritePlan {
    <#
    .SYNOPSIS
        The plan of the rewrite: what would change, and why each other clause is left.
    .DESCRIPTION
        Old is the list of old applications with their deployment types. Detections is a hashtable of stored detections by
        Get-SupersedenceRewriteKey. RequiredCollections is a hashtable of collection names by application title. Guard is a
        script block that takes an application and a clause and returns a verdict. Returns Items (Application, Version,
        DeploymentType, Clause, Skip, Rewrites), Blocked, ClauseCount and ApplicationCount.
    #>
    param([AllowEmptyCollection()][object[]]$Old = @(), [hashtable]$Detections = @{}, [hashtable]$RequiredCollections = @{}, [Parameter(Mandatory)][scriptblock]$Guard)
    $items = @()
    $blocked = @()
    $item = { param($app, $dt, $clause, $skip) [pscustomobject]@{ Application = $app.Title; Version = $app.SoftwareVersion; DeploymentType = $dt.Name; Clause = $clause; Skip = $skip; Rewrites = ($null -ne $clause -and $null -eq $skip) } }
    foreach ($app in @($Old)) {
        $required = @($RequiredCollections[$app.Title])
        if ($required.Count -gt 0 -and $null -ne $required[0]) {
            $blocked += "'$($app.Title)' has a required deployment ($($required -join ', ')). Remove the required deployment, or leave this application out."
            continue
        }
        foreach ($dt in @($app.DeploymentTypes)) {
            $detection = $Detections[(Get-SupersedenceRewriteKey -Application $app.Title -DeploymentType $dt.Name)]
            if ($null -eq $detection -or $null -eq $detection.Clause) {
                $reason = $(if ($detection) { $detection.NotRewritableReason } else { $null })
                $items += (& $item $app $dt $null $(if ($reason) { $reason } else { 'the detection was not read' }))
                continue
            }
            $clause = $detection.Clause
            if ($clause.Kind -eq 'Registry' -and -not (Test-SupersedenceHiveSupported -Hive $clause.Hive)) {
                $items += (& $item $app $dt $clause "the registry hive $($clause.Hive) is not one that the rewrite writes")
            }
            elseif ($clause.Operator -ine 'GreaterEquals') {
                $items += (& $item $app $dt $clause "the operator is $($clause.Operator), not GreaterEquals")
            }
            elseif ($clause.ExpectedValue.Trim() -ine $app.SoftwareVersion.Trim()) {
                $items += (& $item $app $dt $clause "the expected value '$($clause.ExpectedValue)' is not the version of the application")
            }
            else {
                $verdict = & $Guard $app $clause
                if ($verdict.Result -eq 'ExactApplied') { $items += (& $item $app $dt $clause $null) }
                else { $items += (& $item $app $dt $clause ('not rewritten: ' + $verdict.Detail)) }
            }
        }
    }
    $rewriting = @($items | Where-Object { $_.Rewrites })
    return [pscustomobject]@{
        Items            = @($items)
        Blocked          = @($blocked)
        ClauseCount      = $rewriting.Count
        ApplicationCount = @($rewriting | ForEach-Object { $_.Application.ToUpperInvariant() } | Select-Object -Unique).Count
    }
}

function Get-SupersedenceRewriteGuardVerdict {
    # The Exact guard for an old application: the installed value that a test install recorded for the title and the
    # version of that old application must equal the expected value of the clause. Without a value, the clause is not rewritten.
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$TitleKey, [Parameter(Mandatory)]$Application, [Parameter(Mandatory)]$Clause)
    $key = Get-SupersedenceRecordKey -TitleKey $TitleKey -Version $Application.SoftwareVersion
    $installed = $null
    foreach ($entry in @(Get-SupersedenceMapEntry -Map (Get-SupersedenceProperty $Settings 'InstalledValues'))) { if ($entry.Key -ieq $key) { $installed = [string]$entry.Value } }
    if (-not [string]::IsNullOrWhiteSpace($installed)) {
        if (Test-SupersedenceSameVersionText -A $installed -B $Clause.ExpectedValue) { return [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null } }
        return [pscustomobject]@{ Result = 'ExactBlocked'; Detail = "the installed value '$installed' differs from the expected value '$($Clause.ExpectedValue)'" }
    }
    return [pscustomobject]@{ Result = 'ExactBlocked'; Detail = "no installed value is recorded for version $($Application.SoftwareVersion)" }
}

# ---- the log of the rewrite and the restore ----
# The log is the only way back, because the publish script purges the revision history. One JSON object per line.

function ConvertTo-SupersedenceRewriteLogLine {
    param([Parameter(Mandatory)]$Entry)
    return ([ordered]@{
            At = ([datetimeoffset]$Entry.At).ToString('o'); Application = $Entry.Application; DeploymentType = $Entry.DeploymentType
            OldOperator = $Entry.OldOperator; NewOperator = $Entry.NewOperator; LogicalName = $Entry.LogicalName
            Restored = [bool]$Entry.Restored; Line = $Entry.Line
        } | ConvertTo-Json -Compress)
}

function Read-SupersedenceRewriteLog {
    # The entries of a log text. A line that cannot be read is skipped.
    param([AllowNull()][string]$Text)
    $entries = @()
    foreach ($line in (([string]$Text) -split "`n")) {
        $trimmed = $line.Trim()
        if (-not $trimmed) { continue }
        try {
            $o = $trimmed | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $o -or -not $o.PSObject.Properties['Application']) { continue }
            $at = $o.At
            if ($at -isnot [datetime]) { $at = [datetime]::Parse([string]$at, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) }
            $entries += [pscustomobject]@{
                At = ([datetimeoffset]$at); Application = [string]$o.Application; DeploymentType = [string]$o.DeploymentType
                OldOperator = [string]$o.OldOperator; NewOperator = [string]$o.NewOperator; LogicalName = [string]$o.LogicalName
                Restored = [bool]$o.Restored; Line = $(if ($o.PSObject.Properties['Line'] -and $null -ne $o.Line) { [string]$o.Line } else { $null })
            }
        }
        catch { $null = $_ }
    }
    return $entries
}

function Get-SupersedenceRestoreEntry {
    <#
    .SYNOPSIS
        What Restore writes back: for each deployment type of the given applications whose last entry is a rewrite (not a
        restore), the first entry, which holds the operator that the type had before the first rewrite.
    .DESCRIPTION
        With a line (its identity marker), only the entries of that line: the log holds the rewrites of every line, and Restore
        for one line must not touch another. An entry from before the log carried the line has none, and stays open to every line.
    #>
    param([AllowEmptyCollection()][object[]]$Log = @(), [string[]]$Applications = @(), [AllowNull()][string]$Line = $null)
    $selected = @($Log | Where-Object { ($Applications -contains $_.Application) -and ([string]::IsNullOrEmpty($Line) -or [string]::IsNullOrEmpty($_.Line) -or $_.Line -ieq $Line) })
    $groups = $selected | Group-Object -Property { $_.Application.ToUpperInvariant() + [char]1 + $_.DeploymentType.ToUpperInvariant() }
    $restore = @()
    foreach ($group in $groups) {
        $ordered = @($group.Group | Sort-Object -Property At)
        if ($ordered[-1].Restored) { continue }
        $first = $ordered | Where-Object { -not $_.Restored } | Select-Object -First 1
        if ($first) { $restore += $first }
    }
    return $restore
}

# ---- the rewrite on the site ----

function Read-SupersedenceRequiredCollection {
    # The collections that hold a required deployment of an application (OfferTypeID 0).
    param([Parameter(Mandatory)][string]$ApplicationName)
    $found = @()
    foreach ($deployment in @(Get-CMApplicationDeployment -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop)) {
        if ([string]$deployment.ApplicationName -ne $ApplicationName) { continue }
        if ([int]$deployment.OfferTypeID -eq 0) { $found += [string]$deployment.CollectionName }
    }
    return $found
}

function Write-SupersedenceClauseOperator {
    <#
    .SYNOPSIS
        Replaces the version clause of one deployment type with the same clause and another operator.
    .DESCRIPTION
        The clause is removed by its logical name and the same clause is added with the operator (IsEquals or GreaterEquals).
        The clause object is built off the ConfigMgr drive, as the publish does. Needs the ConfigMgr site drive.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode, [Parameter(Mandatory)][string]$ApplicationName, [Parameter(Mandatory)][string]$DeploymentTypeName,
        [Parameter(Mandatory)]$Clause, [Parameter(Mandatory)][ValidateSet('IsEquals', 'GreaterEquals')][string]$NewOperator
    )
    Set-Location C:
    if ($Clause.Kind -eq 'File') {
        $params = @{ Path = $Clause.Path; FileName = $Clause.FileName; PropertyType = 'Version'; Value = $true; ExpressionOperator = $NewOperator; ExpectedValue = $Clause.ExpectedValue }
        if ($Clause.Is64Bit) { $params['Is64Bit'] = $true }
        $newClause = New-CMDetectionClauseFile @params
    }
    else {
        $hive = switch ($Clause.Hive.ToUpperInvariant()) {
            'HKEY_LOCAL_MACHINE' { 'LocalMachine' }
            'HKEY_CURRENT_USER' { 'CurrentUser' }
            default { throw "The registry hive '$($Clause.Hive)' is not one that the rewrite writes." }
        }
        $params = @{ Hive = $hive; KeyName = $Clause.Key; ValueName = $Clause.ValueName; PropertyType = 'Version'; Value = $true; ExpressionOperator = $NewOperator; ExpectedValue = $Clause.ExpectedValue }
        if ($Clause.Is64Bit) { $params['Is64Bit'] = $true }
        $newClause = New-CMDetectionClauseRegistryKeyValue @params
    }
    Set-Location ($SiteCode + ':')
    Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName -RemoveDetectionClause $Clause.LogicalName -AddDetectionClause $newClause -ErrorAction Stop | Out-Null
}

function Get-SupersedenceStoredDetectionMap {
    # The stored detection of each deployment type of an application, by deployment type name.
    param([Parameter(Mandatory)][string]$ApplicationName)
    $map = @{}
    foreach ($item in @(Get-SupersedenceDeploymentTypeXml -ApplicationName $ApplicationName)) { $map[$item.Name] = Read-SupersedenceStoredDetection -DeploymentTypeXml $item.Xml }
    return $map
}

function Get-SupersedenceRewriteScopePlan {
    # Reads the line and the old applications from the site and plans the rewrite. Nothing changes on the site.
    param([Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$TitleKey, [Parameter(Mandatory)]$Settings)
    $old = @(Get-SupersedenceOldOnes -Line @(Get-SupersedenceLineApplication -Marker $Marker))
    $detections = @{}
    $required = @{}
    foreach ($app in $old) {
        $map = Get-SupersedenceStoredDetectionMap -ApplicationName $app.Title
        foreach ($name in $map.Keys) { $detections[(Get-SupersedenceRewriteKey -Application $app.Title -DeploymentType $name)] = $map[$name] }
        $required[$app.Title] = @(Read-SupersedenceRequiredCollection -ApplicationName $app.Title)
    }
    $guard = { param($app, $clause) Get-SupersedenceRewriteGuardVerdict -Settings $Settings -TitleKey $TitleKey -Application $app -Clause $clause }.GetNewClosure()
    return (Get-SupersedenceRewritePlan -Old $old -Detections $detections -RequiredCollections $required -Guard $guard)
}

function Invoke-SupersedenceRewrite {
    <#
    .SYNOPSIS
        Applies a rewrite plan: GreaterEquals to IsEquals, one deployment type at a time.
    .DESCRIPTION
        The log is written ahead: before the site changes, the intent is appended and flushed, so a stop at any point leaves
        an entry for everything that the site may have changed; when the write fails, a restore entry cancels the intent at
        once. A failure of one type is named and the other types are still tried. Returns Written and Failures.
    #>
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$SiteCode, [Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$LogPath)
    $written = @()
    $failures = @()
    foreach ($item in @($Plan.Items | Where-Object { $_.Rewrites })) {
        $clause = $item.Clause
        $intent = [pscustomobject]@{ At = [datetimeoffset]::UtcNow; Application = $item.Application; DeploymentType = $item.DeploymentType
            OldOperator = (ConvertTo-SupersedenceCmdletOperator -StoredOperator $clause.Operator); NewOperator = 'IsEquals'; LogicalName = $clause.LogicalName; Restored = $false; Line = $Marker }
        try { Add-SupersedenceLogLine -Path $LogPath -Entry $intent }
        catch { $failures += ('{0} / {1}: {2}' -f $item.Application, $item.DeploymentType, $_.Exception.Message); break }
        try {
            Write-SupersedenceClauseOperator -SiteCode $SiteCode -ApplicationName $item.Application -DeploymentTypeName $item.DeploymentType -Clause $clause -NewOperator 'IsEquals'
            $written += $intent
            Write-Log ('Rewritten: {0} / {1}: GreaterEquals -> IsEquals' -f $item.Application, $item.DeploymentType)
        }
        catch {
            $cancel = $intent.PSObject.Copy(); $cancel.At = [datetimeoffset]::UtcNow; $cancel.OldOperator = 'IsEquals'; $cancel.NewOperator = 'GreaterEquals'; $cancel.Restored = $true
            Add-SupersedenceLogLine -Path $LogPath -Entry $cancel
            $failures += ('{0} / {1}: {2}' -f $item.Application, $item.DeploymentType, $_.Exception.Message)
        }
    }
    return [pscustomobject]@{ Written = $written; Failures = $failures }
}

function Add-SupersedenceLogLine {
    # Appends one entry and closes the file, so the entry is on disk before the caller goes on. A failure stops the rewrite.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Entry)
    $folder = Split-Path -Parent $Path
    if ($folder -and -not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read, 4096, [System.IO.FileOptions]::WriteThrough)
    try {
        $writer = New-Object System.IO.StreamWriter($stream)
        try { $writer.WriteLine((ConvertTo-SupersedenceRewriteLogLine -Entry $Entry)) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Invoke-SupersedenceRestore {
    # Writes the old operator back for the deployment types that the log says were rewritten and not restored. A restore entry is
    # appended after each success. A log entry for an application that the line no longer holds is named and left.
    param([Parameter(Mandatory)][string]$SiteCode, [Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$LogPath)
    $log = $(if (Test-Path -LiteralPath $LogPath) { @(Read-SupersedenceRewriteLog -Text (Get-Content -LiteralPath $LogPath -Raw -Encoding UTF8)) } else { @() })
    $titles = @(Get-SupersedenceLineApplication -Marker $Marker | ForEach-Object { $_.Title })
    $written = @()
    $failures = @()
    $applications = @($log | ForEach-Object { $_.Application } | Select-Object -Unique)
    foreach ($entry in @(Get-SupersedenceRestoreEntry -Log $log -Applications $applications -Line $Marker)) {
        $known = [bool]($titles | Where-Object { $_ -ieq $entry.Application })
        # An entry from before the log carried the line, whose application is not in this line, belongs to another line: it is left.
        if (-not $entry.Line -and -not $known) { continue }
        if (-not $known) { $failures += ('{0} / {1}: the application is not in this line on the site any more (renamed, removed or without the marker), so it was not restored.' -f $entry.Application, $entry.DeploymentType); continue }
        try {
            $map = Get-SupersedenceStoredDetectionMap -ApplicationName $entry.Application
            $current = $map[$entry.DeploymentType]
            if ($null -eq $current -or $null -eq $current.Clause) { $failures += ('{0} / {1}: the detection is not a single version clause any more, so it was not restored.' -f $entry.Application, $entry.DeploymentType); continue }
            Write-SupersedenceClauseOperator -SiteCode $SiteCode -ApplicationName $entry.Application -DeploymentTypeName $entry.DeploymentType -Clause $current.Clause -NewOperator (ConvertTo-SupersedenceCmdletOperator -StoredOperator $entry.OldOperator)
            # The restore entry carries the line of the entry that it cancels.
            $done = [pscustomobject]@{ At = [datetimeoffset]::UtcNow; Application = $entry.Application; DeploymentType = $entry.DeploymentType; OldOperator = (ConvertTo-SupersedenceCmdletOperator -StoredOperator $current.Clause.Operator)
                NewOperator = $entry.OldOperator; LogicalName = $current.Clause.LogicalName; Restored = $true; Line = $entry.Line }
            Add-SupersedenceLogLine -Path $LogPath -Entry $done
            $written += $done
            Write-Log ('Restored: {0} / {1}: back to {2}' -f $entry.Application, $entry.DeploymentType, $entry.OldOperator)
        }
        catch { $failures += ('{0} / {1}: {2}' -f $entry.Application, $entry.DeploymentType, $_.Exception.Message) }
    }
    return [pscustomobject]@{ Written = $written; Failures = $failures }
}

# ---- the window of the rewrite: the change plan, its run and its result ----
# The window shows what the rewrite changes before anything is written (a port of RewriteOldVersionsPlans of App Packager
# Desktop). The plan, the notes, the confirmation text, the result rows and the summary come from these functions, so the
# window and the tests read the same text. The write itself is Invoke-SupersedenceRewrite and Invoke-SupersedenceRestore.

$script:RestoreConfirmationText = 'The old operator is written back on the site. A device that now reads "installed" through the clause may read it differently.'

function Get-SupersedenceRestoreConfirmationText { return $script:RestoreConfirmationText }

function ConvertTo-SupersedenceRewriteChangeItem {
    # One row of the plan of a rewrite.
    param([Parameter(Mandatory)]$Item)
    return [pscustomobject]@{
        Action = 'Rewrite clause'; Target = ('{0} / {1}' -f $Item.Application, $Item.DeploymentType); Detail = 'GreaterEquals -> IsEquals'
        Application = [string]$Item.Application; DeploymentType = [string]$Item.DeploymentType
    }
}

function ConvertTo-SupersedenceRestoreChangeItem {
    # One row of the plan of a restore: the operator that comes back.
    param([Parameter(Mandatory)]$Entry)
    return [pscustomobject]@{
        Action = 'Restore clause'; Target = ('{0} / {1}' -f $Entry.Application, $Entry.DeploymentType); Detail = ('{0} -> {1}' -f $Entry.NewOperator, $Entry.OldOperator)
        Application = [string]$Entry.Application; DeploymentType = [string]$Entry.DeploymentType
    }
}

function Get-SupersedenceRewriteNote {
    # The lines above the list: the effect on the devices, what is blocked, what is left and why, and the pending restores.
    param([Parameter(Mandatory)]$Plan, [int]$Restorable = 0)
    $notes = @($script:RewriteConfirmationText)
    foreach ($blocked in @($Plan.Blocked)) { $notes += ('Blocked: ' + $blocked) }
    foreach ($item in @($Plan.Items | Where-Object { -not $_.Rewrites })) { $notes += ('Left: {0} / {1}: {2}' -f $item.Application, $item.DeploymentType, $item.Skip) }
    if ($Restorable -gt 0) { $notes += ('{0} earlier rewrite(s) of this line can still be restored with the button Restore Old Versions.' -f $Restorable) }
    return $notes
}

function New-SupersedenceRewriteChangePlan {
    # The plan that the window shows for a rewrite, or $null when the plan has no clause to rewrite.
    param([Parameter(Mandatory)]$Plan, [int]$Restorable = 0)
    $rewrites = @($Plan.Items | Where-Object { $_.Rewrites })
    if ($rewrites.Count -eq 0) { return $null }
    return [pscustomobject]@{
        Kind = 'Rewrite'; Title = 'Rewrite Old Versions'
        ActionName = ('Rewrite {0} {1}' -f $rewrites.Count, $(if ($rewrites.Count -eq 1) { 'Clause' } else { 'Clauses' }))
        Items = @($rewrites | ForEach-Object { ConvertTo-SupersedenceRewriteChangeItem -Item $_ })
        IsDestructive = $true; HasUndo = $true; DestructiveConfirmation = $script:RewriteConfirmationText
        Notes = @(Get-SupersedenceRewriteNote -Plan $Plan -Restorable $Restorable); LogName = 'Rewrite old version clauses'
    }
}

function Get-SupersedenceRewritePending {
    # The rewrites of the line that the log lists and no restore entry cancels. The log holds every line, and a restore for one line
    # leaves the others.
    param([Parameter(Mandatory)][string]$LogPath, [Parameter(Mandatory)][string]$Marker)
    $log = $(if (Test-Path -LiteralPath $LogPath) { @(Read-SupersedenceRewriteLog -Text (Get-Content -LiteralPath $LogPath -Raw -Encoding UTF8)) } else { @() })
    if ($log.Count -eq 0) { return @() }
    return @(Get-SupersedenceRestoreEntry -Log $log -Applications @($log | ForEach-Object { $_.Application } | Select-Object -Unique) -Line $Marker)
}

function New-SupersedenceRestoreChangePlan {
    # The plan that the window shows for a restore, or $null when no earlier rewrite of the line is pending.
    param([AllowEmptyCollection()][object[]]$Pending = @())
    $pending = @($Pending | Where-Object { $null -ne $_ })
    if ($pending.Count -eq 0) { return $null }
    return [pscustomobject]@{
        Kind = 'Restore'; Title = 'Restore Old Versions'
        ActionName = ('Restore {0} {1}' -f $pending.Count, $(if ($pending.Count -eq 1) { 'Clause' } else { 'Clauses' }))
        Items = @($pending | ForEach-Object { ConvertTo-SupersedenceRestoreChangeItem -Entry $_ })
        IsDestructive = $true; HasUndo = $false; DestructiveConfirmation = $script:RestoreConfirmationText
        Notes = @(); LogName = 'Restore old version clauses'
    }
}

function Test-SupersedenceSameRewrite {
    # True when both plans rewrite the same clauses (application, deployment type, setting, operator and expected value).
    param([Parameter(Mandatory)]$Shown, [Parameter(Mandatory)]$Fresh)
    $keys = {
        param($plan)
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($i in @($plan.Items | Where-Object { $_.Rewrites })) { [void]$set.Add(('{0}|{1}|{2}|{3}|{4}' -f $i.Application, $i.DeploymentType, $i.Clause.LogicalName, $i.Clause.Operator, $i.Clause.ExpectedValue)) }
        , $set
    }
    $a = & $keys $Shown
    $b = & $keys $Fresh
    return $a.SetEquals($b)
}

function New-SupersedenceChangeRow {
    param($Item, [bool]$Succeeded, [string]$Message, [bool]$Skipped = $false, [bool]$Partial = $false)
    return [pscustomobject]@{
        Outcome = $(if ($Skipped) { 'Skipped' } elseif ($Succeeded) { 'Changed' } else { 'Failed' }); Action = $Item.Action; Target = $Item.Target
        Message = $Message; Succeeded = $Succeeded; Skipped = $Skipped; PartialChange = $Partial
    }
}

function Get-SupersedenceChangeSummary {
    # The summary line of a result: the counts, and the note about a failed change that was made in part.
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)
    $rows = @($Rows)
    $changed = @($rows | Where-Object { $_.Succeeded -and -not $_.Skipped }).Count
    $failed = @($rows | Where-Object { -not $_.Succeeded -and -not $_.Skipped }).Count
    $skipped = @($rows | Where-Object { $_.Skipped }).Count
    $partial = @($rows | Where-Object { $_.PartialChange -and -not $_.Succeeded -and -not $_.Skipped }).Count
    $text = $(if ($skipped -eq 0) { '{0} changed, {1} failed.' -f $changed, $failed } else { '{0} changed, {1} failed, {2} skipped.' -f $changed, $failed, $skipped })
    if ($partial -eq 1) { $text += ' 1 failed change was made in part; Undo can reverse it.' }
    elseif ($partial -gt 1) { $text += (' {0} failed changes were made in part; Undo can reverse them.' -f $partial) }
    return $text
}

function New-SupersedenceChangeResult {
    param([AllowEmptyCollection()][object[]]$Rows = @(), [ValidateSet('Completed', 'Cancelled', 'Failed')][string]$Stop = 'Completed', [AllowNull()][object]$UndoAvailable = $null, [string]$Summary = '')
    $rows = @($Rows)
    if (-not $Summary) { $Summary = Get-SupersedenceChangeSummary -Rows $rows }
    return [pscustomobject]@{
        Items = $rows; Stop = $Stop; Summary = $Summary
        UndoAvailable = $(if ($null -ne $UndoAvailable) { [bool]$UndoAvailable } else { [bool]@($rows | Where-Object { $_.Succeeded -and -not $_.Skipped }).Count })
    }
}

function ConvertTo-SupersedenceChangeResult {
    # Maps the outcome of a rewrite or a restore (Written, Failures) to one result row per listed clause. A failure text that names the
    # clause ("<application> / <deployment type>: <reason>") is its message.
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items, [Parameter(Mandatory)]$Outcome)
    $rows = @()
    foreach ($item in @($Items)) {
        $done = @(@($Outcome.Written) | Where-Object { $null -ne $_ -and $_.Application -ieq $item.Application -and $_.DeploymentType -ieq $item.DeploymentType }).Count -gt 0
        $prefix = '{0} / {1}:' -f $item.Application, $item.DeploymentType
        $failure = @(@($Outcome.Failures) | Where-Object { $null -ne $_ -and ([string]$_).StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) }) | Select-Object -First 1
        if ($done) { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $true -Message 'Written to the site and the log.') }
        elseif ($failure) { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $false -Message ([string]$failure).Substring($prefix.Length).Trim()) }
        else { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $false -Skipped $true -Message 'Not reached: the run stopped before this clause, or the entry belongs to another line (it is from before the log carried the line) and was left.') }
    }
    return (New-SupersedenceChangeResult -Rows $rows)
}

function Get-SupersedenceStoppedResult {
    # The result of a run that stopped without its own result (cancelled, or failed): a clause that the log holds an intent for, and
    # that no restore entry cancels, may be changed (the run stopped while it wrote); every other clause did not report.
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items, [AllowEmptyCollection()][object[]]$Entries = @(),
        [ValidateSet('Cancelled', 'Failed')][string]$Stop = 'Cancelled', [string]$Detail = '', [ValidateSet('Rewrite', 'Restore')][string]$Kind = 'Rewrite'
    )
    $word = $(if ($Stop -eq 'Cancelled') { 'cancelled' } else { 'failed' })
    $rows = @()
    foreach ($item in @($Items)) {
        $mine = @(@($Entries) | Where-Object { $null -ne $_ -and $_.Application -ieq $item.Application -and $_.DeploymentType -ieq $item.DeploymentType } | Sort-Object -Property At)
        $open = ($mine.Count -gt 0 -and -not $mine[-1].Restored)
        # A restore writes its entry after the site took the operator, so a restore entry is a clause that is back.
        if ($Kind -eq 'Restore') {
            if ($mine.Count -gt 0 -and $mine[-1].Restored) { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $true -Message 'Written to the site and the log.') }
            else { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $false -Message ('No result: the run was {0} before this change reported.' -f $word)) }
        }
        elseif ($open) { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $false -Partial $true -Message ('The run was {0} while this clause was written. It may be changed; Undo restores it.' -f $word)) }
        else { $rows += (New-SupersedenceChangeRow -Item $item -Succeeded $false -Message ('No result: the run was {0} before this change reported.' -f $word)) }
    }
    $maybeChanged = @($rows | Where-Object { $_.PartialChange }).Count
    $backCount = @($rows | Where-Object { $_.Succeeded }).Count
    $text = $(if ($Stop -eq 'Cancelled') { 'The run was cancelled. Changes made before the cancel stay in place.' } else { 'The run failed. The log has the details.' })
    if ($Detail) { $text += (' ' + $Detail) }
    if ($rows.Count -gt 0) { $text += (' Before the {0}: {1} changed, {2} failed.' -f $(if ($Stop -eq 'Cancelled') { 'cancel' } else { 'failure' }), $backCount, ($rows.Count - $backCount)) }
    return (New-SupersedenceChangeResult -Rows $rows -Stop $Stop -UndoAvailable ($maybeChanged -gt 0) -Summary $text)
}

function Get-SupersedenceLogEntrySince {
    # The entries of the log that were written at or after a point in time: what one run wrote.
    param([Parameter(Mandatory)][string]$LogPath, [Parameter(Mandatory)][datetimeoffset]$Since)
    if (-not (Test-Path -LiteralPath $LogPath)) { return @() }
    return @(Read-SupersedenceRewriteLog -Text (Get-Content -LiteralPath $LogPath -Raw -Encoding UTF8) | Where-Object { $_.At -ge $Since })
}

function ConvertFrom-SupersedenceChildRun {
    <#
    .SYNOPSIS
        The result of the child process that ran a rewrite or a restore for the window.
    .DESCRIPTION
        Run holds Cancelled, Result (the APRESULT line of the child, read as JSON, or $null), ExitCode and StdErr. A cancelled run, or a
        run that gave no result, is reported from the log entries that it wrote (Entries).
    #>
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items, [AllowEmptyCollection()][object[]]$Entries = @(), [ValidateSet('Rewrite', 'Restore')][string]$Kind = 'Rewrite')
    if ($Run.Cancelled) { return (Get-SupersedenceStoppedResult -Items $Items -Entries $Entries -Stop 'Cancelled' -Kind $Kind) }
    $result = $(if ($Run.Result -and [string]$Run.Result.Kind -eq 'Result') { $Run.Result.Result } else { $null })
    if ($null -eq $result) {
        $detail = ([string]$Run.StdErr).Trim()
        if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) }
        return (Get-SupersedenceStoppedResult -Items $Items -Entries $Entries -Stop 'Failed' -Detail $detail -Kind $Kind)
    }
    return (New-SupersedenceChangeResult -Rows @($result.Items) -Stop ([string]$result.Stop) -UndoAvailable ([bool]$result.UndoAvailable) -Summary ([string]$result.Summary))
}

function Invoke-SupersedenceRewriteRun {
    <#
    .SYNOPSIS
        The run of the rewrite after the confirmation: reads the site again, writes only when the plan is the one that was shown.
    .DESCRIPTION
        The plan was read before the confirmation, and the site can change while the window is open, so it is read again here,
        before anything is written; when the clauses to rewrite are not the ones that were shown, nothing is written. Replan
        and Write are test hooks: the first returns the fresh plan, the second takes it and returns Written and Failures.
        Returns the result: Items (one row per shown clause), Stop, Summary and UndoAvailable.
    #>
    param(
        [Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$SiteCode, [Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$TitleKey,
        [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$LogPath, [AllowNull()][scriptblock]$Replan = $null, [AllowNull()][scriptblock]$Write = $null
    )
    $items = @($Plan.Items | Where-Object { $_.Rewrites } | ForEach-Object { ConvertTo-SupersedenceRewriteChangeItem -Item $_ })
    $skipAll = {
        param($text)
        New-SupersedenceChangeResult -Rows @($items | ForEach-Object { New-SupersedenceChangeRow -Item $_ -Succeeded $false -Skipped $true -Message $text }) -Stop 'Failed' -UndoAvailable $false
    }
    try { $fresh = $(if ($Replan) { & $Replan } else { Get-SupersedenceRewriteScopePlan -Marker $Marker -TitleKey $TitleKey -Settings $Settings }) }
    catch { return (& $skipAll ('The site could not be read again, so nothing was written: ' + $_.Exception.Message)) }
    if (-not (Test-SupersedenceSameRewrite -Shown $Plan -Fresh $fresh)) {
        return (& $skipAll 'The site changed while the window was open: the clauses to rewrite are not the ones that were shown. Nothing was written. Open the window again.')
    }
    $outcome = $(if ($Write) { & $Write $Plan } else { Invoke-SupersedenceRewrite -Plan $Plan -SiteCode $SiteCode -Marker $Marker -LogPath $LogPath })
    return (ConvertTo-SupersedenceChangeResult -Items $items -Outcome $outcome)
}

function Invoke-SupersedenceRestoreRun {
    # The run of a restore after the confirmation: the pending entries of the line are listed first, so each one has a result row.
    param([Parameter(Mandatory)][string]$SiteCode, [Parameter(Mandatory)][string]$Marker, [Parameter(Mandatory)][string]$LogPath, [AllowNull()][scriptblock]$Write = $null)
    $items = @(Get-SupersedenceRewritePending -LogPath $LogPath -Marker $Marker | ForEach-Object { ConvertTo-SupersedenceRestoreChangeItem -Entry $_ })
    $outcome = $(if ($Write) { & $Write } else { Invoke-SupersedenceRestore -SiteCode $SiteCode -Marker $Marker -LogPath $LogPath })
    return (ConvertTo-SupersedenceChangeResult -Items $items -Outcome $outcome)
}

# ---------------------------------------------------------------------------
# ConfigMgr application creation from manifest
# ---------------------------------------------------------------------------

function New-SingleDetectionClause {
    <#
    .SYNOPSIS
        Builds a single CM detection clause object from a manifest detection block.
    .DESCRIPTION
        Internal helper for New-MECMApplicationFromManifest. Supports
        RegistryKeyValue, RegistryKey, and File detection types.
        Must be called while the current location is a filesystem drive
        (not the CM PSDrive).
    #>
    param([Parameter(Mandatory)][pscustomobject]$Det)

    $type = if ($Det.Type) { $Det.Type } else { 'RegistryKeyValue' }
    # Hive defaults to HKLM; a per-user manifest names CurrentUser, which
    # only evaluates for an InstallForUser deployment type.
    $hive = if ($Det.PSObject.Properties['Hive'] -and [string]$Det.Hive -match '^(CurrentUser|HKCU)$') { 'CurrentUser' } else { 'LocalMachine' }

    switch ($type) {
        'RegistryKeyValue' {
            $operator = if ($Det.Operator) { $Det.Operator } else { 'IsEquals' }
            $expected = if ($Det.ExpectedValue) { $Det.ExpectedValue } else { $Det.DisplayVersion }
            $valName  = if ($Det.ValueName) { $Det.ValueName } else { 'DisplayVersion' }
            $propType = if ($Det.PropertyType) { $Det.PropertyType } else { 'String' }

            $p = @{
                Hive               = $hive
                KeyName            = $Det.RegistryKeyRelative
                ValueName          = $valName
                PropertyType       = $propType
                Value              = $true
                ExpressionOperator = $operator
                ExpectedValue      = $expected
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKeyValue @p)
        }
        'RegistryKey' {
            $p = @{
                Hive      = $hive
                KeyName   = $Det.RegistryKeyRelative
                Existence = $true
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKey @p)
        }
        'File' {
            if ($Det.PropertyType -eq 'Existence') {
                $p = @{
                    Path      = $Det.FilePath
                    FileName  = $Det.FileName
                    Existence = $true
                }
            }
            else {
                $op = if ($Det.Operator) { $Det.Operator } else { 'GreaterEquals' }
                $p = @{
                    Path               = $Det.FilePath
                    FileName           = $Det.FileName
                    PropertyType       = $Det.PropertyType
                    Value              = $true
                    ExpressionOperator = $op
                    ExpectedValue      = $Det.ExpectedValue
                }
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseFile @p)
        }
        default { throw "Unsupported detection clause type: $type" }
    }
}

function Test-PsadtLayout {
    <#
    .SYNOPSIS
        Detects the PSADT toolkit generation in a folder and returns the ConfigMgr
        deployment type command lines for it.

    .DESCRIPTION
        v4 layouts carry Invoke-AppDeployToolkit.ps1 at the root (usually with
        Invoke-AppDeployToolkit.exe beside it); v3 layouts carry
        Deploy-Application.exe / Deploy-Application.ps1. The native .exe
        launcher is preferred when present because it hides the PowerShell
        window; otherwise the command line falls back to powershell.exe -File.

        DeployMode is deliberately omitted from the generated command lines by
        default: PSADT runs Interactive when a user session exists and degrades
        to NonInteractive otherwise, which is the close-app/defer behavior that
        justifies wrapping with PSADT at all. Pass -DeployMode to force one.

    .OUTPUTS
        [pscustomobject] Generation ('v3'|'v4'), EntryPoint (file name),
        InstallCommandLine, UninstallCommandLine.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('', 'Interactive', 'Silent', 'NonInteractive')]
        [string]$DeployMode = ''
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "PSADT toolkit folder not found: $Path"
    }

    $modeSuffix = if ([string]::IsNullOrWhiteSpace($DeployMode)) { '' } else { " -DeployMode $DeployMode" }
    $policy = Resolve-CommonSigningPolicy
    $signed = (Get-PolicyFlag -Policy $policy -Name 'SignDeployment') -or (Get-PolicyFlag -Policy $policy -Name 'RequireDeployment')

    $v4Exe = Join-Path $Path 'Invoke-AppDeployToolkit.exe'
    $v4Ps1 = Join-Path $Path 'Invoke-AppDeployToolkit.ps1'
    $v3Exe = Join-Path $Path 'Deploy-Application.exe'
    $v3Ps1 = Join-Path $Path 'Deploy-Application.ps1'

    # The toolkit's native launchers start PowerShell with their own
    # execution-policy override, which a signed deployment must not carry, so
    # signed mode always enters through the script.
    if (Test-Path -LiteralPath $v4Ps1) {
        if ((Test-Path -LiteralPath $v4Exe) -and -not $signed) {
            $entry = 'Invoke-AppDeployToolkit.exe'
            $prefix = $entry
        }
        else {
            $entry = 'Invoke-AppDeployToolkit.ps1'
            $prefix = Get-DeploymentLauncherCommandLine -ScriptToken $entry -Signed:$signed -ExecutableName 'powershell.exe'
        }
        return [pscustomobject]@{
            Generation           = 'v4'
            EntryPoint           = $entry
            InstallCommandLine   = "$prefix -DeploymentType Install$modeSuffix"
            UninstallCommandLine = "$prefix -DeploymentType Uninstall$modeSuffix"
        }
    }

    if ((Test-Path -LiteralPath $v3Exe) -or (Test-Path -LiteralPath $v3Ps1)) {
        if ((Test-Path -LiteralPath $v3Exe) -and -not $signed) {
            $entry = 'Deploy-Application.exe'
            $prefix = $entry
        }
        elseif (Test-Path -LiteralPath $v3Ps1) {
            $entry = 'Deploy-Application.ps1'
            $prefix = Get-DeploymentLauncherCommandLine -ScriptToken $entry -Signed:$signed -ExecutableName 'powershell.exe'
        }
        else {
            throw "Signed deployment mode needs Deploy-Application.ps1 in '$Path'; the toolkit's .exe launcher applies its own execution-policy override and cannot be the signed entry point."
        }
        return [pscustomobject]@{
            Generation           = 'v3'
            EntryPoint           = $entry
            InstallCommandLine   = "$prefix -DeploymentType `"Install`"$modeSuffix"
            UninstallCommandLine = "$prefix -DeploymentType `"Uninstall`"$modeSuffix"
        }
    }

    throw "No PSADT entry point found in '$Path' (expected Invoke-AppDeployToolkit.ps1 [v4] or Deploy-Application.exe/.ps1 [v3])."
}

function Get-NextPatchVersion {
    <#
    .SYNOPSIS
        Returns the version string with its last dotted numeric component
        incremented by one ('8.0.20' -> '8.0.21').
    .DESCRIPTION
        Used by packagers whose detection accepts the packaged version OR its
        immediate successor, so last month's still-active deployment stops
        reinstalling over an in-place upgrade. Returns $null when the last
        component is not purely numeric (for example preview/rc suffixes), so callers
        can fall back to single-version detection.
    #>
    param([Parameter(Mandatory)][string]$Version)

    # Every dotted component must be purely numeric; '10.0.0-rc.1' splits to a
    # numeric final segment and must not be treated as an incrementable patch.
    $parts = $Version -split '\.'
    $n = 0
    foreach ($p in $parts) {
        if (-not [int]::TryParse($p, [ref]$n)) { return $null }
    }
    $parts[$parts.Count - 1] = [string]([int]$parts[$parts.Count - 1] + 1)
    return ($parts -join '.')
}

function New-ArpEntryDetectionScript {
    <#
    .SYNOPSIS
        Returns a PowerShell detection script that finds a product by its own
        Add/Remove Programs entry and compares its DisplayVersion.
    .DESCRIPTION
        Entry has the shape of a WsusDetection block of type ArpEntry: View
        (32, 64 or Both), DisplayName, or DisplayNamePrefix with an optional
        DisplayNameSuffix, an optional Publisher, WindowsInstaller (true
        keeps to entries that Windows Installer wrote), and Version. The script
        opens each registry view explicitly, so a 32-bit and a 64-bit host
        read the same keys. It writes 'Installed' when an entry matches with a
        DisplayVersion at or above Version, writes nothing otherwise, and
        exits 0. A subkey that cannot be read is skipped; an Uninstall key
        that cannot be listed stops the script with an error, which ConfigMgr
        and Intune report as a failed detection, not as a missing product.
        Versions compare as four numbers, a missing part as zero; a
        DisplayVersion that is not up to four numbers never matches. Names
        and the publisher compare without case.
    #>
    param([Parameter(Mandatory)]$Entry)

    $read = {
        param([string]$Name)
        if ($Entry -is [System.Collections.IDictionary]) {
            if ($Entry.Contains($Name)) { return [string]$Entry[$Name] }
            return ''
        }
        $property = $Entry.PSObject.Properties[$Name]
        if ($property) { return [string]$property.Value }
        return ''
    }
    $quote = { param([string]$Text) "'" + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($Text) + "'" }

    $type = & $read 'Type'
    if ($type -and $type -ne 'ArpEntry') { throw ("The entry type '{0}' is not ArpEntry." -f $type) }
    $exactName = & $read 'DisplayName'
    $prefix = & $read 'DisplayNamePrefix'
    $suffix = & $read 'DisplayNameSuffix'
    $publisher = & $read 'Publisher'
    $windowsInstaller = ((& $read 'WindowsInstaller') -match '^(true|1)$')
    if (-not $exactName -and -not $prefix) { throw 'The entry names no DisplayName or DisplayNamePrefix, so the script cannot find the Add/Remove Programs entry.' }
    $version = (& $read 'Version').Trim()
    $parts = @($version.Split('.'))
    $number = 0
    if ($version -notmatch '^\d+(\.\d+){0,3}$' -or @($parts | Where-Object { -not [int]::TryParse($_, [ref]$number) }).Count -gt 0) {
        throw ("The entry compares DisplayVersion with '{0}', which is not a version of up to four numbers." -f $version)
    }
    while ($parts.Count -lt 4) { $parts += '0' }
    $viewText = & $read 'View'
    $views = switch -Regex ($viewText) {
        '^32$' { @('Registry32') }
        '^64$' { @('Registry64') }
        '^Both$' { @('Registry32', 'Registry64') }
        default { throw ("The entry view '{0}' is not 32, 64 or Both." -f $viewText) }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('$minimum = [version]''{0}''' -f ($parts -join '.')))
    $lines.Add(('foreach ($view in @({0})) {{' -f (($views | ForEach-Object { '[Microsoft.Win32.RegistryView]::' + $_ }) -join ', ')))
    $lines.Add('    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)')
    $lines.Add('    $uninstall = $base.OpenSubKey(''SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'')')
    $lines.Add('    if (-not $uninstall) { continue }')
    $lines.Add('    foreach ($name in $uninstall.GetSubKeyNames()) {')
    $lines.Add('        $key = $null')
    $lines.Add('        try {')
    $lines.Add('            $key = $uninstall.OpenSubKey($name)')
    $lines.Add('            if (-not $key) { continue }')
    $lines.Add('            $displayName = [string]$key.GetValue(''DisplayName'')')
    $lines.Add('            $publisher = [string]$key.GetValue(''Publisher'')')
    $lines.Add('            $displayVersion = [string]$key.GetValue(''DisplayVersion'')')
    $lines.Add('            $windowsInstaller = $key.GetValue(''WindowsInstaller'')')
    $lines.Add('        }')
    $lines.Add('        catch { continue }')
    $lines.Add('        finally { if ($key) { $key.Close() } }')
    if ($exactName) {
        $lines.Add(('        if (-not [string]::Equals($displayName, {0}, [StringComparison]::OrdinalIgnoreCase)) {{ continue }}' -f (& $quote $exactName)))
    }
    else {
        $lines.Add(('        if (-not $displayName.StartsWith({0}, [StringComparison]::OrdinalIgnoreCase)) {{ continue }}' -f (& $quote $prefix)))
        if ($suffix) { $lines.Add(('        if (-not $displayName.EndsWith({0}, [StringComparison]::OrdinalIgnoreCase)) {{ continue }}' -f (& $quote $suffix))) }
    }
    if ($publisher) { $lines.Add(('        if (-not [string]::Equals($publisher, {0}, [StringComparison]::OrdinalIgnoreCase)) {{ continue }}' -f (& $quote $publisher))) }
    if ($windowsInstaller) { $lines.Add('        if ([string]$windowsInstaller -ne ''1'') { continue }') }
    $lines.Add('        $parts = @($displayVersion.Trim().Split(''.''))')
    $lines.Add('        if ($parts.Count -gt 4) { continue }')
    $lines.Add('        $numbers = @(foreach ($part in $parts) { $n = 0; if ([int]::TryParse($part, [ref]$n) -and $n -ge 0) { $n } })')
    $lines.Add('        if ($numbers.Count -ne $parts.Count) { continue }')
    $lines.Add('        while ($numbers.Count -lt 4) { $numbers += 0 }')
    $lines.Add('        if ((New-Object System.Version $numbers[0], $numbers[1], $numbers[2], $numbers[3]) -ge $minimum) { Write-Output ''Installed''; exit 0 }')
    $lines.Add('    }')
    $lines.Add('    $uninstall.Close()')
    $lines.Add('}')
    $lines.Add('exit 0')
    return ($lines -join "`r`n")
}

function Test-MECMApplicationHasDeploymentType {
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$DeploymentTypeName
    )

    if (-not (Get-Command -Name Get-CMDeploymentType -ErrorAction SilentlyContinue)) {
        throw "Get-CMDeploymentType is not available; cannot validate existing application '$ApplicationName'."
    }

    try {
        $deploymentTypes = @(Get-CMDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName -ErrorAction SilentlyContinue)
    }
    catch {
        $deploymentTypes = @()
    }

    if ($deploymentTypes.Count -eq 0) {
        try {
            $deploymentTypes = @(Get-CMDeploymentType -ApplicationName $ApplicationName -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.LocalizedDisplayName -eq $DeploymentTypeName -or
                    $_.DeploymentTypeName -eq $DeploymentTypeName -or
                    $_.Name -eq $DeploymentTypeName
                })
        }
        catch {
            $deploymentTypes = @()
        }
    }

    return ($deploymentTypes.Count -gt 0)
}


function Get-ConditionTemplatesPath {
    return (Join-Path $PSScriptRoot 'condition-templates.json')
}

function Get-DefaultConditionTemplates {
    <#
    .SYNOPSIS
        Returns the built-in condition template document.

    .DESCRIPTION
        Each template maps a stable Id (referenced by requirement specs) to
        a CM global condition. Kind selects the creation cmdlet (BuiltIn is
        never created, only resolved); RuleType selects the requirement
        rule cmdlet the spec value binds through.
    #>
    return [pscustomobject]@{
        SchemaVersion = 1
        Conditions    = @(
            [pscustomobject]@{
                Id                  = 'cpu-arch'
                GlobalConditionName = 'APKG - CPU Architecture'
                Kind                = 'Wql'
                RuleType            = 'CommonValue'
                DataType            = 'Integer'
                Namespace           = 'root\cimv2'
                Class               = 'Win32_Processor'
                Property            = 'Architecture'
                # Win32_OperatingSystem.OSArchitecture returns a localized
                # string; the numeric processor architecture compares the
                # same on every OS language.
                Values              = [pscustomobject]@{ x64 = '9'; ARM64 = '12' }
                Description         = 'Win32_Processor.Architecture (9 = x64, 12 = ARM64).'
            }
            [pscustomobject]@{
                Id                  = 'os-language'
                GlobalConditionName = 'Operating System Language'
                Kind                = 'BuiltIn'
                RuleType            = 'OperatingSystemLanguage'
                # The site ships two conditions under this name; PlatformType
                # 1 is Windows, 2 is Mobile.
                PlatformType        = 1
            }
            [pscustomobject]@{
                Id                  = 'vpn-connected'
                GlobalConditionName = 'APKG - VPN Connected'
                Kind                = 'Script'
                RuleType            = 'Boolean'
                DataType            = 'Boolean'
                AdapterPatterns     = @('Cisco AnyConnect', 'Fortinet', 'Juniper', 'PANGP', 'Zscaler')
                AliasPattern        = 'vpn'
                Description         = 'True when an IP-enabled network adapter description matches a configured VPN client pattern or an interface alias contains the alias pattern.'
            }
        )
    }
}

function Get-ConditionTemplates {
    <#
    .SYNOPSIS
        Returns the condition template document (file override or built-in defaults).

    .DESCRIPTION
        condition-templates.json beside the module overrides the built-in
        defaults when present and parseable; the Deployment Conditions
        options panel writes it. A missing or malformed file falls back to
        defaults so the Package phase never depends on the file existing.
    #>
    $path = Get-ConditionTemplatesPath
    if (Test-Path -LiteralPath $path) {
        try {
            $doc = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($doc -and $doc.Conditions) { return $doc }
            Write-Log ("Condition templates file has no Conditions; using built-in defaults: {0}" -f $path) -Level WARN
        }
        catch {
            Write-Log ("Condition templates file unreadable; using built-in defaults: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
    return Get-DefaultConditionTemplates
}

function Save-ConditionTemplates {
    param([Parameter(Mandatory)][pscustomobject]$Templates)
    $path = Get-ConditionTemplatesPath
    $Templates | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function New-VpnConditionScriptText {
    <#
    .SYNOPSIS
        Builds the client-side detection script for the VPN-connected global condition.

    .DESCRIPTION
        The generated script runs on clients inside the CM agent's
        requirement evaluation, so it must emit exactly one boolean and
        nothing else on the output stream.
    #>
    param(
        [string[]]$AdapterPatterns = @(),
        [string]$AliasPattern = ''
    )

    $quoted = @($AdapterPatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        "'" + (ConvertTo-SingleQuotedContent ([string]$_)) + "'"
    })

    $lines = @(
        ('$patterns = @({0})' -f ($quoted -join ', '))
        '$found = $false'
        'foreach ($adapter in @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=''True''" -ErrorAction SilentlyContinue)) {'
        '    foreach ($p in $patterns) {'
        '        if ($adapter.Description -like (''*'' + $p + ''*'')) { $found = $true }'
        '    }'
        '}'
    )
    if (-not [string]::IsNullOrWhiteSpace($AliasPattern)) {
        $aliasLike = '*' + (ConvertTo-SingleQuotedContent ([string]$AliasPattern)) + '*'
        $lines += @(
            'if (-not $found) {'
            ('    foreach ($ip in @(Get-CimInstance -ClassName MSFT_NetIPAddress -Namespace ''root/StandardCimv2'' -ErrorAction SilentlyContinue)) {')
            ('        if ($ip.InterfaceAlias -like ''{0}'') {{ $found = $true }}' -f $aliasLike)
            '    }'
            '}'
        )
    }
    $lines += '$found'
    return ($lines -join "`r`n")
}

function ConvertFrom-CMEncodedScriptBody {
    <#
    .SYNOPSIS
        Returns the real bytes of a script body stored by the SMS provider.

    .DESCRIPTION
        A signed script is stored as a plain-text preview followed by an
        encoded block whose base64 payload is the imported file byte for byte;
        an unsigned script is stored as plain text with no such block. The
        signature only verifies against those exact bytes, so a read-back
        comparison decodes the block instead of comparing the preview text.

    .OUTPUTS
        [pscustomobject] Bytes, Text, Encoded.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Body)

    $match = [regex]::Match($Body, '#\s*ENCODEDSCRIPT\s*#.*?script block\s*#(?<payload>[A-Za-z0-9+/=\s]+?)#\s*ENCODEDSCRIPT\s*#', 'Singleline, IgnoreCase')
    if ($match.Success) {
        $payload = ($match.Groups['payload'].Value -replace '\s', '')
        try {
            $bytes = [Convert]::FromBase64String($payload)
            return [pscustomobject]@{
                Bytes   = $bytes
                Text    = [System.Text.Encoding]::UTF8.GetString($bytes)
                Encoded = $true
            }
        }
        catch {
            Write-Log ("A stored script carried an encoded block that is not valid base64: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
    return [pscustomobject]@{
        Bytes   = [System.Text.Encoding]::UTF8.GetBytes($Body)
        Text    = $Body
        Encoded = $false
    }
}

function Get-SdmPackageScriptText {
    <#
    .SYNOPSIS
        Extracts one stored script body from an SDMPackageXML document.

    .DESCRIPTION
        A deployment type's detection script is serialized as an Arg named
        ScriptBody, duplicated under CustomData/DetectionScript; a script
        global condition uses the same Arg. DeploymentTypeName selects the
        DeploymentType element by Title, because an application document holds
        every type's body. Nodes are matched by local name so a namespace
        prefix or schema revision does not silently return nothing.

    .OUTPUTS
        [string] the stored body, or $null when the document carries none.
    #>
    param(
        [AllowNull()][string]$SdmPackageXml,
        [AllowEmptyString()][string]$DeploymentTypeName = ''
    )

    if ([string]::IsNullOrWhiteSpace($SdmPackageXml)) { return $null }
    try {
        $document = [xml]$SdmPackageXml
    }
    catch {
        Write-Log ("Could not parse SDMPackageXML for a script read-back: {0}" -f $_.Exception.Message) -Level WARN
        return $null
    }

    $scope = $document
    if (-not [string]::IsNullOrWhiteSpace($DeploymentTypeName)) {
        $escaped = $DeploymentTypeName.Replace("'", "&apos;")
        $node = $document.SelectSingleNode(("//*[local-name()='DeploymentType'][*[local-name()='Title']='{0}']" -f $escaped))
        if (-not $node) {
            # A document holding exactly one type needs no title match; more
            # than one and the wrong body would be verified silently.
            $all = @($document.SelectNodes("//*[local-name()='DeploymentType']"))
            if ($all.Count -eq 1) { $node = $all[0] }
        }
        if (-not $node) { return $null }
        $scope = $node
    }

    foreach ($xpath in @(
            ".//*[local-name()='Arg'][@Name='ScriptBody']",
            ".//*[local-name()='DetectionScript']",
            ".//*[local-name()='ScriptBody']",
            # A script global condition stores its body in this element.
            ".//*[local-name()='DiscoveryScriptBody']")) {
        $found = $scope.SelectSingleNode($xpath)
        if ($found -and -not [string]::IsNullOrWhiteSpace($found.InnerText)) { return [string]$found.InnerText }
    }
    return $null
}

function Get-CMGlobalConditionScriptText {
    <#
    .SYNOPSIS
        Reads back the script body stored on a CM script global condition.

    .DESCRIPTION
        A cmdlet reporting success is not proof the provider stored what was
        sent, so the created condition is read back before it is used. Returns
        $null when the object carries no readable script, which callers treat
        as "cannot compare" rather than as "empty".
    #>
    param([Parameter(Mandatory)]$GlobalCondition)

    $stored = Get-CMGlobalConditionScriptBody -GlobalCondition $GlobalCondition
    if ($null -eq $stored) { return $null }
    return $stored.Text
}

function Get-CMGlobalConditionScriptBody {
    <#
    .SYNOPSIS
        Reads back the decoded script body stored on a CM script global
        condition.

    .OUTPUTS
        [pscustomobject] Bytes, Text, Encoded; $null when none is readable.
    #>
    param([Parameter(Mandatory)]$GlobalCondition)

    $property = $GlobalCondition.PSObject.Properties['SDMPackageXML']
    if (-not $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        try {
            $reloaded = Get-CMGlobalCondition -Id ([string]$GlobalCondition.CI_ID) -ErrorAction Stop
            $property = $reloaded.PSObject.Properties['SDMPackageXML']
        }
        catch {
            return $null
        }
    }
    if (-not $property) { return $null }
    $body = Get-SdmPackageScriptText -SdmPackageXml ([string]$property.Value)
    if ($null -eq $body) { return $null }
    return (ConvertFrom-CMEncodedScriptBody -Body $body)
}

function Test-ConditionScriptTextEqual {
    # The site stores an unsigned script body with LF line ends, whatever
    # was sent. A signed script travels as an encoded block and is compared
    # by bytes instead.
    param([AllowEmptyString()][string]$Left, [AllowEmptyString()][string]$Right)
    return (($Left -replace "`r`n", "`n") -ceq ($Right -replace "`r`n", "`n"))
}

function Find-CMGlobalConditionByName {
    <#
    .SYNOPSIS
        The single site condition matching a name for this template, or $null.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][pscustomobject]$Template
    )

    # A name can match several site conditions (the stock Operating System
    # Language exists once per PlatformType: 1 = Windows, 2 = Mobile).
    # Passing an array to a requirement rule cmdlet's -InputObject fails,
    # so narrow by the template's PlatformType and require a single match.
    $existing = @(Get-CMGlobalCondition -Name $Name)
    if ($existing.Count -gt 1 -and $Template.PSObject.Properties['PlatformType'] -and $Template.PlatformType) {
        $existing = @($existing | Where-Object { [int]$_.PlatformType -eq [int]$Template.PlatformType })
    }
    if ($existing.Count -gt 1) {
        throw "Global condition name '$Name' matches $($existing.Count) site conditions; add a PlatformType to the '$($Template.Id)' template or rename to a unique condition."
    }
    if ($existing.Count -eq 1) { return $existing[0] }
    return $null
}

function Get-ConditionTemplateScriptText {
    <#
    .SYNOPSIS
        The script body of a Script-kind condition template.

    .DESCRIPTION
        The template's own ScriptText when set, otherwise the generated VPN
        detection script. The ConfigMgr global condition and the Intune
        requirement rule use the same text.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Template)

    if ($Template.PSObject.Properties['ScriptText'] -and $Template.ScriptText) {
        return (@($Template.ScriptText) -join "`r`n")
    }
    return (New-VpnConditionScriptText -AdapterPatterns @($Template.AdapterPatterns) -AliasPattern ([string]$Template.AliasPattern))
}

function Get-ConditionScriptContentIdentity {
    <#
    .SYNOPSIS
        Resolves a script condition template to its final text and identity.

    .DESCRIPTION
        The requirement-script signing switch signs the text that the site will
        store, so the identity is computed from the post-signing bytes. A
        required signature that cannot be produced throws here, before any
        site object is created.

    .OUTPUTS
        [pscustomobject] Text, Bytes, Sha256, Signed, Thumbprint, Status.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Template)

    $scriptText = Get-ConditionTemplateScriptText -Template $Template

    $policy = Resolve-CommonSigningPolicy
    $sign = Get-PolicyFlag -Policy $policy -Name 'SignRequirements'
    $require = Get-PolicyFlag -Policy $policy -Name 'RequireRequirements'
    $status = 'NotRequested'
    $thumbprint = ''
    # The Authenticode hash of a script covers its byte-order mark, so the
    # bytes that were signed are the only bytes that verify: the file is
    # written without a BOM and the signed bytes travel to the site verbatim.
    $scriptBytes = [System.Text.Encoding]::UTF8.GetBytes($scriptText)

    if ($sign -or $require) {
        if (-not (Get-Command -Name Invoke-ScriptSigning -ErrorAction SilentlyContinue)) {
            if ($require) { throw "Requirement script signatures are required but the signing service is not available; reinstall AppPackager." }
        }
        else {
            $temporary = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + '.ps1')
            try {
                [System.IO.File]::WriteAllBytes($temporary, $scriptBytes)
                $result = Invoke-ScriptSigning -Path $temporary -Policy $policy -Category Requirements
                $status = [string]$result.Status
                $thumbprint = [string]$result.Thumbprint
                if (@('SignedAndVerified', 'ExistingSignatureValid') -contains $status) {
                    $scriptBytes = [System.IO.File]::ReadAllBytes($temporary)
                    $scriptText = [System.Text.Encoding]::UTF8.GetString($scriptBytes)
                }
                elseif ($require) {
                    throw ("Requirement script signing is required and did not succeed for condition '{0}': {1} ({2})." -f $Template.GlobalConditionName, $status, $result.Reason)
                }
            }
            finally {
                Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash($scriptBytes)
    }
    finally {
        $sha.Dispose()
    }
    return [pscustomobject]@{
        Text       = $scriptText
        Bytes      = $scriptBytes
        Sha256     = (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
        Signed     = (@('SignedAndVerified', 'ExistingSignatureValid') -contains $status)
        Thumbprint = $thumbprint
        Status     = $status
    }
}

function Get-OrCreateGlobalConditionFromTemplate {
    <#
    .SYNOPSIS
        Resolves a condition template to a CM global condition, creating it when absent.

    .DESCRIPTION
        Must be called from the ConfigMgr site drive. Built-in and WQL conditions match
        by name, so an existing site condition of any origin is attached rather
        than duplicated.

        A script condition is a shared site object: changing its body would
        change every application that already references it. A signed or
        changed script therefore never mutates the existing condition; it
        resolves to '<name> (<8 hex of the final text's sha256>)', and the
        condition still referenced by other applications is named in the log.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Template)

    $name = [string]$Template.GlobalConditionName
    if ([string]::IsNullOrWhiteSpace($name)) {
        throw "Condition template '$($Template.Id)' has no GlobalConditionName."
    }

    $kind = [string]$Template.Kind
    if ($kind -ne 'Script') {
        $existing = Find-CMGlobalConditionByName -Name $name -Template $Template
        if ($existing) {
            Write-Log ("Global condition (existing)  : {0}" -f $name)
            return $existing
        }
        switch ($kind) {
            'BuiltIn' {
                throw "Built-in global condition '$name' was not found on the site."
            }
            'Wql' {
                Write-Log ("Global condition (creating)  : {0}" -f $name)
                return New-CMGlobalConditionWqlQuery `
                    -Name $name `
                    -DataType ([string]$Template.DataType) `
                    -Namespace ([string]$Template.Namespace) `
                    -Class ([string]$Template.Class) `
                    -Property ([string]$Template.Property) `
                    -Description ([string]$Template.Description)
            }
            default {
                throw "Condition template '$($Template.Id)' has unsupported Kind '$kind'."
            }
        }
    }

    $identity = Get-ConditionScriptContentIdentity -Template $Template
    $baseCondition = Find-CMGlobalConditionByName -Name $name -Template $Template

    if (-not $identity.Signed) {
        if (-not $baseCondition) {
            Write-Log ("Global condition (creating)  : {0}" -f $name)
            return New-CMScriptGlobalConditionVerified -Name $name -Template $Template -Identity $identity
        }
        $storedText = Get-CMGlobalConditionScriptText -GlobalCondition $baseCondition
        if ($null -eq $storedText -or (Test-ConditionScriptTextEqual -Left $storedText -Right $identity.Text)) {
            Write-Log ("Global condition (existing)  : {0}" -f $name)
            return $baseCondition
        }
    }

    $versionedName = '{0} ({1})' -f $name, $identity.Sha256.Substring(0, 8)
    $versioned = Find-CMGlobalConditionByName -Name $versionedName -Template $Template
    if ($versioned) {
        # The name carries the hash of the script it was created with; the
        # stored script can still have been edited on the site afterwards.
        $storedBody = Get-CMGlobalConditionScriptBody -GlobalCondition $versioned
        $sameScript = $false
        if ($null -ne $storedBody) {
            if ($identity.Signed) { $sameScript = Compare-ByteSequence -Left $storedBody.Bytes -Right ([byte[]]$identity.Bytes) }
            else { $sameScript = Test-ConditionScriptTextEqual -Left ([string]$storedBody.Text) -Right ([string]$identity.Text) }
        }
        if ($null -eq $storedBody -and -not $identity.Signed) {
            Write-Log ("Global condition read-back   : '{0}' stored no readable script body; the site copy could not be compared." -f $versionedName) -Level WARN
        }
        elseif (-not $sameScript) {
            throw ("Global condition '{0}' exists on the site and its stored script is not the script this run built. Delete or rename that condition, then run again." -f $versionedName)
        }
        Write-Log ("Global condition (existing)  : {0}" -f $versionedName)
        return $versioned
    }
    if ($baseCondition) {
        Write-Log ("Global condition unchanged   : '{0}' keeps its current script and its existing application references; this run uses '{1}'." -f $name, $versionedName) -Level WARN
    }
    Write-Log ("Global condition (creating)  : {0}" -f $versionedName)
    return New-CMScriptGlobalConditionVerified -Name $versionedName -Template $Template -Identity $identity
}

function New-CMScriptGlobalConditionVerified {
    <#
    .SYNOPSIS
        Creates a script global condition and verifies what the site stored.

    .DESCRIPTION
        The provider can refuse a write while the cmdlet reports success, and a
        signature that survives locally can still be altered in transport, so
        the stored script is read back and compared, and verified against the
        signing service when the text was signed.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][pscustomobject]$Template,
        [Parameter(Mandatory)][pscustomobject]$Identity
    )

    $temporary = ''
    try {
        $parameters = @{
            Name           = $Name
            DataType       = [string]$Template.DataType
            ScriptLanguage = 'PowerShell'
            Description    = [string]$Template.Description
        }
        $signedBytes = $null
        if ($Identity.PSObject.Properties['Bytes'] -and $null -ne $Identity.Bytes) { $signedBytes = [byte[]]$Identity.Bytes }
        if ($Identity.Signed) {
            # The provider refuses a signed script passed as text; only the
            # file parameter set carries a signature through the import, and
            # only the exact signed bytes verify afterwards.
            if ($null -eq $signedBytes) { $signedBytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Identity.Text) }
            $temporary = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + '.ps1')
            [System.IO.File]::WriteAllBytes($temporary, $signedBytes)
            $parameters['FilePath'] = $temporary
        }
        else {
            $parameters['ScriptText'] = $Identity.Text
        }
        $created = New-CMGlobalConditionScript @parameters
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($temporary)) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }

    $stored = Get-CMGlobalConditionScriptBody -GlobalCondition $created
    if ($null -eq $stored) {
        if ($Identity.Signed) {
            throw ("Global condition '{0}' was created but its stored script could not be read back, so the signed script is not verified." -f $Name)
        }
        Write-Log ("Global condition read-back   : '{0}' stored no readable script body; the site copy could not be compared." -f $Name) -Level WARN
        return $created
    }
    if ($Identity.Signed) {
        if (-not (Compare-ByteSequence -Left $stored.Bytes -Right $signedBytes)) {
            throw ("Global condition '{0}' read back different script bytes than were sent; the site did not store the signed script." -f $Name)
        }
        $expectedThumbprint = ''
        if ($Identity.PSObject.Properties['Thumbprint']) { $expectedThumbprint = [string]$Identity.Thumbprint }
        Test-StoredScriptSignature -Bytes $stored.Bytes -Context ("global condition '$Name'") -ExpectedThumbprint $expectedThumbprint
    }
    elseif (-not (Test-ConditionScriptTextEqual -Left ([string]$stored.Text) -Right ([string]$Identity.Text))) {
        throw ("Global condition '{0}' read back different script content than was sent; the site did not store the finalized script." -f $Name)
    }
    Write-Log ("Global condition read-back   : {0} verified" -f $Name)
    return $created
}

function Test-StoredScriptSignature {
    <#
    .SYNOPSIS
        Verifies that read-back script text still carries a valid signature.

    .DESCRIPTION
        A round trip through XML, JSON or a provider write can re-encode a
        signed script and invalidate its signature while leaving the text
        readable, so verification runs on the bytes that came back. The test is
        integrity, not host trust: the build host is not required to trust the
        signer chain, and an untrusted chain says nothing about the endpoint.
        A hash mismatch, a missing signature, or a different signer is fatal.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes,
        [Parameter(Mandatory)][string]$Context,
        [AllowEmptyString()][string]$ExpectedThumbprint = ''
    )

    if (-not (Get-Command -Name Test-ScriptSignatureBytes -ErrorAction SilentlyContinue)) {
        throw ("Signature verification for {0} was requested but the signing service is not available; reinstall AppPackager." -f $Context)
    }
    $verification = Test-ScriptSignatureBytes -Bytes $Bytes
    # An older signing service reports only Valid, which folds trust into the result.
    $intact = if ($verification.PSObject.Properties['SignatureIntact']) { [bool]$verification.SignatureIntact } else { [bool]$verification.Valid }
    if (-not $intact) {
        throw ("The signature on the stored script for {0} did not verify after read-back: {1} ({2})." -f $Context, $verification.Status, $verification.Reason)
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedThumbprint) -and
        $verification.Thumbprint -ne $ExpectedThumbprint) {
        throw ("The stored script for {0} is signed by {1}, not by the certificate this build signed with ({2})." -f $Context, $verification.Thumbprint, $ExpectedThumbprint)
    }
    # Chain trust on this host is recorded, never required: under AllSigned the signer
    # must be trusted as a publisher on the client, which is a client configuration.
    $trusted = ($verification.PSObject.Properties['TrustedOnThisHost'] -and $verification.TrustedOnThisHost)
    $trustState = if ($trusted) { 'chain trusted on this host' } else { ("chain not trusted on this host: {0}" -f $verification.Status) }
    Write-Log ("Signature verified           : {0} ({1}); {2}; endpoint publisher trust is a separate client configuration." -f $Context, $verification.Thumbprint, $trustState)
}

function Assert-ScriptSignatureBeforeSend {
    <#
    .SYNOPSIS
        Refuses to send a script that must be signed when its bytes carry no
        intact signature.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes,
        [Parameter(Mandatory)][string]$Context
    )

    if (-not (Get-Command -Name Test-ScriptSignatureBytes -ErrorAction SilentlyContinue)) {
        throw ("A signature check for the {0} was requested but the signing service is not available; reinstall AppPackager." -f $Context)
    }
    $verification = Test-ScriptSignatureBytes -Bytes $Bytes
    $intact = if ($verification.PSObject.Properties['SignatureIntact']) { [bool]$verification.SignatureIntact } else { [bool]$verification.Valid }
    if (-not $intact) {
        throw ("The {0} carries no intact signature ({1}) and was not sent. Re-run the Stage phase with signing enabled." -f $Context, $verification.Status)
    }
}

function Compare-ByteSequence {
    <#
    .SYNOPSIS
        True when two byte arrays are identical in length and content.
    #>
    param(
        [AllowNull()][byte[]]$Left,
        [AllowNull()][byte[]]$Right
    )
    if ($null -eq $Left -or $null -eq $Right) { return $false }
    if ($Left.Length -ne $Right.Length) { return $false }
    for ($i = 0; $i -lt $Left.Length; $i++) {
        if ($Left[$i] -ne $Right[$i]) { return $false }
    }
    return $true
}

function Get-DeploymentTypeRequirementSpecs {
    <#
    .SYNOPSIS
        Collects requirement rule specs for the deployment type being created.

    .DESCRIPTION
        Two sources merge: a manifest Requirements array and the
        APP_PACKAGER_REQUIREMENTS environment JSON the GUI sets per app
        ({ SchemaVersion, Rules: [...] }). Malformed environment JSON
        throws instead of packaging without the rules the operator
        configured.
    #>
    param(
        [pscustomobject]$Manifest,
        [switch]$IgnoreEnvironment
    )

    $specs = @()
    if ($Manifest -and $Manifest.PSObject.Properties['Requirements'] -and $Manifest.Requirements) {
        $specs += @($Manifest.Requirements)
    }

    $envJson = $(if ($IgnoreEnvironment) { $null } else { $env:APP_PACKAGER_REQUIREMENTS })
    if (-not [string]::IsNullOrWhiteSpace($envJson)) {
        try {
            $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "APP_PACKAGER_REQUIREMENTS is not valid JSON: $($_.Exception.Message)"
        }
        if ($parsed -and $parsed.PSObject.Properties['Rules'] -and $parsed.Rules) {
            $specs += @($parsed.Rules)
        }
    }

    # No comma wrap: an empty array must unroll to zero pipeline objects
    # so callers' @(...) sees Count 0, not a one-element array-of-array.
    return $specs
}

function New-DeploymentTypeRequirementRules {
    <#
    .SYNOPSIS
        Builds CM requirement rule objects from condition templates and specs.

    .DESCRIPTION
        Must be called from the ConfigMgr site drive. Unknown condition ids and
        unmapped values throw so a package run never creates a deployment
        type missing rules the operator asked for.

    .OUTPUTS
        Requirement rule array; empty when no specs are configured.
    #>
    param(
        [pscustomobject]$Manifest,
        [switch]$IgnoreEnvironment
    )

    $specs = @(Get-DeploymentTypeRequirementSpecs -Manifest $Manifest -IgnoreEnvironment:$IgnoreEnvironment)
    if ($specs.Count -eq 0) { return @() }

    $doc = Get-ConditionTemplates
    $rules = @()
    foreach ($spec in $specs) {
        $condId = [string]$spec.ConditionId
        $found = @($doc.Conditions | Where-Object { [string]$_.Id -eq $condId })
        if ($found.Count -eq 0) {
            throw "Requirement rule references unknown condition id '$condId'."
        }
        $template = $found[0]
        $gc = Get-OrCreateGlobalConditionFromTemplate -Template $template

        switch ([string]$template.RuleType) {
            'CommonValue' {
                $value1 = [string]$spec.Value
                if ($template.PSObject.Properties['Values'] -and $template.Values) {
                    $mapped = $template.Values.PSObject.Properties[$value1]
                    if (-not $mapped) {
                        $known = (@($template.Values.PSObject.Properties | ForEach-Object { $_.Name }) -join ', ')
                        throw "Condition '$condId' has no mapping for value '$value1' (known: $known)."
                    }
                    $value1 = [string]$mapped.Value
                }
                if ([string]::IsNullOrWhiteSpace($value1)) {
                    throw "Condition '$condId' requires a Value."
                }
                $rules += New-CMRequirementRuleCommonValue -InputObject $gc -RuleOperator IsEquals -Value1 $value1
            }
            'OperatingSystemLanguage' {
                $cultures = @()
                foreach ($c in @($spec.Cultures)) {
                    if ([string]::IsNullOrWhiteSpace([string]$c)) { continue }
                    $cultures += [System.Globalization.CultureInfo]::GetCultureInfo([string]$c)
                }
                if ($cultures.Count -eq 0) {
                    throw "Condition '$condId' requires at least one culture code (for example de-DE)."
                }
                $rules += New-CMRequirementRuleOperatingSystemLanguageValue -InputObject $gc -RuleOperator OneOf -Culture $cultures
            }
            'Boolean' {
                # [bool]'false' is $true; Convert.ToBoolean parses the
                # string form and throws on anything else.
                $boolValue = $false
                if ($spec.Value -is [bool]) { $boolValue = $spec.Value }
                else { $boolValue = [System.Convert]::ToBoolean([string]$spec.Value) }
                $rules += New-CMRequirementRuleBooleanValue -InputObject $gc -Value $boolValue
            }
            default {
                throw "Condition '$condId' has unsupported RuleType '$($template.RuleType)'."
            }
        }
        Write-Log ("Requirement rule             : {0} ({1})" -f $condId, $template.GlobalConditionName)
    }

    return $rules
}


function Get-ManifestDeploymentTypeSpecs {
    <#
    .SYNOPSIS
        Resolves a stage manifest into an ordered deployment type spec list.

    .DESCRIPTION
        A manifest without a DeploymentTypes array yields one spec built
        from the base fields (today's single-DT behavior, byte for byte).
        A DeploymentTypes array yields one spec per entry in manifest
        order; CM assigns deployment type priority by creation order and
        the client installs the first deployment type whose requirements
        pass, so authors list the most specific variant first and the
        unconditional fallback last. Entry fields override the base
        manifest field of the same name; ContentSubpath is relative to
        the network content root. On a multi-DT manifest the
        APP_PACKAGER_REQUIREMENTS environment JSON is ignored (per-DT
        Requirements arrays are authoritative) - the caller logs that.

    .OUTPUTS
        [pscustomobject[]] DtName, ContentLocation, InstallCommand,
        UninstallCommand, Detection, RequirementSource, plus behavior
        overrides (PostExecutionBehavior, InstallationBehaviorType,
        LogonRequirementType, RequireUserInteraction).
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [Parameter(Mandatory)][string]$AppName
    )

    $entries = @()
    if ($Manifest.PSObject.Properties['DeploymentTypes'] -and $Manifest.DeploymentTypes) {
        $entries = @($Manifest.DeploymentTypes)
    }

    $base = {
        param($Entry, $Field)
        if ($Entry -and $Entry.PSObject.Properties[$Field] -and $null -ne $Entry.$Field -and ('' -ne [string]$Entry.$Field -or $Entry.$Field -is [bool])) { return $Entry.$Field }
        if ($Manifest.PSObject.Properties[$Field]) { return $Manifest.$Field }
        return $null
    }

    if ($entries.Count -eq 0) {
        # RequirementSource $Manifest keeps the single-DT contract: manifest
        # Requirements plus the environment JSON, exactly as before.
        return ,([pscustomobject]@{
            DtName                   = $AppName
            ContentLocation          = $NetworkContentPath
            InstallCommand           = $(if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.InstallCommandLine)) { [string]$Manifest.InstallCommandLine } else { 'install.bat' })
            UninstallCommand         = $(if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.UninstallCommandLine)) { [string]$Manifest.UninstallCommandLine } else { 'uninstall.bat' })
            Detection                = $Manifest.Detection
            RequirementSource        = $Manifest
            Timing                   = $Manifest.Timing
            PostExecutionBehavior    = $Manifest.PostExecutionBehavior
            InstallationBehaviorType = $Manifest.InstallationBehaviorType
            LogonRequirementType     = $Manifest.LogonRequirementType
            RequireUserInteraction   = ($Manifest.RequireUserInteraction -eq $true)
        })
    }

    $specs = @()
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $entries) {
        $suffix = [string]$entry.NameSuffix
        if ([string]::IsNullOrWhiteSpace($suffix)) {
            throw "Every DeploymentTypes entry requires a NameSuffix; the deployment type name becomes '<AppName> - <NameSuffix>'."
        }
        $dtName = "$AppName - $suffix"
        if (-not $seen.Add($dtName)) {
            throw "DeploymentTypes entries produce a duplicate deployment type name '$dtName'."
        }

        $content = $NetworkContentPath
        if (-not [string]::IsNullOrWhiteSpace([string]$entry.ContentSubpath)) {
            $content = Join-Path $NetworkContentPath ([string]$entry.ContentSubpath)
        }

        $detection = & $base $entry 'Detection'
        if (-not $detection) {
            throw "DeploymentTypes entry '$suffix' has no Detection and the manifest has no base Detection."
        }

        # Requirements resolve strictly from the entry: a variant without
        # rules is the unconditional fallback by design, and inheriting the
        # base or environment rules here would gate the fallback by accident.
        # @() over a missing property yields @($null) with one element; the
        # Where-Object keeps a rule-less entry a true unconditional fallback.
        $reqSource = [pscustomobject]@{ Requirements = @(@($entry.Requirements) | Where-Object { $null -ne $_ }) }

        $installOverride = [string](& $base $entry 'InstallCommandLine')
        $uninstallOverride = [string](& $base $entry 'UninstallCommandLine')
        $specs += [pscustomobject]@{
            DtName                   = $dtName
            ContentLocation          = $content
            InstallCommand           = $(if (-not [string]::IsNullOrWhiteSpace($installOverride)) { $installOverride } else { 'install.bat' })
            UninstallCommand         = $(if (-not [string]::IsNullOrWhiteSpace($uninstallOverride)) { $uninstallOverride } else { 'uninstall.bat' })
            Detection                = $detection
            RequirementSource        = $reqSource
            Timing                   = (& $base $entry 'Timing')
            PostExecutionBehavior    = (& $base $entry 'PostExecutionBehavior')
            InstallationBehaviorType = (& $base $entry 'InstallationBehaviorType')
            LogonRequirementType     = (& $base $entry 'LogonRequirementType')
            RequireUserInteraction   = ((& $base $entry 'RequireUserInteraction') -eq $true)
        }
    }
    return $specs
}

function Resolve-DetectionScriptTransport {
    <#
    .SYNOPSIS
        Chooses how a script detection reaches Add-CMScriptDeploymentType.

    .DESCRIPTION
        A finalized detection names a file under the content root; that file
        carries the exact bytes the signature covers, so it is imported with
        -ScriptFile rather than re-serialized text. A file that is not
        reachable from this host is copied byte for byte to a temporary path
        first; re-reading it as text would re-encode it and invalidate the
        signature. A manifest with no ScriptFile keeps today's -ScriptText.

    .OUTPUTS
        [pscustomobject] Transport ('File'|'Text'), Path, TempPath, Text,
        Bytes, Signed.
    #>
    param(
        [Parameter(Mandatory)]$Detection,
        [Parameter(Mandatory)][string]$ContentLocation,
        [AllowEmptyString()][string]$SignatureStatus = ''
    )

    $signed = @('SignedAndVerified', 'ExistingSignatureValid') -contains $SignatureStatus
    $requireSigned = Get-PolicyFlag -Policy (Resolve-CommonSigningPolicy) -Name 'RequireDetection'

    $scriptFile = ''
    if ($Detection.PSObject.Properties['ScriptFile']) { $scriptFile = [string]$Detection.ScriptFile }
    if ([string]::IsNullOrWhiteSpace($scriptFile)) {
        if ($signed) {
            # The provider rejects a signed script supplied as text outright,
            # so a signed detection without a finalized file has no transport.
            throw 'The detection script is signed but the manifest names no Detection.ScriptFile; a signed script cannot be imported as script text. Re-run the Stage phase so the finalized script is written to the content folder.'
        }
        $text = [string]$Detection.ScriptText
        if ($requireSigned) {
            Assert-ScriptSignatureBeforeSend -Bytes ([System.Text.Encoding]::UTF8.GetBytes($text)) -Context 'detection script'
        }
        return [pscustomobject]@{
            Transport = 'Text'; Path = ''; TempPath = ''
            Text      = $text
            Bytes     = [System.Text.Encoding]::UTF8.GetBytes($text)
            Signed    = $signed
        }
    }

    $absolute = $scriptFile
    if (-not [System.IO.Path]::IsPathRooted($absolute)) {
        $absolute = Join-Path $ContentLocation $scriptFile
    }
    # After Connect-CMSite the session sits on the CMSite drive, where a bare
    # UNC path resolves through the CM provider and reports missing files as
    # absent; the FileSystem:: qualifier keeps the check on the filesystem.
    if (-not (Test-Path -LiteralPath ("FileSystem::" + $absolute) -PathType Leaf)) {
        throw ("Detection script '{0}' was not found under the content location '{1}'; re-run the Stage phase." -f $scriptFile, $ContentLocation)
    }
    $bytes = [System.IO.File]::ReadAllBytes($absolute)
    # The status in the manifest is from stage time. The file can have changed
    # since, and these are the bytes the deployment type receives.
    if ($signed -or $requireSigned) { Assert-ScriptSignatureBeforeSend -Bytes $bytes -Context ("detection script '$scriptFile'") }
    # The signature block counts toward the site's script size limit, so the
    # check runs on the finalized bytes rather than the authored text.
    if (Get-Command -Name Get-ConfigMgrDetectionScriptMaxBytes -ErrorAction SilentlyContinue) {
        $maxBytes = [int](Get-ConfigMgrDetectionScriptMaxBytes)
        if ($maxBytes -gt 0 -and $bytes.Length -gt $maxBytes) {
            throw ("The finalized detection script is {0} bytes, above the {1}-byte limit a deployment type accepts; shorten the detector or use a native detection rule." -f $bytes.Length, $maxBytes)
        }
    }

    $temporary = ''
    $importPath = $absolute
    if ($absolute -like '\\*') {
        # The site server reads -ScriptFile from the console host, so a share
        # the console cannot reach is staged locally with identical bytes.
        $temporary = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + '.ps1')
        [System.IO.File]::WriteAllBytes($temporary, $bytes)
        $importPath = $temporary
    }

    return [pscustomobject]@{
        Transport = 'File'
        Path      = $importPath
        TempPath  = $temporary
        Text      = [System.Text.Encoding]::UTF8.GetString($bytes)
        Bytes     = $bytes
        Signed    = $signed
    }
}

function Get-CMDeploymentTypeDetectionScript {
    <#
    .SYNOPSIS
        Reads back the detection script a deployment type stored.

    .DESCRIPTION
        The read-back reads the application document, not the deployment type
        object: Get-CMDeploymentType returns the first type's SDMPackageXML for
        every type of an application, so a multi-type app would verify the
        wrong body. The stored body is decoded, because a signed script is
        carried as an encoded block rather than as its own text.

    .OUTPUTS
        [pscustomobject] Bytes, Text, Encoded; $null when none is readable.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$DeploymentTypeName
    )

    try {
        $application = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
    }
    catch {
        Write-Log ("Detection read-back unavailable: {0}" -f $_.Exception.Message) -Level WARN
        return $null
    }
    $application = @($application) | Select-Object -First 1
    if (-not $application) { return $null }
    $property = $application.PSObject.Properties['SDMPackageXML']
    if (-not $property) { return $null }

    $body = Get-SdmPackageScriptText -SdmPackageXml ([string]$property.Value) -DeploymentTypeName $DeploymentTypeName
    if ($null -eq $body) { return $null }
    return ConvertFrom-CMEncodedScriptBody -Body $body
}

function Test-StoredDetectionScript {
    <#
    .SYNOPSIS
        Verifies the detection script the site stored against what was sent.

    .DESCRIPTION
        A CM cmdlet reports success without the provider having accepted the
        write, and a signed script can survive import as readable text while
        its signature no longer verifies. A signed detection that does not
        come back intact stops the run rather than deploying a detector the
        client will refuse under AllSigned.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$DeploymentTypeName,
        [Parameter(Mandatory)][pscustomobject]$Transport,
        [AllowEmptyString()][string]$ExpectedThumbprint = ''
    )

    try {
        $stored = Get-CMDeploymentTypeDetectionScript -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName
        $require = Get-PolicyFlag -Policy (Resolve-CommonSigningPolicy) -Name 'RequireDetection'
        if ($null -eq $stored) {
            if ($Transport.Signed -or $require) {
                throw ("The detection script stored on deployment type '{0}' could not be read back, so its signature cannot be verified." -f $DeploymentTypeName)
            }
            Write-Log ("Detection read-back          : '{0}' stored no readable script body; not compared." -f $DeploymentTypeName) -Level WARN
            return
        }
        if ($stored.Encoded -or $Transport.Signed) {
            # A signature covers bytes, so a signed import is compared byte for
            # byte rather than through a text normalization.
            if (-not (Compare-ByteSequence -Left $stored.Bytes -Right $Transport.Bytes)) {
                throw ("Deployment type '{0}' read back different detection script bytes than were sent." -f $DeploymentTypeName)
            }
        }
        elseif (([string]$stored.Text).TrimEnd("`r", "`n") -ne ([string]$Transport.Text).TrimEnd("`r", "`n")) {
            throw ("Deployment type '{0}' read back different detection script content than was sent." -f $DeploymentTypeName)
        }
        if ($Transport.Signed -or $require) {
            Test-StoredScriptSignature -Bytes $stored.Bytes -Context ("deployment type '$DeploymentTypeName'") -ExpectedThumbprint $ExpectedThumbprint
        }
        Write-Log ("Detection read-back          : {0} verified" -f $DeploymentTypeName)
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace([string]$Transport.TempPath)) {
            Remove-Item -LiteralPath $Transport.TempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-ResolvedDeploymentCommand {
    <#
    .SYNOPSIS
        Re-checks the resolved deployment-type command lines for execution-policy
        relaxation at Package time.

    .DESCRIPTION
        The stage-time launcher check cannot see operator overrides or per-variant
        commands resolved later. This runs the same check over the final commands and
        throws before any site mutation whenever the deployment category is signed or
        the policy requires deployment signatures. Without the signing module the build
        is a legacy unsigned build and nothing is enforced.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Specs,
        $Manifest
    )

    $checker = Get-Command -Name Test-DeploymentLauncherChain -ErrorAction SilentlyContinue
    if (-not $checker) { return $null }

    $policy = $null
    if (Get-Command -Name Get-SigningPolicy -ErrorAction SilentlyContinue) { $policy = Get-SigningPolicy }

    $deploymentStatus = ''
    if ($Manifest -and $Manifest.PSObject.Properties['ScriptSigning'] -and $Manifest.ScriptSigning -and
        $Manifest.ScriptSigning.PSObject.Properties['Deployment'] -and $Manifest.ScriptSigning.Deployment) {
        $deploymentStatus = [string]$Manifest.ScriptSigning.Deployment.Status
    }
    $signed = @('SignedAndVerified', 'ExistingSignatureValid') -contains $deploymentStatus
    $required = [bool]($policy -and ($policy.SignDeployment -or $policy.RequireDeployment))
    if (-not ($signed -or $required)) { return $null }

    # A synthetic manifest carries only the resolved commands: the stage tree was
    # already inspected during finalization and is not re-read here.
    $commandManifest = [pscustomobject]@{
        InstallCommandLine   = ''
        UninstallCommandLine = ''
        DeploymentTypes      = @($Specs | ForEach-Object {
            [pscustomobject]@{
                InstallCommand   = [string]$_.InstallCommand
                UninstallCommand = [string]$_.UninstallCommand
            }
        })
    }
    # Enforcement follows the manifest's own signed state as well as the policy, so a
    # signed build is refused even when the child runs without a policy in scope.
    $effectivePolicy = if ($policy) { $policy } else { [pscustomobject]@{ SignDeployment = $false; RequireDeployment = $false } }
    if ($signed) { $effectivePolicy = [pscustomobject]@{ SignDeployment = $true; RequireDeployment = $true } }

    # The stage root is deliberately a path that cannot exist: only the commands are
    # inspected here, the staged files were covered at finalization time.
    $commandsOnlyRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ap-commands-only-' + [guid]::NewGuid().ToString('N'))
    return Test-DeploymentLauncherChain -StageRoot $commandsOnlyRoot -Manifest $commandManifest -Policy $effectivePolicy
}

function Get-CMApplicationCurrentCIId {
    # Every saved revision of an application has its own CI_ID, so an object
    # captured before a later Set-CMApplication or deployment type write names
    # a revision that the history cleanup removes.
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][UInt32]$Fallback)

    $current = @(Get-CMApplication -Name $Name -DisableWildcardHandling -ErrorAction SilentlyContinue)
    if ($current.Count -eq 1 -and $current[0].CI_ID) { return [UInt32]$current[0].CI_ID }
    Write-Log ("Application CI_ID            : the final read of '{0}' returned {1} object(s); reporting the CI_ID from before the revision cleanup ({2})." -f $Name, $current.Count, $Fallback) -Level WARN
    return $Fallback
}

function New-MECMApplicationFromManifest {
    <#
    .SYNOPSIS
        Creates a ConfigMgr application with Script deployment type from a stage manifest.

    .DESCRIPTION
        Reads a stage manifest object and creates a ConfigMgr application with a single
        Script deployment type. Supports all detection methods:

          RegistryKeyValue  Single registry value comparison (IsEquals or GreaterEquals)
          RegistryKey       Registry key existence check
          File              File existence or version comparison
          Script            PowerShell script-based detection
          Compound          Multiple clauses joined by AND or OR

        Handles ConfigMgr site connection, duplicate app check, New-CMApplication with
        task-sequence eligibility, detection clause creation, Add-CMScriptDeploymentType
        with gold standard parameters, optional PostExecutionBehavior, and
        revision history cleanup.

        Backward compatible: manifests without a Detection.Type field default
        to RegistryKeyValue; DisplayVersion is accepted as an alias for
        ExpectedValue.

    .OUTPUTS
        [UInt32] CI_ID of the created or already-complete application.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$SiteCode,
        [AllowEmptyString()][string]$Comment = '',
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30,
        # Empty means unset: Resolve-OnExistingBehavior falls through to the
        # environment variable and then to Skip.
        [AllowEmptyString()][string]$OnExisting = ''
    )

    $orig = Get-Location

    $verifyContent = {
        $contentVerification = Compare-StageFileHashes -Root $NetworkContentPath -Expected $Manifest.FileHashes
        if ($contentVerification.Skipped) {
            Write-Log ("Package integrity verification : skipped ({0})" -f $contentVerification.Reason) -Level WARN
        }
        elseif (-not $contentVerification.Pass) {
            throw ("Package integrity verification failed: {0}" -f (Format-StageFileHashComparison -Comparison $contentVerification))
        }
        else {
            Write-Log ("Package integrity verified     : {0} file(s)" -f $contentVerification.ExpectedCount)
        }
    }

    $deferredSync = $script:DeferredContentSync
    $script:DeferredContentSync = $null
    if ($deferredSync -and -not [string]::Equals([string]$deferredSync.NetworkContentPath, $NetworkContentPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $deferredSync = $null
    }
    if (-not $deferredSync) { & $verifyContent }

    # Tracks the operation in flight so the catch block can name the step
    # that failed. CM cmdlet errors (e.g. "Key cannot be null. Parameter
    # name: key") rarely identify their own call site.
    $step = 'initialization'

    try {
        $step = "Connect-CMSite (SiteCode=$SiteCode)"
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "ConfigMgr site connection failed."
        }

        $appName = Get-PackagedApplicationName -AppName $Manifest.AppName -Version $Manifest.SoftwareVersion
        if ([string]::IsNullOrWhiteSpace([string]$appName)) {
            throw "Stage manifest AppName is null or empty; cannot create a ConfigMgr application. Re-run the Stage phase and verify the manifest."
        }

        Write-Log ("Manifest fields              : AppName='{0}' Publisher='{1}' SoftwareVersion='{2}' DetectionType='{3}'" -f $appName, $Manifest.Publisher, $Manifest.SoftwareVersion, $Manifest.Detection.Type) -Level DEBUG

        # The signing service records what it achieved per category; a signed
        # detector must be imported and read back as bytes, not as text.
        $detectionSignatureStatus = ''
        $detectionSignatureThumbprint = ''
        if ($Manifest.PSObject.Properties['ScriptSigning'] -and $Manifest.ScriptSigning -and
            $Manifest.ScriptSigning.PSObject.Properties['Detection'] -and $Manifest.ScriptSigning.Detection) {
            $detectionSignatureStatus = [string]$Manifest.ScriptSigning.Detection.Status
            if ($Manifest.ScriptSigning.Detection.PSObject.Properties['Thumbprint']) {
                $detectionSignatureThumbprint = [string]$Manifest.ScriptSigning.Detection.Thumbprint
            }
        }

        # The version rule and the supersedence mode (App Packager Desktop rules). A refused plan creates nothing: it
        # throws before the first write to the site. The plan carries the manifest with the version rule written.
        $step = 'Supersedence plan'
        $supersedenceContentPath = $NetworkContentPath
        $supersedencePlan = Get-SupersedencePlan -Manifest $Manifest -Settings (Get-SupersedenceSettings) -CommandLine (Get-SupersedenceCommandLine) `
            -PublishesToConfigMgr $true -PublishesToIntune $false `
            -MsiProductVersion { param($m) Get-SupersedenceMsiProductVersion -Manifest $m -ContentPath $supersedenceContentPath }.GetNewClosure() `
            -FallbackApplicationId (Get-SupersedenceFallbackApplicationId)
        foreach ($supersedenceLine in $supersedencePlan.Log) { Write-Log $supersedenceLine }
        if ($supersedencePlan.Refusal) { throw $supersedencePlan.Refusal }
        $Manifest = $supersedencePlan.ConfigMgrManifest

        $step = 'Deployment type spec resolution'
        $dtSpecs = @(Get-ManifestDeploymentTypeSpecs -Manifest $Manifest -NetworkContentPath $NetworkContentPath -AppName $appName)
        $contentOptions = Get-DeploymentTypeContentOptions
        Write-Log ("Deployment type content      : fallback={0}, slow network={1}" -f $(if ($contentOptions.ContentFallback) { 'Allow' } else { 'Deny' }), $contentOptions.SlowNetworkDeploymentMode)
        # AutoInstall is application-wide. Every resolved DT must run as
        # system without needing a logged-on user or desktop interaction.
        # Missing overrides use the same defaults as dtParams below.
        $allowTaskSequence = @($dtSpecs | Where-Object {
            ($_.InstallationBehaviorType -and $_.InstallationBehaviorType -ne 'InstallForSystem') -or
            ($_.LogonRequirementType -eq 'OnlyWhenUserLoggedOn') -or
            ($_.RequireUserInteraction -eq $true)
        }).Count -eq 0
        $isMultiDt = ($dtSpecs.Count -gt 1 -or $dtSpecs[0].DtName -ne $appName)
        if ($isMultiDt) {
            Write-Log ("Deployment types (manifest)  : {0}" -f (($dtSpecs | ForEach-Object { $_.DtName }) -join ', '))
            if (-not [string]::IsNullOrWhiteSpace($env:APP_PACKAGER_REQUIREMENTS)) {
                Write-Log "APP_PACKAGER_REQUIREMENTS ignored: a DeploymentTypes manifest carries its own per-deployment-type Requirements." -Level WARN
            }
        }

        # Operator command overrides replace the deployment type commands
        # on single-DT apps; a multi-DT manifest carries per-variant
        # commands, so the flat override is refused rather than applied to
        # every variant.
        $commandOverrides = Get-RequestedCommandOverrides
        if ($commandOverrides) {
            if ($isMultiDt) {
                Write-Log "APP_PACKAGER_COMMANDS ignored: a DeploymentTypes manifest carries per-deployment-type commands." -Level WARN
            }
            else {
                if (-not [string]::IsNullOrWhiteSpace($commandOverrides.Install)) {
                    $dtSpecs[0].InstallCommand = $commandOverrides.Install
                    Write-Log ("Install command (override)   : {0}" -f $commandOverrides.Install)
                }
                if (-not [string]::IsNullOrWhiteSpace($commandOverrides.Uninstall)) {
                    $dtSpecs[0].UninstallCommand = $commandOverrides.Uninstall
                    Write-Log ("Uninstall command (override) : {0}" -f $commandOverrides.Uninstall)
                }
            }
        }

        # Stage-time signing inspected the stage tree; the commands that reach the
        # deployment types can still change here (overrides, per-variant specs), so
        # the resolved command lines are re-checked before the first site write.
        [void](Test-ResolvedDeploymentCommand -Specs $dtSpecs -Manifest $Manifest)

        # The chain check warns and never blocks. It runs only when the application carries the marker that finds the old one.
        $supersedenceActive = ($supersedencePlan.Mode -ne 'Off' -and -not [string]::IsNullOrWhiteSpace([string]$supersedencePlan.Marker))
        if ($supersedenceActive) {
            $step = 'Supersedence chain check'
            foreach ($warning in @(Invoke-SupersedenceChainCheck -NewTitle $appName -NewVersion ([string]$Manifest.SoftwareVersion) -Marker $supersedencePlan.Marker `
                        -NewRule $supersedencePlan.ConfigMgrRule -Mode $supersedencePlan.Mode -NewManifest $Manifest)) {
                Write-Log ("Chain check                  : {0}" -f $warning) -Level WARN
            }
        }

        $step = "Get-CMApplication duplicate check ('$appName')"
        $existing = Get-CMApplication -Name $appName -DisableWildcardHandling -ErrorAction Stop
        $cmApp = $null
        $replaceDtNames = @()
        $replaceExisting = $false
        if ($existing) {
            $existingApps = @($existing)
            if ($existingApps.Count -gt 1) {
                throw "Multiple existing ConfigMgr applications matched '$appName'; refusing to package until the duplicate names are resolved."
            }
            $cmApp = $existingApps[0]

            $missingDts = @($dtSpecs | Where-Object { -not (Test-MECMApplicationHasDeploymentType -ApplicationName $appName -DeploymentTypeName $_.DtName) })
            $existingVersion = [string]$cmApp.SoftwareVersion

            # All name collisions require a policy, including browser updates
            # using a perpetual (versionless) application name.
            $existingPolicy = Resolve-OnExistingBehavior -Requested $OnExisting
            Write-Log ("On-existing behavior         : {0} (source: {1})" -f $existingPolicy.Behavior, $existingPolicy.Source)
            if ($existingPolicy.Behavior -ne 'Overwrite') {
                if ($existingPolicy.Behavior -eq 'Fail') {
                    throw ("Existing ConfigMgr application '$appName' is already at version $existingVersion and OnExisting=Fail was requested.")
                }

                if ($existingPolicy.Behavior -eq 'Skip') {
                    if ($missingDts.Count -eq 0) {
                        # Repair the old unconditional AutoInstall setting even
                        # when the same-version deployment types are retained.
                        if (-not $allowTaskSequence -and $cmApp.AutoInstall -ne $false) {
                            $step = "Set-CMApplication task-sequence eligibility ('$appName')"
                            Set-CMApplication -Name $appName -AutoInstall $false -ErrorAction Stop | Out-Null
                            Write-Log "Task-sequence install        : disabled (deployment type requires user context or interaction)"
                        }
                        Write-Log "Application already exists    : $appName (v$existingVersion, unchanged)" -Level WARN
                        if ($deferredSync) {
                            Write-Log "Network content              : unchanged; this stage differs from the content of the existing application and was not copied" -Level WARN
                        }
                        Write-Log ("Deployment type(s) validated : {0}" -f (($dtSpecs | ForEach-Object { $_.DtName }) -join ', '))
                        # Machine-readable companion to the line above: the GUI
                        # parses it to offer the operator an overwrite re-run.
                        Write-Log ("{0} app='{1}' version='{2}'" -f $script:OnExistingConflictMarker, $appName, $existingVersion)
                        # An idempotent rerun still applies the icon, so a package
                        # created before icon support (or before a pack install)
                        # gains one without a version change.
                        Set-CMApplicationIconFromManifest -Manifest $Manifest -AppName $appName -NetworkContentPath $NetworkContentPath
                        # An idempotent rerun runs the supersedence step again: it reads the relations first and adds the missing ones.
                        if ($supersedenceActive) { [void](Invoke-SupersedenceStepAndLog -ApplicationName $appName -Version ([string]$Manifest.SoftwareVersion) -Plan $supersedencePlan) }
                        return (Get-CMApplicationCurrentCIId -Name $appName -Fallback ([UInt32]$cmApp.CI_ID))
                    }
                    throw ("Existing ConfigMgr application '$appName' is missing deployment type(s): {0}. This looks like a partial prior package run; fix or remove the partial app before packaging again." -f (($missingDts | ForEach-Object { $_.DtName }) -join ', '))
                }

            }
            else {
                # Version change on a reused application name (version-less CMName by
                # design): replace every deployment type with the new set pointing at
                # the new content.
                Write-Log "Application already exists    : $appName (v$existingVersion -> v$($Manifest.SoftwareVersion), replacing deployment types, operator choice)" -Level WARN
                $replaceExisting = $true
            }

            # New deployment types are created under staging names first,
            # because a deployed application refuses to remove its last
            # deployment type.
            if ($replaceExisting) {
                $step = "Get-CMDeploymentType inventory ('$appName')"
                $replaceDtNames = @(Get-CMDeploymentType -ApplicationName $appName -ErrorAction Stop | ForEach-Object { [string]$_.LocalizedDisplayName })
                if ($replaceDtNames.Count -eq 0) {
                    Write-Log "Existing app has no deployment type; adding the new set." -Level WARN
                }
            }
        }

        if ($deferredSync) {
            # Connect-CMSite leaves the CM drive current; UNC paths resolve
            # through the CM provider there and Test-Path fails on them.
            $step = "Network content copy ('$NetworkContentPath')"
            Push-Location -LiteralPath $env:SystemRoot
            try {
                Sync-StagedContentToNetwork -LocalContentPath $deferredSync.LocalContentPath -NetworkContentPath $deferredSync.NetworkContentPath -Manifest $deferredSync.Manifest -Exclude $deferredSync.Exclude -NoDefer
                & $verifyContent
            }
            finally { Pop-Location }
        }

        # Requirement rules resolve for every deployment type before any
        # create/replace so a bad spec fails the run without leaving a
        # partial application behind. Rule objects must be built on the CM
        # drive; they survive the later detection-clause drive switch. On a
        # multi-DT manifest the per-entry Requirements are authoritative and
        # the environment JSON is excluded.
        $step = 'Requirement rules'
        foreach ($spec in $dtSpecs) {
            $spec | Add-Member -NotePropertyName RequirementRules -NotePropertyValue @(New-DeploymentTypeRequirementRules -Manifest $spec.RequirementSource -IgnoreEnvironment:$isMultiDt)
        }

        # Clear the incompatible flag before adding replacement user DTs.
        # Preserve an operator's existing setting for compatible system apps.
        if ($cmApp -and -not $allowTaskSequence -and $cmApp.AutoInstall -ne $false) {
            $step = "Set-CMApplication task-sequence eligibility ('$appName')"
            Set-CMApplication -Name $appName -AutoInstall $false -ErrorAction Stop | Out-Null
            Write-Log "Task-sequence install        : disabled (deployment type requires user context or interaction)"
        }

        if (-not $cmApp) {
            Write-Log "Creating ConfigMgr application      : $appName"
            $step = "New-CMApplication ('$appName')"
            $cmAppParams = @{
                Name             = $appName
                Publisher        = $Manifest.Publisher
                SoftwareVersion  = $Manifest.SoftwareVersion
                Description      = $Comment
                AutoInstall      = $allowTaskSequence
                ErrorAction      = 'Stop'
            }
            # Set Software Center display name if provided (omits channel/arch details)
            if ($Manifest.DisplayName) {
                $cmAppParams['LocalizedApplicationName'] = $Manifest.DisplayName
                Write-Log "Software Center name         : $($Manifest.DisplayName)"
            }
            $cmApp = New-CMApplication @cmAppParams

            Write-Log "Application CI_ID            : $($cmApp.CI_ID)"
        }

        # Create every deployment type in spec order: CM assigns priority by
        # creation order and the client installs the first deployment type
        # whose requirements pass, so the most specific variant is listed
        # first and the unconditional fallback last. On a version replace
        # each new deployment type starts under a staging name; the old set
        # is removed and the staged names take the canonical names only
        # after every new deployment type created successfully.
        $stagedRenames = @()
        foreach ($spec in $dtSpecs) {
            $det = $spec.Detection
            $detType = if ($det.Type) { $det.Type } else { 'RegistryKeyValue' }

            $dtName = $spec.DtName
            $dtCreateName = if ($replaceDtNames.Count -gt 0) { "$dtName (staging)" } else { $dtName }
            # Deployment type command lines default to the generated .bat
            # wrappers; manifests may override both (PSADT-wrapped apps point
            # at the toolkit entry, e.g. Invoke-AppDeployToolkit.exe
            # -DeploymentType Install).
            $installCommand = [string]$spec.InstallCommand
            $uninstallCommand = [string]$spec.UninstallCommand
            if ($installCommand -ne 'install.bat') {
                Write-Log "Install command (manifest)   : $installCommand"
            }
            if ($uninstallCommand -ne 'uninstall.bat') {
                Write-Log "Uninstall command (manifest) : $uninstallCommand"
            }

            # The manifest's own timing wins over the caller's parameters: it is
            # the value the operator saved with the profile, and a run that
            # silently used the GUI defaults would misreport install duration
            # and shorten the maintenance-window budget.
            $timing = Resolve-DeploymentTypeTiming -Timing $spec.Timing `
                -DefaultEstimated $EstimatedRuntimeMins -DefaultMaximum $MaximumRuntimeMins `
                -Context ("deployment type '$dtName'")
            if ($timing.Source -eq 'manifest') {
                Write-Log ("Runtime minutes (manifest)   : estimated {0}, maximum {1}" -f $timing.EstimatedRuntimeMins, $timing.MaximumRuntimeMins)
            }

            $dtParams = @{
                ApplicationName           = $appName
                DeploymentTypeName        = $dtCreateName
                ContentLocation           = $spec.ContentLocation
                InstallCommand            = $installCommand
                UninstallCommand          = $uninstallCommand
                InstallationBehaviorType  = 'InstallForSystem'
                LogonRequirementType      = 'WhetherOrNotUserLoggedOn'
                EstimatedRuntimeMins      = $timing.EstimatedRuntimeMins
                MaximumRuntimeMins        = $timing.MaximumRuntimeMins
                ContentFallback           = [bool]$contentOptions.ContentFallback
                SlowNetworkDeploymentMode = [string]$contentOptions.SlowNetworkDeploymentMode
                UserInteractionMode       = 'Hidden'
                ErrorAction               = 'Stop'
            }

            # Manifest field name matches the cmdlet's parameter TYPE
            # (PostExecutionBehavior); the actual parameter name is
            # -RebootBehavior. See Add-CMScriptDeploymentType docs.
            # Default to BasedOnExitCode when the manifest doesn't specify so
            # MSI 3010 "reboot required" exits actually propagate to the CCM
            # client. The cmdlet's own default is NoAction, which silently drops
            # 3010s and breaks install chains that need a reboot between apps.
            if ($spec.PostExecutionBehavior) {
                $dtParams['RebootBehavior'] = $spec.PostExecutionBehavior
            }
            else {
                $dtParams['RebootBehavior'] = 'BasedOnExitCode'
            }

            if ($spec.InstallationBehaviorType) {
                $dtParams['InstallationBehaviorType'] = $spec.InstallationBehaviorType
            }
            if ($spec.LogonRequirementType) {
                $dtParams['LogonRequirementType'] = $spec.LogonRequirementType
            }
            if ($spec.RequireUserInteraction -eq $true) {
                $dtParams['RequireUserInteraction'] = $true
            }

            $detectionTransport = $null
            if ($detType -eq 'Script') {
                # Script-based detection: pass script text, no clause objects needed
                $lang = if ($det.ScriptLanguage) { $det.ScriptLanguage } else { 'PowerShell' }
                $dtParams['ScriptLanguage'] = $lang
                $detectionTransport = Resolve-DetectionScriptTransport -Detection $det -ContentLocation $spec.ContentLocation -SignatureStatus $detectionSignatureStatus
                if ($detectionTransport.Transport -eq 'File') {
                    $dtParams['ScriptFile'] = $detectionTransport.Path
                    Write-Log ("Detection script (file)      : {0}" -f $detectionTransport.Path)
                }
                else {
                    $dtParams['ScriptText'] = $detectionTransport.Text
                }
            }
            else {
                # Clause-based detection: leave CM PSDrive to create clause objects
                # (CM PSDrive context can interfere with parameter binding)
                $step = "New detection clause(s) (type=$detType, dt='$dtName')"
                Set-Location C: -ErrorAction Stop

                if ($detType -eq 'Compound') {
                    $clauses = @()
                    foreach ($c in $det.Clauses) {
                        $clauses += New-SingleDetectionClause -Det $c
                    }
                    $dtParams['AddDetectionClause'] = $clauses

                    $groupSizes = @()
                    if ($det.PSObject.Properties.Name -contains 'GroupSizes' -and $null -ne $det.GroupSizes) {
                        $groupSizes = @($det.GroupSizes | ForEach-Object { [int]$_ })
                    }

                    if ($groupSizes.Count -gt 0) {
                        # (run1 AND ...) OR (run2 AND ...): exactly two contiguous
                        # clause runs. Only the second run is passed to
                        # -GroupDetectionClauses; the cmdlet's left-associative
                        # expression build parenthesizes the first run on its own,
                        # so grouping both runs is impossible (String[] holds one
                        # group) and unnecessary. OR attaches to the first clause
                        # of the second run; clauses inside a run keep AND.
                        if ($groupSizes.Count -ne 2 -or (($groupSizes[0] + $groupSizes[1]) -ne $clauses.Count) -or ($groupSizes -contains 0)) {
                            throw "Detection.GroupSizes must be exactly two non-zero sizes summing to the clause count (got: '$($groupSizes -join ',')' for $($clauses.Count) clauses)."
                        }
                        $secondStart = $groupSizes[0]
                        # A group holds two or more clauses; a one-clause second
                        # run takes the OR connector alone.
                        if ($groupSizes[1] -ge 2) {
                            $dtParams['GroupDetectionClauses'] = @($clauses[$secondStart..($clauses.Count - 1)] | ForEach-Object { $_.Setting.LogicalName })
                        }
                        $dtParams['DetectionClauseConnector'] = @(@{
                            LogicalName = $clauses[$secondStart].Setting.LogicalName
                            Connector   = 'OR'
                        })
                        Write-Log ("Detection expression         : (clauses 1..{0}) OR (clauses {1}..{2})" -f $groupSizes[0], ($secondStart + 1), $clauses.Count)
                    }
                    # OR connector: specify OR for each clause beyond the first
                    # AND is the default and needs no explicit connector
                    elseif ($det.Connector -eq 'Or' -and $clauses.Count -ge 2) {
                        $connectors = @()
                        for ($i = 1; $i -lt $clauses.Count; $i++) {
                            $connectors += @{
                                LogicalName = $clauses[$i].Setting.LogicalName
                                Connector   = 'OR'
                            }
                        }
                        $dtParams['DetectionClauseConnector'] = $connectors
                    }
                }
                else {
                    # Single clause: RegistryKeyValue, RegistryKey, or File
                    $clause = New-SingleDetectionClause -Det $det
                    $dtParams['AddDetectionClause'] = @($clause)
                }

                # Reconnect to CM site for Add-CMScriptDeploymentType
                $step = "Connect-CMSite reconnect (SiteCode=$SiteCode)"
                if (-not (Connect-CMSite -SiteCode $SiteCode)) {
                    throw "ConfigMgr site reconnection failed."
                }
            }

            $requirementRules = @($spec.RequirementRules)
            if ($requirementRules.Count -gt 0) {
                $dtParams['AddRequirement'] = $requirementRules
                Write-Log ("Requirement rules attached   : {0} ('{1}')" -f $requirementRules.Count, $dtName)
            }

            Write-Log ("Deployment type parameters   : {0}" -f (($dtParams.Keys | Sort-Object | ForEach-Object { "{0}='{1}'" -f $_, $dtParams[$_] }) -join ' ')) -Level DEBUG

            Write-Log "Adding Script Deployment Type : $dtCreateName"
            $step = "Add-CMScriptDeploymentType ('$dtCreateName')"
            Add-CMScriptDeploymentType @dtParams | Out-Null

            if ($detectionTransport) {
                $step = "Get-CMDeploymentType detection read-back ('$dtCreateName')"
                Test-StoredDetectionScript -ApplicationName $appName -DeploymentTypeName $dtCreateName `
                    -Transport $detectionTransport -ExpectedThumbprint $detectionSignatureThumbprint
            }

            if ($dtCreateName -ne $dtName) {
                $stagedRenames += [pscustomobject]@{ From = $dtCreateName; To = $dtName }
            }
        }

        foreach ($oldDt in $replaceDtNames) {
            $step = "Remove-CMDeploymentType ('$oldDt')"
            Remove-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $oldDt -Force -ErrorAction Stop
            Write-Log "Removed old deployment type  : $oldDt"
        }
        foreach ($rename in $stagedRenames) {
            $step = "Set-CMDeploymentType rename ('$($rename.From)' -> '$($rename.To)')"
            Set-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $rename.From -NewDeploymentTypeName $rename.To -ErrorAction Stop
            Write-Log "Renamed deployment type      : $($rename.To)"
        }

        if ($existing) {
            $step = "Set-CMApplication version update ('$appName')"
            $setAppParams = @{
                Name            = $appName
                SoftwareVersion = [string]$Manifest.SoftwareVersion
                ErrorAction     = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($Comment)) { $setAppParams['Description'] = $Comment }
            if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.Publisher)) { $setAppParams['Publisher'] = [string]$Manifest.Publisher }
            Set-CMApplication @setAppParams
            Write-Log "Updated application version  : $($Manifest.SoftwareVersion)"
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$supersedencePlan.Marker)) {
            $step = "Identity marker ('$appName')"
            [void](Set-CMApplicationIdentityMarker -ApplicationName $appName -Marker $supersedencePlan.Marker)
        }

        $step = "Set-CMApplication icon ('$appName')"
        Set-CMApplicationIconFromManifest -Manifest $Manifest -AppName $appName -NetworkContentPath $NetworkContentPath

        $step = "Remove-CMApplicationRevisionHistory (CI_ID=$($cmApp.CI_ID))"
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        $step = "Get-CMApplication final read ('$appName')"
        $finalCiId = Get-CMApplicationCurrentCIId -Name $appName -Fallback ([UInt32]$cmApp.CI_ID)

        # Optional auto-distribute to a DP group, and optional test-collection
        # deployment after successful distribution. Settings live in
        # AppPackager.preferences.json alongside the GUI. Packagers invoked
        # from the CLI with no prefs file silently skip both steps.
        try {
            $prefsPath = Join-Path $PSScriptRoot '..\AppPackager.preferences.json'
            if (Test-Path -LiteralPath $prefsPath) {
                $prefsRaw = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($prefsRaw)) {
                    $prefsObj = $prefsRaw | ConvertFrom-Json -ErrorAction Stop
                    $cd = $prefsObj.ContentDistribution
                    if ($cd -and [bool]$cd.AutoDistribute -and -not [string]::IsNullOrWhiteSpace([string]$cd.DPGroupName)) {
                        $dpGroup = [string]$cd.DPGroupName
                        Write-Log "Distributing content         : DP group '$dpGroup'"
                        try {
                            Start-CMContentDistribution -ApplicationName $appName -DistributionPointGroupName $dpGroup -ErrorAction Stop
                            Write-Log "Content distribution         : initiated"
                        } catch {
                            # "already been targeted" is the canonical pattern for a DP group that already holds this app
                            if ($_.Exception.Message -match 'already been targeted|already distributed') {
                                Write-Log "Content distribution         : already targeted (treated as success)"
                            } else {
                                Write-Log "Content distribution failed  : $($_.Exception.Message)" -Level WARN
                            }
                        }

                        # Availability of the test-deployment option is gated in the
                        # GUI (requires auto-distribute + DP group); at runtime the
                        # prefs are taken as-is.
                        $testCollection = if ($null -ne $cd.TestCollectionName) { ([string]$cd.TestCollectionName).Trim() } else { '' }
                        if ($cd.DeployToTestCollection -and -not [string]::IsNullOrWhiteSpace($testCollection)) {
                            try {
                                $isUserCollection = ($cd.PSObject.Properties['TestCollectionIsUser'] -and [bool]$cd.TestCollectionIsUser)
                                $limitingName = if ($isUserCollection) { 'All Users' } else { 'All Systems' }
                                $collection = if ($isUserCollection) { Get-CMUserCollection -Name $testCollection -ErrorAction SilentlyContinue } else { Get-CMDeviceCollection -Name $testCollection -ErrorAction SilentlyContinue }
                                if (-not $collection -and [bool]$cd.CreateTestCollectionIfMissing) {
                                    Write-Log "Creating test collection     : '$testCollection' (direct membership, limited to $limitingName)"
                                    if ($isUserCollection) { New-CMUserCollection -Name $testCollection -LimitingCollectionName $limitingName -ErrorAction Stop | Out-Null }
                                    else { New-CMDeviceCollection -Name $testCollection -LimitingCollectionName $limitingName -ErrorAction Stop | Out-Null }
                                    $collection = if ($isUserCollection) { Get-CMUserCollection -Name $testCollection -ErrorAction SilentlyContinue } else { Get-CMDeviceCollection -Name $testCollection -ErrorAction SilentlyContinue }
                                }
                                if (-not $collection) {
                                    Write-Log "Test deployment skipped      : collection '$testCollection' not found (enable 'Create collection if it does not exist' or create it first)" -Level WARN
                                }
                                else {
                                    Write-Log "Deploying to test collection : '$testCollection' (Available, immediate)"
                                    New-CMApplicationDeployment `
                                        -Name $appName `
                                        -CollectionName $testCollection `
                                        -DeployAction Install `
                                        -DeployPurpose Available `
                                        -AvailableDateTime (Get-Date) `
                                        -ErrorAction Stop | Out-Null
                                    Write-Log "Test deployment              : created"
                                }
                            } catch {
                                if ($_.Exception.Message -match 'already (exists|deployed|been deployed)') {
                                    Write-Log "Test deployment              : already exists (treated as success)"
                                } else {
                                    Write-Log "Test deployment failed       : $($_.Exception.Message)" -Level WARN
                                }
                            }
                        }
                    }
                }
            }
        } catch {
            Write-Log "Could not read prefs for auto-distribute: $($_.Exception.Message)" -Level WARN
        }

        # The supersedence step runs when the application is complete. A failure is reported and never fails the run: the
        # application stays, and the log says "Published, supersedence not applied".
        if ($supersedenceActive) {
            $step = 'Supersedence step'
            [void](Invoke-SupersedenceStepAndLog -ApplicationName $appName -Version ([string]$Manifest.SoftwareVersion) -Plan $supersedencePlan)
        }

        Write-Log ""
        Write-Log "Created ConfigMgr application : $appName"

        return $finalCiId
    }
    catch {
        Write-LogErrorRecord -ErrorRecord $_ -Context ("New-MECMApplicationFromManifest failed during step: {0}" -f $step)
        throw
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Intune Win32 content prep
# ---------------------------------------------------------------------------

function Install-IntuneWinAppUtil {
    <#
    .SYNOPSIS
        Downloads IntuneWinAppUtil.exe into a local tool folder.

    .DESCRIPTION
        Fetches the Microsoft Win32 Content Prep Tool from its official
        GitHub repository and verifies the Authenticode signature is Valid
        and Microsoft-signed before the file lands at its final path. A
        download that fails verification is deleted and the function
        throws. The tool is downloaded per workstation on first use and is
        never redistributed with AppPackager.

    .OUTPUTS
        [string] Full path to the verified IntuneWinAppUtil.exe.
    #>
    param(
        [Parameter(Mandatory)][string]$DestinationFolder,
        [string]$DownloadUrl = 'https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe'
    )

    Initialize-Folder -Path $DestinationFolder
    $target = Join-Path $DestinationFolder 'IntuneWinAppUtil.exe'
    $temp   = Join-Path $DestinationFolder ('IntuneWinAppUtil.' + [Guid]::NewGuid().ToString('N') + '.tmp')

    try {
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $temp

        $sig = Get-AuthenticodeSignature -LiteralPath $temp
        if ($sig.Status -ne 'Valid') {
            throw ("IntuneWinAppUtil.exe download signature status is '{0}', expected 'Valid'; file discarded." -f $sig.Status)
        }
        $subject = if ($sig.SignerCertificate) { [string]$sig.SignerCertificate.Subject } else { '' }
        if ($subject -notmatch 'O=Microsoft Corporation') {
            throw ("IntuneWinAppUtil.exe download signer is '{0}', expected a Microsoft Corporation certificate; file discarded." -f $subject)
        }

        Move-Item -LiteralPath $temp -Destination $target -Force
        Write-Log ("IntuneWinAppUtil.exe verified and installed: {0}" -f $target)
        return $target
    }
    finally {
        if (Test-Path -LiteralPath $temp) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-IntuneWinPackage {
    <#
    .SYNOPSIS
        Produces a .intunewin file from a staged content folder.

    .DESCRIPTION
        Runs the Microsoft Win32 Content Prep Tool against ContentFolder
        with SetupFile as the setup reference. The tool writes
        <setup-basename>.intunewin into OutputFolder; when OutputName is
        given the file is renamed to it. OutputFolder must not be the
        content folder itself: stage-manifest hash verification treats any
        file added to staged or network content as an integrity failure.

    .OUTPUTS
        [pscustomobject] IntuneWinPath, SizeBytes, Sha256, ToolExitCode,
        DurationSec.
    #>
    param(
        [Parameter(Mandatory)][string]$ToolPath,
        [Parameter(Mandatory)][string]$ContentFolder,
        [Parameter(Mandatory)][string]$SetupFile,
        [Parameter(Mandatory)][string]$OutputFolder,
        [string]$OutputName = '',
        [int]$TimeoutSec = 900
    )

    if (-not (Test-Path -LiteralPath $ToolPath)) {
        throw "IntuneWinAppUtil.exe not found: $ToolPath"
    }
    if (-not (Test-Path -LiteralPath $ContentFolder)) {
        throw "Content folder not found: $ContentFolder"
    }
    $setupPath = Join-Path $ContentFolder $SetupFile
    if (-not (Test-Path -LiteralPath $setupPath)) {
        throw "Setup file not found in content folder: $setupPath"
    }
    $contentFull = (Resolve-Path -LiteralPath $ContentFolder).ProviderPath.TrimEnd('\')
    Initialize-Folder -Path $OutputFolder
    $outputFull = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath.TrimEnd('\')
    if ($outputFull -ieq $contentFull) {
        throw "OutputFolder must differ from ContentFolder; a .intunewin inside the content folder fails stage hash verification."
    }

    $expected = Join-Path $outputFull ([IO.Path]::GetFileNameWithoutExtension($SetupFile) + '.intunewin')

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = $ToolPath
    # -q answers the tool's overwrite prompt; without it a pre-existing
    # output file stalls the run waiting for console input.
    $psi.Arguments = ('-c "{0}" -s "{1}" -o "{2}" -q' -f $contentFull, $setupPath, $outputFull)
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $stdOutTask = $proc.StandardOutput.ReadToEndAsync()
        $stdErrTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            try { $proc.Kill() } catch { <# Killing a timed-out process is best-effort; the timeout is reported after. #> $null = $_ }
            throw ("IntuneWinAppUtil.exe did not finish within {0}s; process terminated." -f $TimeoutSec)
        }
        $sw.Stop()
        $exitCode = $proc.ExitCode
        $stdOut = [string]$stdOutTask.Result
        $stdErr = [string]$stdErrTask.Result
    }
    finally {
        $proc.Dispose()
    }

    if ($exitCode -ne 0) {
        $tail = @((($stdErr, $stdOut) -join "`n") -split "`r?`n" | Where-Object { $_ } | Select-Object -Last 5) -join ' | '
        throw ("IntuneWinAppUtil.exe exit code {0}: {1}" -f $exitCode, $tail)
    }
    if (-not (Test-Path -LiteralPath $expected)) {
        throw ("IntuneWinAppUtil.exe reported success but the output file is missing: {0}" -f $expected)
    }

    $final = $expected
    if (-not [string]::IsNullOrWhiteSpace($OutputName)) {
        $final = Join-Path $outputFull $OutputName
        Move-Item -LiteralPath $expected -Destination $final -Force
    }

    $item = Get-Item -LiteralPath $final
    return [pscustomobject]@{
        IntuneWinPath = [string]$item.FullName
        SizeBytes     = [long]$item.Length
        Sha256        = (Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash
        ToolExitCode  = [int]$exitCode
        DurationSec   = [int]$sw.Elapsed.TotalSeconds
    }
}

# ---------------------------------------------------------------------------
# Packager preferences
# ---------------------------------------------------------------------------

function Get-PackagerRunOptions {
    <#
    .SYNOPSIS
        Reads the values of this run from APP_PACKAGER_RUN_OPTIONS, or returns $null.

    .DESCRIPTION
        The review before staging sets them for one Stage run: Preferences holds the
        top-level keys of packager-preferences.json that the run changes, and Files maps
        the name of a packager settings file to the path of the copy that this run reads.
        Nothing in the saved files changes.
    #>
    $text = [string]$env:APP_PACKAGER_RUN_OPTIONS
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return ($text | ConvertFrom-Json -ErrorAction Stop) }
    catch {
        Write-Log "APP_PACKAGER_RUN_OPTIONS is not valid JSON and is ignored: $($_.Exception.Message)" -Level WARN
        return $null
    }
}

function Resolve-PackagerRunFile {
    <#
    .SYNOPSIS
        Returns the path of a packager settings file for this run.

    .DESCRIPTION
        Returns the copy that the review before staging made for this run, when there is
        one, otherwise the saved file given as Default.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Default
    )
    $run = Get-PackagerRunOptions
    if ($run -and $run.PSObject.Properties['Files'] -and $run.Files -and $run.Files.PSObject.Properties[$Name]) {
        $path = [string]$run.Files.$Name
        if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            Write-Log "Using the values of this run for $Name"
            return $path
        }
    }
    return $Default
}

function Get-PackagerPreferences {
    <#
    .SYNOPSIS
        Reads packager-preferences.json from the Packagers folder.

    .DESCRIPTION
        The keys that the review before staging changed for this run
        (APP_PACKAGER_RUN_OPTIONS) replace the saved ones in the returned object.
    #>
    $prefsPath = Join-Path $PSScriptRoot "packager-preferences.json"
    $prefs = $null
    if (-not (Test-Path -LiteralPath $prefsPath)) {
        Write-Log "Preferences file not found: $prefsPath" -Level WARN
    }
    else {
        $json = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
        $prefs = ($json | ConvertFrom-Json)
    }
    $run = Get-PackagerRunOptions
    if ($run -and $run.PSObject.Properties['Preferences'] -and $run.Preferences) {
        if ($null -eq $prefs) { $prefs = [pscustomobject]@{} }
        foreach ($property in $run.Preferences.PSObject.Properties) {
            $prefs | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value -Force
        }
    }
    return $prefs
}

function Get-LocalSourceFolder {
    <#
    .SYNOPSIS
        Returns the local source folder for a packager: the override when
        given, otherwise the folder saved in packager preferences, or ''.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagerName,
        [string]$Override
    )

    if (-not [string]::IsNullOrWhiteSpace($Override)) { return $Override }
    $prefs = Get-PackagerPreferences
    if ($prefs -and $prefs.PSObject.Properties['LocalSourceFolders'] -and $prefs.LocalSourceFolders -and
        $prefs.LocalSourceFolders.PSObject.Properties[$PackagerName]) {
        return [string]$prefs.LocalSourceFolders.$PackagerName
    }
    return ''
}

function Resolve-LocalSourceInstaller {
    <#
    .SYNOPSIS
        Returns the highest-version installer matching a file pattern in a
        hybrid packager's local source folder.

    .DESCRIPTION
        Hybrid packagers read the latest version from a public page, but the
        vendor download requires a sign-in, so the installer comes from a
        folder chosen in Options or at the Stage prompt. Vendors reuse one
        file name across releases, so the file version decides, not the
        download time. Throws when no folder is set, the folder is missing, or
        it holds no matching installer.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagerName,
        [Parameter(Mandatory)][string]$Filter,
        [string[]]$Exclude = @(),
        [string]$Override
    )

    $folder = Get-LocalSourceFolder -PackagerName $PackagerName -Override $Override

    if ([string]::IsNullOrWhiteSpace($folder)) {
        throw "No installer source folder is set for $PackagerName. Choose one in Options > Applications > Local Installer Sources, or run Stage from the app to be prompted."
    }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        throw "Installer source folder for $PackagerName not found: $folder"
    }

    $candidates = @(Get-ChildItem -LiteralPath $folder -Filter $Filter -File -ErrorAction SilentlyContinue |
        Where-Object { $name = $_.Name; -not @($Exclude | Where-Object { $name -like $_ }) })
    if ($candidates.Count -eq 0) {
        throw "No installer matching '$Filter' found in $folder"
    }

    $ranked = foreach ($c in $candidates) {
        $v = $null
        # FileVersion may carry a build label after the number ("10.0.1 (WinBuild...)").
        $raw = ([string][System.Diagnostics.FileVersionInfo]::GetVersionInfo($c.FullName).FileVersion).Trim()
        if ($raw -match '^\d+(\.\d+){1,3}') { [void][version]::TryParse($Matches[0], [ref]$v) }
        if (-not $v) { $v = [version]'0.0' }
        [pscustomobject]@{ File = $c; Version = $v }
    }
    $best = $ranked | Sort-Object -Property @{ Expression = 'Version'; Descending = $true }, @{ Expression = { $_.File.LastWriteTimeUtc }; Descending = $true } |
        Select-Object -First 1
    return $best.File.FullName
}

# ---------------------------------------------------------------------------
# ODT config XML generation
# ---------------------------------------------------------------------------

function ConvertTo-XmlAttributeValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Get-OdtGeneratedOption {
    <#
    .SYNOPSIS
        The generated-configuration values of the Microsoft 365 packagers, from the
        M365Options object of the packager preferences, as a hashtable that splats
        into New-OdtConfigXml. A missing object or value keeps the packager default.
    #>
    param([AllowNull()]$Preferences)

    $options = @{
        Languages               = @('en-us')
        OfficeMgmtCom           = $true
        UpdatesEnabled          = $true
        ForceAppShutdown        = $true
        RemoveMsi               = $true
        MigrateArch             = $false
        SharedComputerLicensing = $false
        DeviceBasedLicensing    = $false
    }
    if ($Preferences -and $Preferences.PSObject.Properties['M365Options'] -and $Preferences.M365Options) {
        $stored = $Preferences.M365Options
        foreach ($name in @('OfficeMgmtCom', 'UpdatesEnabled', 'ForceAppShutdown', 'RemoveMsi', 'MigrateArch', 'SharedComputerLicensing', 'DeviceBasedLicensing')) {
            if ($stored.PSObject.Properties[$name] -and $null -ne $stored.$name) { $options[$name] = [bool]$stored.$name }
        }
        if ($stored.PSObject.Properties['Languages'] -and $null -ne $stored.Languages) {
            $codes = @($stored.Languages | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -match '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){1,2}$' })
            if ($codes.Count -gt 0) { $options.Languages = $codes }
        }
    }
    return $options
}

function ConvertFrom-OdtConfigText {
    # Reads a configuration.xml text into an XmlDocument. A DOCTYPE and external entities are refused.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $clean = $Text.TrimStart([char]0xFEFF)
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($clean)), $settings)
    try {
        $document = New-Object System.Xml.XmlDocument
        $document.XmlResolver = $null
        $document.PreserveWhitespace = $false
        $document.Load($reader)
        return $document
    }
    finally { $reader.Dispose() }
}

function Test-OdtConfigurationXml {
    <#
    .SYNOPSIS
        Checks a configuration.xml for the Microsoft 365 packagers: well-formed XML, a
        Configuration root, and an Add element with at least one Product that has an ID.

    .OUTPUTS
        Valid, Errors (what stops the import), Notes (what the packager changes) and
        ProductIds.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $errors = New-Object System.Collections.Generic.List[string]
    $notes = New-Object System.Collections.Generic.List[string]
    $productIds = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Text)) {
        $errors.Add('The file is empty.')
        return [pscustomobject]@{ Valid = $false; Errors = $errors.ToArray(); Notes = @(); ProductIds = @() }
    }

    $document = $null
    try { $document = ConvertFrom-OdtConfigText -Text $Text }
    catch { $errors.Add('The file is not well-formed XML: ' + $_.Exception.Message) }

    if ($document) {
        $root = $document.DocumentElement
        if ($root.LocalName -cne 'Configuration') { $errors.Add(("The root element is '{0}'. An Office Deployment Tool configuration has the root element Configuration." -f $root.LocalName)) }
        else {
            $add = @($root.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Add' }) | Select-Object -First 1
            if (-not $add) { $errors.Add('The file has no Add element. The packager needs an Add element with at least one Product.') }
            else {
                $products = @($add.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Product' })
                if ($products.Count -eq 0) { $errors.Add('The Add element has no Product element.') }
                foreach ($product in $products) {
                    $id = $product.GetAttribute('ID')
                    if ([string]::IsNullOrWhiteSpace($id)) { $errors.Add('A Product element has no ID attribute.') } else { $productIds.Add($id) }
                }
                $channel = $add.GetAttribute('Channel')
                if ($channel -and $channel -notin @('MonthlyEnterprise', 'Current')) {
                    $notes.Add(("Channel '{0}' is not supported by the packagers. They use the channel from Options." -f $channel))
                }
                $notes.Add('OfficeClientEdition, Version and SourcePath are set by the packager.')
                $notes.Add('Every Microsoft 365 Apps, Project and Visio package uses this file with its Product elements, so list the products that you want in the file.')
                $display = @($root.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Display' }) | Select-Object -First 1
                if (-not $display -or $display.GetAttribute('Level') -cne 'None' -or $display.GetAttribute('AcceptEULA') -cne 'TRUE') {
                    $notes.Add('Display is set to Level None and AcceptEULA TRUE, so that the install is silent.')
                }
            }
        }
    }
    return [pscustomobject]@{ Valid = ($errors.Count -eq 0); Errors = $errors.ToArray(); Notes = $notes.ToArray(); ProductIds = $productIds.ToArray() }
}

function Get-OdtUserConfigInfo {
    <#
    .SYNOPSIS
        The channel and the deploy mode that an imported configuration.xml selects.
        Channel is MonthlyEnterprise or Current, or empty when the file names no
        supported channel. DeployMode is Managed when the Add element has a
        SourcePath and Online when it has none, which is how the Office
        customization tool describes a local source and the CDN.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $document = ConvertFrom-OdtConfigText -Text $Text
    $add = @($document.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Add' }) | Select-Object -First 1
    $channel = ''
    $mode = 'Online'
    if ($add) {
        $value = $add.GetAttribute('Channel')
        if ($value -in @('MonthlyEnterprise', 'Current')) { $channel = $value }
        if ($add.GetAttribute('SourcePath')) { $mode = 'Managed' }
    }
    return [pscustomobject]@{ Channel = $channel; DeployMode = $mode }
}

function Get-OdtUserConfigPath {
    <#
    .SYNOPSIS
        The path of the configuration.xml that Options imported, or '' when the packagers
        generate the configuration. Throws when the preference selects the file and
        the file is gone.
    #>
    param([AllowNull()]$Preferences)

    if (-not $Preferences -or -not $Preferences.PSObject.Properties['M365UseCustomConfig'] -or -not [bool]$Preferences.M365UseCustomConfig) { return '' }
    $path = [string]$Preferences.M365CustomConfigPath
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw 'Options selects your own configuration.xml for Microsoft 365, and the stored file is missing. Import it again in Options, Applications, Microsoft 365 Apps.'
    }
    return $path
}

function ConvertTo-OdtUserConfigXml {
    <#
    .SYNOPSIS
        Turns an imported configuration.xml into the install or download XML of one
        packager. The packager sets OfficeClientEdition, Version and SourcePath and a
        silent Display; every other element and attribute of the file stays.

    .PARAMETER Channel
        The channel to write when the file names none that is supported.
    #>
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$OfficeClientEdition,
        [string]$Version,
        [string]$SourcePath,
        [ValidateSet('MonthlyEnterprise', 'Current')][string]$Channel = 'MonthlyEnterprise'
    )

    $check = Test-OdtConfigurationXml -Text $Text
    if (-not $check.Valid) { throw ('The imported configuration.xml is not usable: ' + ($check.Errors -join ' ')) }
    $document = ConvertFrom-OdtConfigText -Text $Text
    $root = $document.DocumentElement
    $add = @($root.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Add' }) | Select-Object -First 1

    $add.SetAttribute('OfficeClientEdition', $OfficeClientEdition)
    if ($add.GetAttribute('Channel') -notin @('MonthlyEnterprise', 'Current')) { $add.SetAttribute('Channel', $Channel) }
    if ($Version -and $Version -ne 'Latest') { $add.SetAttribute('Version', $Version) } else { [void]$add.RemoveAttribute('Version') }
    if ($SourcePath) { $add.SetAttribute('SourcePath', $SourcePath) } else { [void]$add.RemoveAttribute('SourcePath') }

    $display = @($root.ChildNodes | Where-Object { $_.NodeType -eq 'Element' -and $_.LocalName -ceq 'Display' }) | Select-Object -First 1
    if (-not $display) { $display = $document.CreateElement('Display'); [void]$root.AppendChild($display) }
    $display.SetAttribute('Level', 'None')
    $display.SetAttribute('AcceptEULA', 'TRUE')

    $writerSettings = New-Object System.Xml.XmlWriterSettings
    $writerSettings.OmitXmlDeclaration = $true
    $writerSettings.Indent = $true
    $writerSettings.IndentChars = '  '
    $writerSettings.NewLineChars = "`r`n"
    $writerSettings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
    $builder = New-Object System.Text.StringBuilder
    $writer = [System.Xml.XmlWriter]::Create($builder, $writerSettings)
    try { $document.Save($writer) } finally { $writer.Dispose() }

    # The packagers write the file as ASCII, so characters above U+007F leave as numeric character references.
    $xml = $builder.ToString()
    $ascii = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $xml.Length; $i++) {
        $c = $xml[$i]
        if ([int]$c -lt 128) { [void]$ascii.Append($c) }
        elseif ([char]::IsHighSurrogate($c) -and ($i + 1) -lt $xml.Length) { [void]$ascii.AppendFormat('&#x{0:X};', [char]::ConvertToUtf32($c, $xml[$i + 1])); $i++ }
        else { [void]$ascii.AppendFormat('&#x{0:X};', [int]$c) }
    }
    return $ascii.ToString()
}

function New-OdtConfigXml {
    <#
    .SYNOPSIS
        Generates a full ODT configuration XML string for download or install.

    .DESCRIPTION
        Builds the XML of the Microsoft 365 packagers from the product list, the
        excluded apps and the install behavior values. With UserConfigPath, the
        XML comes from the configuration.xml that Options imported, and the
        packager sets only the edition, the version and the source path.

    .PARAMETER OfficeClientEdition
        Architecture: "32" or "64".

    .PARAMETER Version
        Full M365 version string (for example "16.0.19127.20532").

    .PARAMETER ProductIds
        Array of product IDs (for example @('O365ProPlusRetail') or
        @('O365ProPlusRetail', 'VisioProRetail')).

    .PARAMETER SourcePath
        SourcePath attribute for the Add element. For download: local content
        folder path. For install: ".".

    .PARAMETER Channel
        ODT channel name. Valid values: MonthlyEnterprise, Current.
        Default: MonthlyEnterprise.

    .PARAMETER CompanyName
        Value for the AppSettings Company name. Omit or pass empty to skip
        the AppSettings block entirely.

    .PARAMETER Languages
        Language IDs of each product. The first language sets the shell UI language.

    .PARAMETER MigrateArch
        Writes MigrateArch="TRUE", which changes an installed Office to OfficeClientEdition.
        The ODT default is off.

    .PARAMETER SharedComputerLicensing
        Writes the SharedComputerLicensing property with the value 1. The ODT default is 0.
    #>
    param(
        [Parameter(Mandatory)][string]$OfficeClientEdition,
        [string]$Version,
        [Parameter(Mandatory)][string[]]$ProductIds,
        [string]$SourcePath,
        [ValidateSet('MonthlyEnterprise','Current')]
        [string]$Channel = 'MonthlyEnterprise',
        [string]$CompanyName,
        # ExcludeApp IDs per product (e.g. 'Groove','Lync','Teams'). Pass @() to include everything.
        [string[]]$ExcludeApps = @('Groove','Lync','OneDrive','Teams'),
        [string[]]$Languages = @('en-us'),
        [bool]$OfficeMgmtCom = $true,
        [bool]$UpdatesEnabled = $true,
        [bool]$ForceAppShutdown = $true,
        [bool]$RemoveMsi = $true,
        [bool]$MigrateArch = $false,
        [bool]$SharedComputerLicensing = $false,
        [bool]$DeviceBasedLicensing = $false,
        [string]$UserConfigPath = ''
    )

    if ($UserConfigPath) {
        $userText = [System.IO.File]::ReadAllText($UserConfigPath, [System.Text.Encoding]::UTF8)
        return (ConvertTo-OdtUserConfigXml -Text $userText -OfficeClientEdition $OfficeClientEdition -Version $Version -SourcePath $SourcePath -Channel $Channel)
    }

    $addAttrs = @()
    if ($SourcePath) {
        $addAttrs += 'SourcePath="{0}"' -f (ConvertTo-XmlAttributeValue $SourcePath)
    }
    $addAttrs += 'OfficeClientEdition="{0}"' -f (ConvertTo-XmlAttributeValue $OfficeClientEdition)
    $addAttrs += 'Channel="{0}"' -f (ConvertTo-XmlAttributeValue $Channel)
    if ($OfficeMgmtCom) { $addAttrs += 'OfficeMgmtCOM="TRUE"' }
    if ($Version -and $Version -ne 'Latest') {
        $addAttrs += 'Version="{0}"' -f (ConvertTo-XmlAttributeValue $Version)
    }
    if ($MigrateArch) { $addAttrs += 'MigrateArch="TRUE"' }

    $languageIds = @($Languages | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($languageIds.Count -eq 0) { $languageIds = @('en-us') }

    $lines = @('<Configuration>')
    $lines += '  <Add {0}>' -f ($addAttrs -join ' ')

    foreach ($prodId in $ProductIds) {
        $lines += '    <Product ID="{0}">' -f (ConvertTo-XmlAttributeValue $prodId)
        foreach ($languageId in $languageIds) { $lines += '      <Language ID="{0}" />' -f (ConvertTo-XmlAttributeValue $languageId) }
        foreach ($appId in $ExcludeApps) {
            if (-not [string]::IsNullOrWhiteSpace($appId)) {
                $lines += '      <ExcludeApp ID="{0}" />' -f (ConvertTo-XmlAttributeValue $appId)
            }
        }
        $lines += '    </Product>'
    }

    $lines += '  </Add>'
    if ($SharedComputerLicensing) { $lines += '  <Property Name="SharedComputerLicensing" Value="1" />' }
    if ($ForceAppShutdown) { $lines += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />' }
    if ($DeviceBasedLicensing) { $lines += '  <Property Name="DeviceBasedLicensing" Value="1" />' }
    $lines += '  <Updates Enabled="{0}" />' -f $(if ($UpdatesEnabled) { 'TRUE' } else { 'FALSE' })
    if ($RemoveMsi) { $lines += '  <RemoveMSI />' }
    if ($CompanyName) {
        $lines += '  <AppSettings>'
        $lines += '    <Setup Name="Company" Value="{0}" />' -f (ConvertTo-XmlAttributeValue $CompanyName)
        $lines += '  </AppSettings>'
    }
    $lines += '  <Display Level="None" AcceptEULA="TRUE" />'
    $lines += '  <Logging Level="Standard" Path="%programdata%\Appdeploy\Office2016" />'
    $lines += '</Configuration>'

    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Java vendor release helpers
# ---------------------------------------------------------------------------

function Get-LatestTemurinRelease {
    <#
    .SYNOPSIS
        Queries the Eclipse Adoptium API for the latest Temurin release.
    .DESCRIPTION
        Returns a hashtable with Version, DownloadUrl, and FileName for the
        latest Temurin JRE or JDK MSI installer. The -LTS suffix is stripped
        from the version string.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER ImageType
        'jre' or 'jdk'.
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [Parameter(Mandatory)][ValidateSet('jre','jdk')][string]$ImageType,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.adoptium.net/v3/assets/latest/$FeatureVersion/hotspot?architecture=$Architecture&image_type=$ImageType&os=windows"
    Write-Log "Adoptium API URL             : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Adoptium API." }

        $data = ConvertFrom-Json $json

        $asset = $data | Where-Object { $_.binary.installer.name -match '\.msi$' } | Select-Object -First 1
        if (-not $asset) { throw "No MSI installer found for Temurin $ImageType $FeatureVersion ($Architecture)." }

        $downloadUrl = $asset.binary.installer.link
        $fileName    = $asset.binary.installer.name
        $rawVersion  = $asset.version.semver

        if ([string]::IsNullOrWhiteSpace($rawVersion)) { throw "version.semver is empty in Adoptium API response." }

        $version = $rawVersion -replace '[\.\-]\d*\.?LTS$', ''

        Write-Log ("Temurin {0} {1} version      : {2}" -f $ImageType, $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Temurin release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}


function Get-LatestCorrettoRelease {
    <#
    .SYNOPSIS
        Queries the GitHub API for the latest Amazon Corretto JDK release.
    .DESCRIPTION
        Returns a hashtable with Version (4-part normalized), DownloadUrl, and
        FileName. Corretto uses 5-part versioning; the 5th part (Corretto patch)
        is stripped to produce a 4-part version compatible with the GUI regex
        and the .NET [version] type.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.github.com/repos/corretto/corretto-$FeatureVersion/releases/latest"
    Write-Log "Corretto GitHub API URL      : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error -A "PowerShell" @(Get-GitHubApiCurlArgs) $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Corretto GitHub API." }

        $release = ConvertFrom-Json $json

        $tagVersion = $release.tag_name
        if ([string]::IsNullOrWhiteSpace($tagVersion)) { throw "tag_name is empty in Corretto release response." }

        # Construct MSI filename from known pattern
        # v8: amazon-corretto-{TAG}-windows-{ARCH}-jdk.msi
        # v11+: amazon-corretto-{TAG}-windows-{ARCH}.msi
        if ($FeatureVersion -le 8) {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture-jdk.msi"
        }
        else {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture.msi"
        }
        $downloadUrl = "https://corretto.aws/downloads/resources/$tagVersion/$fileName"

        # Normalize 5-part version to 4 parts (strip Corretto patch)
        $parts = $tagVersion -split '\.'
        if ($parts.Count -ge 5) {
            $version = ($parts[0..3] -join '.')
        }
        else {
            $version = $tagVersion
        }

        Write-Log ("Corretto {0} raw version     : {1}" -f $FeatureVersion, $tagVersion) -Quiet:$Quiet
        Write-Log ("Corretto {0} normalized      : {1}" -f $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Corretto release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}

# ---------------------------------------------------------------------------
# Packager history (per-app timestamps for batch/scheduled workflow)
# ---------------------------------------------------------------------------
# Storage lives under %LOCALAPPDATA%\AppPackager\app-history.json. It is NEVER
# inside the install folder and is auto-created on first use.
#
# Schema (top level is a dictionary keyed by packager base name, e.g. 'package-chrome'):
# {
#   "package-chrome": {
#     "LastChecked":      "2026-04-19T17:00:00Z",
#     "LastStaged":       "2026-04-19T17:05:00Z",
#     "LastPackaged":     "2026-04-19T17:10:00Z",
#     "LastKnownVersion": "140.0.7339.80",
#     "LastResult":       "Updated",    # Checked | NoChange | Updated | Failed
#     "LastCmChecked":    "2026-04-19T17:00:00Z",
#     "LastCmVersion":    "139.0.7258.66",
#     "VulnSummary":      { ... }       # Get-VulnerabilitySummary of AppPackagerVulnerability.psm1
#   },
#   ...
# }

function Get-PackagerHistoryPath {
    <#
    .SYNOPSIS
        Returns the user-profile path where packager history is persisted.
    #>
    $dir = Join-Path $env:LOCALAPPDATA 'AppPackager'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Join-Path $dir 'app-history.json'
}

function Convert-PackagerHistoryTimestampToString {
    param($Value)

    if ($Value -is [datetime]) {
        return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    return $Value
}

function Read-PackagerHistory {
    <#
    .SYNOPSIS
        Loads the packager history dictionary. Returns an empty hashtable if
        the file doesn't exist or is unreadable.
    #>
    $path = Get-PackagerHistoryPath
    if (-not (Test-Path -LiteralPath $path)) { return @{} }
    try {
        $raw  = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
        $obj  = $raw | ConvertFrom-Json -ErrorAction Stop
        $hash = @{}
        foreach ($p in $obj.PSObject.Properties) {
            $entry = $p.Value
            if ($entry -and $entry.PSObject -and $entry.PSObject.Properties) {
                foreach ($datePropName in @('LastChecked','LastStaged','LastPackaged','LastCmChecked')) {
                    $dateProp = $entry.PSObject.Properties[$datePropName]
                    if ($dateProp) {
                        $dateProp.Value = Convert-PackagerHistoryTimestampToString -Value $dateProp.Value
                    }
                }
            }
            $hash[$p.Name] = $entry
        }
        return $hash
    }
    catch {
        Write-Log ("Read-PackagerHistory: could not read {0}: {1}" -f $path, $_.Exception.Message) -Level WARN
        return @{}
    }
}

function Save-PackagerHistory {
    <#
    .SYNOPSIS
        Writes the history dictionary back to disk as UTF-8 JSON.
    .PARAMETER History
        Hashtable keyed by packager base name.
    #>
    param([Parameter(Mandatory)][hashtable]$History)
    $path = Get-PackagerHistoryPath
    try {
        ($History | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $path -Encoding UTF8
    }
    catch {
        Write-Log ("Save-PackagerHistory: could not write {0}: {1}" -f $path, $_.Exception.Message) -Level ERROR
        throw
    }
}

function Get-PackagerHistoryActedVersion {
    <#
    .SYNOPSIS
        The version a history entry shows as done for one kind of run, or ''
        when none is on record.

    .DESCRIPTION
        Checked: the newest version seen. Staged and Packaged: the version a
        completed stage or publish recorded, for the same scope when both
        sides name one. An entry written before these versions were recorded
        counts only when its stage or publish is not older than its latest
        check; otherwise a later check may have seen a newer version and the
        entry cannot say which one was built.
    #>
    param(
        $Entry,
        [Parameter(Mandatory)][ValidateSet('Checked', 'Staged', 'Packaged')][string]$Event,
        [string]$Scope = ''
    )

    if ($null -eq $Entry) { return '' }
    $read = {
        param($name)
        if ($Entry -is [System.Collections.IDictionary]) { if ($Entry.Contains($name)) { return $Entry[$name] }; return $null }
        if ($Entry.PSObject.Properties[$name]) { return $Entry.$name }
        return $null
    }

    $known = [string](& $read 'LastKnownVersion')
    if ($Event -eq 'Checked') { return $known }

    $recorded = [string](& $read "Last${Event}Version")
    if ($recorded) {
        $recordedScope = [string](& $read "Last${Event}Scope")
        if ($Scope -and $recordedScope -and $recordedScope -ne $Scope) { return '' }
        return $recorded
    }

    $when = [string](& $read "Last$Event")
    if (-not $when) { return '' }
    $checked = [string](& $read 'LastChecked')
    if ($checked -and [string]::CompareOrdinal($when, $checked) -lt 0) { return '' }
    return $known
}

function Update-PackagerHistory {
    <#
    .SYNOPSIS
        Records an event against a packager in the history file.
    .PARAMETER PackagerName
        Script base name without extension, for example 'package-chrome'.
    .PARAMETER Event
        What happened: Checked | Staged | Packaged | CmChecked | VulnChecked.
        CmChecked records the version ConfigMgr holds (LastCmVersion,
        LastCmChecked) and leaves LastKnownVersion alone. VulnChecked stores
        the -Data summary as VulnSummary.
    .PARAMETER Version
        Version string associated with the event (the latest version checked,
        staged, or packaged; the ConfigMgr version for CmChecked). Optional
        for non-version events.
    .PARAMETER Result
        Outcome classification used by the summary: Checked | NoChange |
        Updated | Failed. Optional.
    .PARAMETER Data
        VulnChecked: the summary from Get-VulnerabilitySummary, stored as
        VulnSummary. CmChecked: a SiteCode entry, stored as LastCmSite.
        Ignored for other events.
    .PARAMETER Scope
        Staged and Packaged: where the version went (the staging root, or the
        site and share). A later run compares it before it treats that version
        as already done for its own destination.
    #>
    param(
        [Parameter(Mandatory)][string]$PackagerName,
        [Parameter(Mandatory)][ValidateSet('Checked','Staged','Packaged','CmChecked','VulnChecked')][string]$Event,
        [string]$Version,
        [ValidateSet('Checked','NoChange','Updated','Failed')][string]$Result,
        [System.Collections.IDictionary]$Data,
        [string]$Scope = ''
    )
    # A product that Cancel stopped in flight leaves no trace in the history (the run reverts its row).
    if (Test-RunAborted) { return }


    $history = Read-PackagerHistory
    if (-not $history.ContainsKey($PackagerName)) {
        $history[$PackagerName] = [pscustomobject]@{
            LastChecked      = $null
            LastStaged       = $null
            LastPackaged     = $null
            LastKnownVersion = $null
            LastResult       = $null
        }
    }
    $entry = $history[$PackagerName]

    # PSCustomObject from ConvertFrom-Json is read-only for property-set; normalize to hashtable-like.
    if ($entry -isnot [hashtable]) {
        $h = @{}
        foreach ($p in $entry.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $entry = $h
        $history[$PackagerName] = $entry
    }

    $nowUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    switch ($Event) {
        'Checked'     { $entry['LastChecked']  = $nowUtc }
        'Staged'      { $entry['LastStaged']   = $nowUtc }
        'Packaged'    { $entry['LastPackaged'] = $nowUtc }
        'CmChecked'   {
            $entry['LastCmChecked'] = $nowUtc
            $entry['LastCmVersion'] = $(if ($PSBoundParameters.ContainsKey('Version')) { [string]$Version } else { '' })
            if ($Data -and $Data.Contains('SiteCode')) { $entry['LastCmSite'] = [string]$Data['SiteCode'] }
        }
        'VulnChecked' {
            $entry['VulnSummary'] = $(if ($Data) { [pscustomobject]$Data } else { $null })
        }
    }
    if ($Event -notin @('CmChecked', 'VulnChecked') -and $PSBoundParameters.ContainsKey('Version') -and -not [string]::IsNullOrWhiteSpace($Version)) {
        $entry['LastKnownVersion'] = $Version
    }
    # LastKnownVersion is the newest version seen; these two are the versions
    # a stage or a publish actually completed for, which is what decides
    # whether a run still has work to do.
    if ($Event -in @('Staged', 'Packaged') -and -not [string]::IsNullOrWhiteSpace($Version)) {
        $entry["Last${Event}Version"] = $Version
        $entry["Last${Event}Scope"]   = $Scope
    }
    if ($PSBoundParameters.ContainsKey('Result')) {
        $entry['LastResult'] = $Result
    }

    Save-PackagerHistory -History $history
}


# ---------------------------------------------------------------------------
# Ad-hoc drop intake: analyze a dropped installer, stage it through the
# normal manifest path, optionally graduate it into the packager catalog.
# ---------------------------------------------------------------------------

function Get-InstallerAnalysis {
    <#
    .SYNOPSIS
        Runs the vendored installer analysis over one file and returns the
        deployment-relevant digest.

    .DESCRIPTION
        Orchestrates InstallerAnalysisCommon (vendored at
        ..\Lib\InstallerAnalysisCommon): engine identification by binary
        signature, MSI properties when applicable, package metadata,
        silent-switch prediction, and the aggregated deployment fields.
        Confidence is the roadmap's gate: MSI-derived identity is
        authoritative; everything else is predicted and requires operator
        confirmation before Package.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the aggregate analysis for one installer.')]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Installer not found: $Path"
    }
    if (-not (Get-Module -Name InstallerAnalysisCommon)) {
        Import-Module -Name (Join-Path $PSScriptRoot '..\Lib\InstallerAnalysisCommon\InstallerAnalysisCommon.psd1') -Global -DisableNameChecking
    }

    $fileInfo = Get-InstallerFileInfo -Path $Path
    $type     = Get-InstallerType -Path $Path

    $msiProps = $null
    $msiSummary = $null
    if ($type -eq 'MSI') {
        $msiProps   = Get-MsiProperties -MsiPath $Path
        $msiSummary = Get-MsiSummaryInfo -MsiPath $Path
    }
    $pkgMeta  = Get-PackageMetadataFor -Path $Path -InstallerType $type
    $switches = Get-SilentSwitches -InstallerType $type -FilePath $Path -MsiProperties $msiProps -PackageMetadata $pkgMeta
    $fields   = Get-DeploymentFields -FileInfo $fileInfo -MsiProperties $msiProps -Switches $switches `
        -PackageMetadata $pkgMeta -InstallerType $type -MsiSummary $msiSummary

    $productCode = ''
    if ($msiProps -and $msiProps.Contains('ProductCode')) { $productCode = [string]$msiProps['ProductCode'] }

    # Get-SilentSwitches returns full command lines; the wrapper generators
    # want bare arguments, so strip the leading installer filename.
    $installCommand   = if ($switches) { [string]$switches.Install }   else { '' }
    $uninstallCommand = if ($fields -and $fields.SilentUninstallString) { [string]$fields.SilentUninstallString } elseif ($switches) { [string]$switches.Uninstall } else { '' }
    $installArgs = ''
    $uninstallArgs = ''
    if ($type -eq 'MSI') {
        $installArgs   = '/qn /norestart'
        $uninstallArgs = '/qn /norestart'
    }
    else {
        $fileNameEscaped = [regex]::Escape((Split-Path -Leaf $Path))
        if ($installCommand -match ('^"?' + $fileNameEscaped + '"?\s+(?<args>\S.*)$')) {
            $installArgs = $Matches['args'].Trim()
        }
    }

    # Install context and registry placement: an NSIS header names the ARP
    # hive and the 32/64-bit view the script wrote under; other formats
    # fall back to the predicted key's hive prefix.
    $installContext = ''
    $registryHive   = ''
    $registryView   = ''
    $installDir     = ''
    $metaProp = { param($name) if ($pkgMeta -and $pkgMeta.PSObject.Properties[$name]) { [string]$pkgMeta.$name } else { '' } }
    if ($pkgMeta) {
        $installContext = & $metaProp 'InstallContext'
        # Format metadata spells the context in prose ('Per-user');
        # every consumer downstream compares against PerUser/PerMachine.
        if ($installContext -match '^Per-?user$')        { $installContext = 'PerUser' }
        elseif ($installContext -match '^Per-?machine$') { $installContext = 'PerMachine' }
        $registryHive   = & $metaProp 'RegistryHive'
        $registryView   = & $metaProp 'RegistryView'
        $installDir     = & $metaProp 'InstallDirWindows'
    }
    $predictedKey = [string]$fields.UninstallRegistryKey
    if (-not $registryHive -and $predictedKey -match '^(HKLM|HKCU):') { $registryHive = $Matches[1] }
    if (-not $registryView -and $predictedKey -match '(?i)\\WOW6432Node\\') { $registryView = '32' }
    $scriptDerived = ($pkgMeta -and $pkgMeta.PSObject.Properties['HeaderAvailable'] -and $pkgMeta.HeaderAvailable)
    # The manifest decides the context only for an installer whose format
    # says nothing about it: Inno Setup, WiX Burn and BitRock stubs are
    # asInvoker and elevate themselves at run time, and NSIS and Inno
    # headers carry the context outright.
    $manifestDecides = ($type -eq 'Unknown')
    if (-not $installContext) {
        if ($registryHive -eq 'HKCU') { $installContext = 'PerUser' }
        elseif ($manifestDecides -and $fileInfo.PSObject.Properties['RequestedExecutionLevel'] -and $fileInfo.RequestedExecutionLevel -eq 'asInvoker') { $installContext = 'PerUser' }
        elseif ($registryHive -eq 'HKLM' -or $type -eq 'MSI') { $installContext = 'PerMachine' }
    }
    # A convention-predicted HKLM key for an installer that runs without
    # elevation lands under HKCU, where no WOW6432Node redirection applies.
    if ($installContext -eq 'PerUser' -and -not $scriptDerived -and $type -ne 'MSI' -and $predictedKey -match '^HKLM:') {
        $predictedKey = ($predictedKey -replace '^HKLM:', 'HKCU:') -replace '(?i)\\WOW6432Node\\', '\'
        $registryHive = 'HKCU'
        $registryView = ''
    }

    # Install modes: a switchable installer (NSIS with /allusers, Inno Setup
    # with the command-line override) carries one resolved branch per mode;
    # the top-level fields describe the default branch.
    $installModes = @()
    $installMode = ''
    $modeVariants = @{}
    if ($pkgMeta -and $pkgMeta.PSObject.Properties['InstallModes'] -and $pkgMeta.InstallModes) {
        $installModes = @($pkgMeta.InstallModes)
        $installMode = & $metaProp 'InstallMode'
        if ($pkgMeta.PSObject.Properties['ModeVariants'] -and $pkgMeta.ModeVariants) { $modeVariants = $pkgMeta.ModeVariants }
    }
    if (-not $installMode) { $installMode = if ($installContext -eq 'PerMachine') { 'AllUsers' } elseif ($installContext -eq 'PerUser') { 'CurrentUser' } else { '' } }

    # A Velopack setup stub's own PE architecture can differ from the
    # application it carries; the deployment describes the application.
    $architecture = if ($msiSummary -and $msiSummary.Architecture) { [string]$msiSummary.Architecture } else { [string]$fileInfo.Architecture }
    $metaArchitecture = & $metaProp 'Architecture'
    if ($type -eq 'Velopack' -and $metaArchitecture -in @('x86', 'x64', 'arm64')) { $architecture = $metaArchitecture }

    [pscustomobject]@{
        Path                 = (Resolve-Path -LiteralPath $Path).Path
        FileName             = Split-Path -Leaf $Path
        InstallerType        = $type
        AppName              = [string]$fields.DisplayName
        Publisher            = [string]$fields.Vendor
        SoftwareVersion      = [string]$fields.DisplayVersion
        ProductCode          = $productCode
        InstallArgs          = $installArgs
        UninstallArgs        = $uninstallArgs
        InstallCommand       = $installCommand
        UninstallCommand     = $uninstallCommand
        UninstallRegistryKey = $predictedKey
        UninstallRegistryHive = $registryHive
        RegistryView         = $registryView
        InstallDir           = $installDir
        InstallContext       = $installContext
        InstallMode          = $installMode
        InstallModes         = $installModes
        ModeVariants         = $modeVariants
        RequestedExecutionLevel = if ($fileInfo.PSObject.Properties['RequestedExecutionLevel']) { [string]$fileInfo.RequestedExecutionLevel } else { '' }
        Architecture         = $architecture
        Confidence           = if ($type -eq 'MSI') { 'Authoritative' } else { 'Predicted' }
        Switches             = $switches
        Fields               = $fields
        PackageMetadata      = $pkgMeta
    }
}

function Assert-ArpDetectionKey {
    <#
    .SYNOPSIS
        Throws when the staged installer names a different ARP key or registry
        view than the detection clause.
    .DESCRIPTION
        The clause carries a literal key confirmed on a client. A vendor that
        renames the key or moves it between registry views in a later build
        would otherwise stage a rule that can never evaluate true.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$ExpectedKey,
        [Parameter(Mandatory)][bool]$Is64BitView
    )

    $analysis = Get-InstallerAnalysis -Path $InstallerPath
    $decoded  = [string]$analysis.UninstallRegistryKey
    if ([string]::IsNullOrWhiteSpace($decoded)) {
        Write-Log "Installer header names no ARP key; detection keeps the confirmed literal." -Level WARN
        return
    }

    $relative = $decoded -replace '^HK(LM|CU):\\', ''
    $relative = $relative -replace '(?i)^SOFTWARE\\WOW6432Node\\', 'SOFTWARE\'
    if ($relative -ne $ExpectedKey) {
        throw "Installer registers '$relative' but detection expects '$ExpectedKey'."
    }

    $decodedView = [string]$analysis.RegistryView
    if ($decodedView) {
        $decoded64 = ($decodedView -ne '32')
        if ($decoded64 -ne $Is64BitView) {
            throw "Installer writes its ARP key in the $decodedView-bit view but detection expects the other view."
        }
    }
}

Export-ModuleMember -Function Assert-ArpDetectionKey
function Set-InstallerAnalysisMode {
    <#
    .SYNOPSIS
        Returns a copy of an installer analysis with every mode-dependent
        field taken from one install mode's branch.
    .DESCRIPTION
        Install arguments and command, uninstall command, install folder,
        ARP key with hive and view, and install context all come from the
        branch the installer script resolves for that mode, so a stage built
        from the copy never pairs a per-machine install with a per-user
        uninstall or detection. Throws when the analysis does not offer the
        requested mode.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject]$Analysis,
        [Parameter(Mandatory)][ValidateSet('CurrentUser', 'AllUsers')][string]$Mode
    )

    $variants = if ($Analysis.PSObject.Properties['ModeVariants'] -and $Analysis.ModeVariants) { $Analysis.ModeVariants } else { @{} }
    if (-not $variants.ContainsKey($Mode)) {
        throw ("Installer '{0}' offers install mode(s) {1}, not {2}." -f $Analysis.FileName, (@($Analysis.InstallModes) -join ', '), $Mode)
    }
    $branch = $variants[$Mode]

    $copy = $Analysis.PSObject.Copy()
    $copy.InstallMode          = $Mode
    $copy.InstallArgs          = [string]$branch.InstallArgs
    $copy.InstallCommand       = ('"' + $Analysis.FileName + '" ' + [string]$branch.InstallArgs).Trim()
    $copy.UninstallCommand     = [string]$branch.SilentUninstallCommand
    $copy.UninstallRegistryKey = [string]$branch.UninstallRegistryKey
    $copy.UninstallRegistryHive = [string]$branch.RegistryHive
    $copy.RegistryView         = [string]$branch.RegistryView
    $copy.InstallDir           = [string]$branch.InstallDirWindows
    $copy.InstallContext       = [string]$branch.InstallContext
    return $copy
}

function New-AdHocStage {
    <#
    .SYNOPSIS
        Stages a dropped installer as a versioned content folder with
        wrappers and a stage manifest - the same shape every packager
        produces, so Package consumes it unchanged.

    .DESCRIPTION
        Layout: <DownloadRoot>\<AppFolder>\<Version>\ holding the installer,
        the install/uninstall wrapper set, and stage-manifest.json with
        SHA256 hashes. Detection is RegistryKeyValue on the ARP key: the
        ProductCode key for an MSI (deterministic), or the analysis-
        predicted uninstall key otherwise (operator-confirmed upstream).
    #>
    param(
        [Parameter(Mandatory)]$Analysis,
        [Parameter(Mandatory)][string]$DownloadRoot,
        [string]$AppName,
        [string]$Publisher,
        [string]$SoftwareVersion,
        [string]$InstallArgs,
        [string]$UninstallArgs,
        [string]$UninstallCommand
    )

    if (-not $AppName)         { $AppName         = [string]$Analysis.AppName }
    if (-not $Publisher)       { $Publisher       = [string]$Analysis.Publisher }
    if (-not $SoftwareVersion) { $SoftwareVersion = [string]$Analysis.SoftwareVersion }
    if (-not $InstallArgs)     { $InstallArgs     = [string]$Analysis.InstallArgs }
    if (-not $PSBoundParameters.ContainsKey('UninstallArgs') -or $null -eq $UninstallArgs) { $UninstallArgs = [string]$Analysis.UninstallArgs }
    if ([string]::IsNullOrWhiteSpace($AppName))         { throw 'Ad-hoc stage requires an application name.' }
    if ([string]::IsNullOrWhiteSpace($SoftwareVersion)) { throw 'Ad-hoc stage requires a version.' }
    if ($Analysis.InstallerType -eq 'MSP') {
        throw 'MSI patches (.msp) are not supported for ad-hoc staging; they apply against an installed base MSI, not as a standalone deployment.'
    }
    if ($Analysis.InstallerType -ne 'MSI' -and [string]::IsNullOrWhiteSpace($InstallArgs)) {
        throw 'Ad-hoc stage requires install arguments for a non-MSI installer.'
    }
    if ([string]::IsNullOrWhiteSpace($Publisher))       { $Publisher = 'Unknown' }

    $installerName = [string]$Analysis.FileName
    # Wrapper .ps1 files are written -Encoding ASCII (deployment targets read
    # them with Windows PowerShell defaults); a non-ASCII filename would be
    # silently mangled inside the wrapper and fail only on target devices.
    if ($installerName -match '[^\x00-\x7F]') {
        throw "Installer filename contains non-ASCII characters; rename the file to plain ASCII before staging: $installerName"
    }

    # Folder-name sanitation: identity fields become path segments. A value
    # that sanitizes to nothing must not silently collapse into the parent
    # folder (Join-Path with '' returns the parent).
    $vendorFolder = ConvertTo-SafePathSegment $Publisher
    $appFolder    = ConvertTo-SafePathSegment $AppName
    $versionFolder = ConvertTo-SafePathSegment $SoftwareVersion
    if ([string]::IsNullOrWhiteSpace($vendorFolder))  { $vendorFolder = 'Unknown' }
    if ([string]::IsNullOrWhiteSpace($appFolder))     { throw "Application name '$AppName' contains no usable path characters." }
    if ([string]::IsNullOrWhiteSpace($versionFolder)) { throw "Version '$SoftwareVersion' contains no usable path characters." }

    $baseRoot = Join-Path (Join-Path $DownloadRoot $vendorFolder) $appFolder
    $stagedPath = Join-Path $baseRoot $versionFolder
    Assert-PathUnderRoot -Path $stagedPath -Root $DownloadRoot
    Initialize-Folder -Path $stagedPath
    $stagedInstaller = Join-Path $stagedPath $installerName
    Copy-Item -LiteralPath $Analysis.Path -Destination $stagedInstaller -Force
    Write-Log "Staged dropped installer     : $stagedInstaller"

    $isMsi = ($Analysis.InstallerType -eq 'MSI')
    # New-ExeWrapperContent interpolates InstallArgs/UninstallArgs into an
    # @() literal, so plain space-separated args become a quoted element list.
    $toArgList = {
        param($s)
        (($s -split '\s+') | Where-Object { $_ } | ForEach-Object { "'" + (ConvertTo-SingleQuotedContent $_) + "'" }) -join ', '
    }
    if ($isMsi) {
        # The wrapper already passes /i, /x, /qn and /norestart; whatever else
        # the operator entered (MSI properties, transforms) rides along.
        $toMsiExtras = {
            param($s)
            @([regex]::Matches([string]$s, '(?:[^\s"]+|"[^"]*")+') | ForEach-Object { $_.Value } |
                Where-Object { $_ -notmatch '^(?i)[/-](qn|qb|qr|qf|q|quiet|passive|norestart|i|x)$' })
        }
        $wrapperContent = New-MsiWrapperContent -MsiFileName $installerName `
            -ExtraInstallArgs (& $toMsiExtras $InstallArgs) `
            -ExtraUninstallArgs (& $toMsiExtras $UninstallArgs)
    }
    else {
        if ([string]::IsNullOrWhiteSpace($UninstallCommand)) { $UninstallCommand = [string]$Analysis.UninstallCommand }
        # Split "path args" into -FilePath and -ArgumentList halves; the path
        # may be quoted (analysis emits '"uninstall.exe" /S').
        $uninstallExe  = $UninstallCommand
        $uninstallTail = $UninstallArgs
        if ($UninstallCommand -match '^\s*"(?<exe>[^"]+)"\s*(?<tail>.*)$' -or
            $UninstallCommand -match '^\s*(?<exe>\S+)\s+(?<tail>.+)$') {
            $uninstallExe = $Matches['exe']
            if ([string]::IsNullOrWhiteSpace($uninstallTail)) { $uninstallTail = $Matches['tail'].Trim() }
        }
        if ([string]::IsNullOrWhiteSpace($uninstallExe)) {
            # Wrapper still ships; the operator-visible manifest carries the
            # gap so the generated uninstall is an explicit no-op, not a lie.
            $uninstallExe = 'cmd.exe'
            $uninstallTail = '/c exit 0'
        }
        $wrapperContent = New-ExeWrapperContent -InstallerFileName $installerName `
            -InstallArgs (& $toArgList $InstallArgs) `
            -UninstallCommand $uninstallExe `
            -UninstallArgs (& $toArgList $uninstallTail)
    }
    Write-ContentWrappers -OutputPath $stagedPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    $is64 = ([string]$Analysis.Architecture) -notmatch '^(x86|Intel)$'
    $hive = 'LocalMachine'
    if ($isMsi -and $Analysis.ProductCode) {
        $arpRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $Analysis.ProductCode
    }
    else {
        # Analysis predicts a full hive path; the manifest wants the
        # hive-relative form. A WOW6432Node prediction means a 32-bit view.
        $predictedKey = [string]$Analysis.UninstallRegistryKey
        if ($predictedKey -match '^HKCU:') { $hive = 'CurrentUser' }
        $arpRelative = $predictedKey -replace '^HK(LM|CU):\\', ''
        if ($arpRelative -match '(?i)\\?WOW6432Node\\') {
            $arpRelative = $arpRelative -replace '(?i)WOW6432Node\\', ''
            $is64 = $false
        }
        elseif ($Analysis.PSObject.Properties['RegistryView'] -and [string]$Analysis.RegistryView -eq '32') {
            $is64 = $false
        }
        if ([string]::IsNullOrWhiteSpace($arpRelative)) {
            $arpRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $AppName
        }
    }

    # A per-user installer writes under the invoking user's profile and
    # HKCU, so the deployment type must install and detect in the user
    # context; a system-context run would land in the SYSTEM profile and
    # never satisfy an HKCU detection.
    $perUser = ($Analysis.PSObject.Properties['InstallContext'] -and [string]$Analysis.InstallContext -eq 'PerUser') -or ($hive -eq 'CurrentUser')
    if ($perUser -and $hive -ne 'CurrentUser' -and -not $isMsi -and $arpRelative -match '(?i)^SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\') {
        $hive = 'CurrentUser'
    }

    $manifest = @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $SoftwareVersion
        InstallerFile   = $installerName
        InstallerType   = if ($isMsi) { 'MSI' } else { 'EXE' }
        DetectedEngine  = [string]$Analysis.InstallerType
        InstallArgs     = $InstallArgs
        UninstallArgs   = $UninstallArgs
        UninstallCommand = if ($isMsi) { '' } else { $UninstallCommand }
        ProductCode     = [string]$Analysis.ProductCode
        # The workbench and the Intune adapter read architecture off the
        # manifest; a BYO app has no packager script header to fall back on.
        Architecture    = $(if ([string]$Analysis.Architecture -match '^(?i)(arm64|aarch64)$') { 'arm64' } elseif ($is64) { 'x64' } else { 'x86' })
        RunningProcess  = @()
        AdHocSource     = [string]$Analysis.Path
        InstallContext  = if ($perUser) { 'PerUser' } else { 'PerMachine' }
        Detection       = @{
            Type                = 'RegistryKeyValue'
            Hive                = $hive
            RegistryKeyRelative = $arpRelative
            ValueName           = 'DisplayVersion'
            DisplayName         = $AppName
            DisplayVersion      = $SoftwareVersion
            Is64Bit             = $is64
        }
    }
    if ($perUser) {
        $manifest['InstallationBehaviorType'] = 'InstallForUser'
        $manifest['LogonRequirementType']     = 'OnlyWhenUserLoggedOn'
    }

    $manifestPath = Join-Path $stagedPath 'stage-manifest.json'
    Write-StageManifest -Path $manifestPath -ManifestData $manifest
    Write-Log ("Install context              : {0}" -f $manifest['InstallContext'])
    Write-Log "Ad-hoc stage complete        : $stagedPath"

    [pscustomobject]@{
        StagedPath   = $stagedPath
        ManifestPath = $manifestPath
        VendorFolder = $vendorFolder
        AppFolder    = $appFolder
    }
}

function Invoke-AdHocPackage {
    <#
    .SYNOPSIS
        Packages an ad-hoc staged folder: copies content to the network
        version folder and creates the ConfigMgr application from the manifest -
        the same path every packager takes.
    #>
    param(
        [Parameter(Mandatory)][string]$StagedPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder,
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$Comment = '',
        [ValidateSet('Nested', 'Flat')][string]$ContentLayout = 'Nested',
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30
    )

    # The caller may already sit on the site drive (a background runspace that
    # connected before this call). Local and UNC paths resolved through the
    # site provider collapse to a null path, so the manifest read, the network
    # sync and the integrity check all run from the file system.
    $origLocation = Get-Location
    try {
        Set-Location -LiteralPath "$env:SystemDrive\" -ErrorAction Stop

        $manifest = Read-StageManifest -Path (Join-Path $StagedPath 'stage-manifest.json')
        if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
            throw "Network root path not accessible: $FileServerPath"
        }

        # Version becomes a path segment; strip the same characters the stage
        # sanitized so local and network folder names agree.
        $versionSegment = ConvertTo-SafePathSegment ([string]$manifest.SoftwareVersion)
        if ([string]::IsNullOrWhiteSpace($versionSegment)) {
            throw ("Manifest version '{0}' contains no usable path characters." -f $manifest.SoftwareVersion)
        }
        $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath `
            -VendorFolder $VendorFolder -AppFolder $AppFolder `
            -Version $versionSegment -Layout $ContentLayout
        Initialize-Folder -Path $networkContentPath

        Sync-StagedContentToNetwork -LocalContentPath $StagedPath -NetworkContentPath $networkContentPath -Manifest $manifest
        Write-Log "Ad-hoc content on network    : $networkContentPath"

        return New-MECMApplicationFromManifest `
            -Manifest $manifest `
            -SiteCode $SiteCode `
            -Comment $Comment `
            -NetworkContentPath $networkContentPath `
            -EstimatedRuntimeMins $EstimatedRuntimeMins `
            -MaximumRuntimeMins $MaximumRuntimeMins
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function New-PackagerFromDrop {
    <#
    .SYNOPSIS
        Graduates an analyzed drop into the packager catalog: writes
        Packagers\package-<slug>.ps1 from the matching template with the
        analysis-filled identity values.

    .DESCRIPTION
        Identity, folder segments, and installer filename are baked; the
        download-source resolution stays a TODO by nature - a dropped file
        carries no URL. The generated file satisfies the grid's discovery
        contract immediately (header metadata, GetLatestVersionOnly, phase
        switches) and refuses to overwrite an existing packager.
    #>
    param(
        [Parameter(Mandatory)]$Analysis,
        [Parameter(Mandatory)][string]$PackagersRoot,
        [string]$AppName,
        [string]$Publisher,
        [string]$SoftwareVersion
    )

    if (-not $AppName)         { $AppName         = [string]$Analysis.AppName }
    if (-not $Publisher)       { $Publisher       = [string]$Analysis.Publisher }
    if (-not $SoftwareVersion) { $SoftwareVersion = [string]$Analysis.SoftwareVersion }
    if ([string]::IsNullOrWhiteSpace($AppName)) { throw 'A packager needs an application name.' }

    $isMsi = ($Analysis.InstallerType -eq 'MSI')
    $templateName = if ($isMsi) { 'package-template-msi.ps1' } else { 'package-template-exe.ps1' }
    $templatePath = Join-Path (Split-Path $PackagersRoot -Parent) (Join-Path 'Samples' $templateName)
    if (-not (Test-Path -LiteralPath $templatePath)) { throw "Template not found: $templatePath" }

    $slug = (($AppName.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-'))
    $target = Join-Path $PackagersRoot ("package-{0}.ps1" -f $slug)
    if (Test-Path -LiteralPath $target) { throw "Packager already exists: $target" }

    $vendorFolder = ConvertTo-SafePathSegment $Publisher
    if ([string]::IsNullOrWhiteSpace($vendorFolder)) { $vendorFolder = 'Unknown' }
    $appFolder    = ConvertTo-SafePathSegment $AppName
    if ([string]::IsNullOrWhiteSpace($appFolder)) { throw "Application name '$AppName' contains no usable path characters." }

    # Identity values originate in the dropped file's own version resource /
    # MSI tables: untrusted text. Substitutions into live double-quoted
    # literals must neutralize `, $, and every character PowerShell reads as a
    # double quote (the typographic quotes U+201C to U+201F included) or the
    # resource string executes as code the next time the generated packager
    # runs; substitutions into the doc comment must not carry a comment
    # terminator. Line breaks are flattened so a value stays on its own line.
    $escapeDq = {
        param($s)
        $flat = [string]$s -replace '[\r\n\u0085\u2028\u2029]+', ' '
        (($flat -replace '`', '``') -replace '\$', '`$') -replace '(["\u201C\u201D\u201E\u201F])', '`$1'
    }
    $escapeComment = { param($s) ([string]$s -replace '[\r\n\u0085\u2028\u2029]+', ' ') -replace '#>', '# >' }
    $appNameDq   = & $escapeDq $AppName
    $publisherDq = & $escapeDq $Publisher
    $appNameCmt   = & $escapeComment $AppName
    $publisherCmt = & $escapeComment $Publisher

    # Literal String.Replace throughout: values may contain characters regex
    # replacement strings would reinterpret.
    $c = Get-Content -LiteralPath $templatePath -Raw
    $c = $c.Replace('Vendor: TODO', "Vendor: $publisherCmt")
    $c = $c.Replace('App: TODO', "App: $appNameCmt")
    $c = $c.Replace('CMName: TODO', "CMName: $appNameCmt")
    $c = $c.Replace('$VendorFolder = "TODO"', ('$VendorFolder = "{0}"' -f (& $escapeDq $vendorFolder)))
    $c = $c.Replace('$AppFolder    = "TODO"', ('$AppFolder    = "{0}"' -f (& $escapeDq $appFolder)))
    $c = $c.Replace('Packages TODO (x64)', "Packages $appNameCmt")
    $c = $c.Replace('"TODO - STAGE phase"', ('"{0} - STAGE phase"' -f $appNameDq))
    $c = $c.Replace('"TODO - PACKAGE phase"', ('"{0} - PACKAGE phase"' -f $appNameDq))
    $c = $c.Replace('"TODO Auto-Packager starting"', ('"{0} Auto-Packager starting"' -f $appNameDq))
    if ($isMsi) {
        $c = $c.Replace('$MsiFileName      = "TODO-installer.msi"', ('$MsiFileName      = "{0}"' -f (& $escapeDq $Analysis.FileName)))
    }
    else {
        $c = $c.Replace('AppName          = "TODO"', ('AppName          = "{0}"' -f $appNameDq))
        $c = $c.Replace('Publisher        = "TODO"', ('Publisher        = "{0}"' -f $publisherDq))
        if (-not [string]::IsNullOrWhiteSpace([string]$Analysis.InstallArgs)) {
            $c = $c.Replace('$installArgs   = "/S"', ('$installArgs   = "{0}"' -f (& $escapeDq $Analysis.InstallArgs)))
        }
        # An uninstall command the analysis resolved to a real path ("<path>"
        # /S) fills the uninstall pair; a bare placeholder stays a TODO.
        $uninstallCommand = [string]$Analysis.UninstallCommand
        if ($uninstallCommand -match '^\s*"(?<path>[^"]+)"\s*(?<args>.*)$' -and $Matches['path'] -match '[\\%]') {
            $c = $c.Replace('$uninstallCmd  = "C:\Program Files\TODO\uninstall.exe /S"      # TODO: confirm uninstall path',
                            ('$uninstallCmd  = "{0}"' -f (& $escapeDq $Matches['path'])))
            $c = $c.Replace('$uninstallArgs = ""', ('$uninstallArgs = "{0}"' -f (& $escapeDq $Matches['args'].Trim())))
        }
        if ($Analysis.PSObject.Properties['InstallDir'] -and [string]$Analysis.InstallDir -match '^(%[^%]+%|[A-Za-z]:\\)') {
            $c = $c.Replace('$detectionPath = "C:\Program Files\TODO"', ('$detectionPath = "{0}"' -f (& $escapeDq $Analysis.InstallDir)))
        }
        if ($Analysis.PSObject.Properties['InstallContext'] -and [string]$Analysis.InstallContext -eq 'PerUser') {
            $c = $c.Replace('InstallationBehaviorType = "InstallForSystem"', 'InstallationBehaviorType = "InstallForUser"')
            $c = $c.Replace('LogonRequirementType     = "WhetherOrNotUserLoggedOn"', 'LogonRequirementType     = "OnlyWhenUserLoggedOn"')
        }
    }

    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($c, [ref]$null, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw ("Generated packager does not parse ({0}); refusing to write it. First error: {1}" -f $parseErrors.Count, $parseErrors[0].Message)
    }

    Set-Content -LiteralPath $target -Value $c -Encoding UTF8
    Write-Log "Packager written             : $target (download source is a TODO; identity and folders are filled)"
    return $target
}


# ---------------------------------------------------------------------------
# Module export (belt-and-suspenders with .psd1 FunctionsToExport)
# ---------------------------------------------------------------------------

Export-ModuleMember -Function *

function Get-RequestedPackagerVariants {
    <#
    .SYNOPSIS
        Parses the APP_PACKAGER_VARIANTS environment JSON the GUI sets for
        a SupportsVariants packager.

    .DESCRIPTION
        Returns $null when no variant split is requested. Malformed JSON
        throws instead of silently packaging without the operator's split.

    .OUTPUTS
        [pscustomobject] Split ('Architecture' | 'Language' | 'Network')
        and Languages (culture codes; only meaningful for a Language
        split), or $null.
    #>
    $envJson = $env:APP_PACKAGER_VARIANTS
    if ([string]::IsNullOrWhiteSpace($envJson)) { return $null }
    try {
        $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "APP_PACKAGER_VARIANTS is not valid JSON: $($_.Exception.Message)"
    }
    $split = [string]$parsed.Split
    if ($split -notin @('Architecture', 'Language', 'Network')) {
        throw "APP_PACKAGER_VARIANTS Split must be Architecture, Language, or Network (got '$split')."
    }
    $languages = @()
    if ($parsed.PSObject.Properties['Languages'] -and $parsed.Languages) {
        $languages = @($parsed.Languages | ForEach-Object { [string]$_ } | Where-Object { $_ })
    }
    if ($split -eq 'Language' -and $languages.Count -eq 0) {
        throw 'APP_PACKAGER_VARIANTS requests a Language split with no Languages.'
    }
    return [pscustomobject]@{ Split = $split; Languages = $languages }
}

Export-ModuleMember -Function Get-RequestedPackagerVariants

function Get-RequestedCommandOverrides {
    <#
    .SYNOPSIS
        Parses the APP_PACKAGER_COMMANDS environment JSON the GUI sets
        when the operator overrides an app's install or uninstall command.

    .DESCRIPTION
        Returns $null when no override is configured. Malformed JSON or an
        override with neither command throws instead of silently packaging
        with the wrong commands.

    .OUTPUTS
        [pscustomobject] Install and Uninstall (either may be empty), or
        $null.
    #>
    $envJson = $env:APP_PACKAGER_COMMANDS
    if ([string]::IsNullOrWhiteSpace($envJson)) { return $null }
    try {
        $parsed = $envJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "APP_PACKAGER_COMMANDS is not valid JSON: $($_.Exception.Message)"
    }
    $install = [string]$parsed.Install
    $uninstall = [string]$parsed.Uninstall
    if ([string]::IsNullOrWhiteSpace($install) -and [string]::IsNullOrWhiteSpace($uninstall)) {
        throw 'APP_PACKAGER_COMMANDS carries neither an Install nor an Uninstall command.'
    }
    return [pscustomobject]@{ Install = $install.Trim(); Uninstall = $uninstall.Trim() }
}

Export-ModuleMember -Function Get-RequestedCommandOverrides

function Get-RequestedInstallMode {
    <#
    .SYNOPSIS
        Reads the APP_PACKAGER_INSTALL_MODE environment value the GUI sets for
        a SupportsInstallModes packager.

    .DESCRIPTION
        Returns 'CurrentUser' or 'AllUsers', or $null when no mode is
        requested. Any other value throws instead of staging the packager's
        default mode against the operator's choice.
    #>
    $value = [string]$env:APP_PACKAGER_INSTALL_MODE
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $value = $value.Trim()
    if ($value -notin @('CurrentUser', 'AllUsers')) {
        throw "APP_PACKAGER_INSTALL_MODE must be CurrentUser or AllUsers (got '$value')."
    }
    return $value
}

Export-ModuleMember -Function Get-RequestedInstallMode

function Set-StageManifestInstallMode {
    <#
    .SYNOPSIS
        Rewrites a stage manifest and its wrappers for the requested install
        mode of an installer that accepts a mode switch.

    .DESCRIPTION
        Reads the staged installer with Get-InstallerAnalysis and takes the
        requested mode's branch: install arguments gain that mode's switch
        (the other mode's switch is removed), the uninstall command moves to
        that branch's uninstaller path, registry detection follows the
        branch's hive and view, per-user file detection becomes HKCU uninstall
        version detection, system file detection follows its folder, and a per-user branch sets
        InstallationBehaviorType InstallForUser. Wrappers in the standard
        New-ExeWrapperContent shape are regenerated; other wrappers only
        have the switch token replaced. Throws when the installer does not
        offer the requested mode, so a package never ships in a mode the
        operator did not choose.

    .OUTPUTS
        $true when the manifest was rewritten.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [Parameter(Mandatory)][ValidateSet('CurrentUser', 'AllUsers')][string]$Mode
    )

    if ([string]$ManifestData['InstallerType'] -ne 'EXE') {
        Write-Log "Install mode $Mode requested; only EXE installers carry a mode switch, manifest left unchanged." -Level WARN
        return $false
    }
    $installerName = [string]$ManifestData['InstallerFile']
    $installerPath = Join-Path $StageRoot $installerName
    if ([string]::IsNullOrWhiteSpace($installerName) -or -not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        throw "Install mode $Mode requested but the staged installer '$installerName' was not found under $StageRoot."
    }

    $analysis = Get-InstallerAnalysis -Path $installerPath
    $modes = @($analysis.InstallModes)
    if ($modes -notcontains $Mode) {
        $offered = if ($modes.Count -gt 0) { $modes -join ', ' } else { 'no mode switch' }
        throw "Install mode $Mode requested but $installerName offers $offered."
    }
    $branch = Set-InstallerAnalysisMode -Analysis $analysis -Mode $Mode
    $meta = $analysis.PackageMetadata
    $switchFor = @{ AllUsers = [string]$meta.AllUsersSwitch; CurrentUser = [string]$meta.CurrentUserSwitch }
    $wanted = $switchFor[$Mode]
    $knownSwitches = @($switchFor.Values | Where-Object { $_ })
    $isNsis = ([string]$analysis.InstallerType -eq 'NSIS')

    # Token-level swap keeps every other argument the packager authored.
    $swapSwitch = {
        param([string]$arguments, [bool]$append)
        $tokens = @(($arguments -split '\s+') | Where-Object { $_ -and ($knownSwitches -notcontains $_) })
        if ($append -and $wanted) { $tokens += $wanted }
        return ($tokens -join ' ')
    }
    $originalInstallArgs = [string]$ManifestData['InstallArgs']
    $originalUninstallArgs = [string]$ManifestData['UninstallArgs']
    $ManifestData['InstallArgs'] = & $swapSwitch $originalInstallArgs $true
    # NSIS uninstallers take the same switch; an Inno Setup uninstaller reads its mode from the install.
    $ManifestData['UninstallArgs'] = & $swapSwitch $originalUninstallArgs $isNsis

    $originalUninstallCommand = [string]$ManifestData['UninstallCommand']
    $branchUninstall = [string]$branch.UninstallCommand
    $newUninstallCommand = $originalUninstallCommand
    if ($branchUninstall -match '^\s*"([^"]+)"' -or $branchUninstall -match '^\s*(\S+\.exe)') {
        $newUninstallCommand = $Matches[1]
        $ManifestData['UninstallCommand'] = $newUninstallCommand
    }

    $folderFor = @{}
    foreach ($m in @('AllUsers', 'CurrentUser')) {
        if ($analysis.ModeVariants -and $analysis.ModeVariants[$m] -and $analysis.ModeVariants[$m].InstallDirWindows) {
            $folderFor[$m] = [string]$analysis.ModeVariants[$m].InstallDirWindows
        }
    }
    $otherMode = if ($Mode -eq 'AllUsers') { 'CurrentUser' } else { 'AllUsers' }

    $det = $ManifestData['Detection']
    if ($det -is [hashtable]) {
        $detType = [string]$det['Type']
        switch -Regex ($detType) {
            '^RegistryKey(Value)?$' {
                $hive = [string]$branch.UninstallRegistryHive
                if ($hive -eq 'HKCU') {
                    $det['Hive'] = 'CurrentUser'
                }
                else {
                    $det['Hive'] = 'LocalMachine'
                    $det['Is64Bit'] = ([string]$branch.RegistryView -ne '32')
                }
                $branchKey = [string]$branch.UninstallRegistryKey
                if ($branchKey -and $branchKey -notmatch '[\$\{]') {
                    $relative = ($branchKey -replace '^HK(LM|CU):\\', '') -replace '(?i)^Software\\WOW6432Node\\', 'SOFTWARE\'
                    $current = [string]$det['RegistryKeyRelative']
                    if ($current -and ($current -replace '(?i)^Software\\WOW6432Node\\', 'SOFTWARE\') -ine $relative) {
                        $det['RegistryKeyRelative'] = $relative
                    }
                }
            }
            '^File$' {
                if ([string]$branch.InstallContext -eq 'PerUser') {
                    # Do not resolve a profile path in the packaging or SYSTEM context.
                    # Use only a concrete uninstall key supplied by installer analysis.
                    $userKey = [string]$branch.UninstallRegistryKey
                    $keySuffix = $userKey -replace '^HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\', ''
                    $unresolved = $keySuffix -replace '\{[0-9a-fA-F-]{36}\}', ''
                    if ([string]$branch.UninstallRegistryHive -ne 'HKCU' -or
                        $keySuffix -eq $userKey -or [string]::IsNullOrWhiteSpace($keySuffix) -or
                        $unresolved -match '[\$\{\}\*\?%]' -or
                        [string]::IsNullOrWhiteSpace([string]$ManifestData['SoftwareVersion'])) {
                        throw "Cannot create user detection for ${installerName}: installer analysis did not provide a concrete HKCU uninstall key and application version."
                    }
                    $ManifestData['Detection'] = @{
                        Type = 'RegistryKeyValue'
                        Hive = 'CurrentUser'
                        RegistryKeyRelative = ($userKey -replace '^HKCU:\\', '')
                        ValueName = 'DisplayVersion'
                        PropertyType = 'Version'
                        Operator = 'GreaterEquals'
                        ExpectedValue = [string]$ManifestData['SoftwareVersion']
                        Is64Bit = $false
                    }
                    break
                }
                $filePath = [string]$det['FilePath']
                $mapped = $false
                if ($filePath -and $folderFor[$Mode] -and $folderFor[$otherMode]) {
                    $otherExpanded = [Environment]::ExpandEnvironmentVariables($folderFor[$otherMode]).TrimEnd('\')
                    $otherRaw = $folderFor[$otherMode].TrimEnd('\')
                    foreach ($prefix in @($otherExpanded, $otherRaw)) {
                        if ($filePath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $det['FilePath'] = $folderFor[$Mode].TrimEnd('\') + $filePath.Substring($prefix.Length)
                            $mapped = $true
                            break
                        }
                    }
                    if (-not $mapped) {
                        $thisExpanded = [Environment]::ExpandEnvironmentVariables($folderFor[$Mode]).TrimEnd('\')
                        if ($filePath.StartsWith($thisExpanded, [System.StringComparison]::OrdinalIgnoreCase) -or $filePath.StartsWith($folderFor[$Mode].TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) { $mapped = $true }
                    }
                }
                if (-not $mapped) {
                    Write-Log ("File detection path '{0}' is not under either install folder; left unchanged for mode {1}." -f $filePath, $Mode) -Level WARN
                }
            }
            default {
                Write-Log "Detection type '$detType' is not adjusted for install mode $Mode; review it before deploying." -Level WARN
            }
        }
    }

    $perUser = ([string]$branch.InstallContext -eq 'PerUser')
    if ($perUser) {
        $ManifestData['InstallationBehaviorType'] = 'InstallForUser'
        $ManifestData['LogonRequirementType']     = 'OnlyWhenUserLoggedOn'
    }
    else {
        $ManifestData['InstallationBehaviorType'] = 'InstallForSystem'
        if ($ManifestData.ContainsKey('LogonRequirementType')) { $ManifestData.Remove('LogonRequirementType') }
    }
    $ManifestData['InstallMode']    = $Mode
    $ManifestData['InstallContext'] = [string]$branch.InstallContext

    # Wrappers: regenerate the standard shape; otherwise swap the switch token in place.
    $installPs1   = Join-Path $StageRoot 'install.ps1'
    $uninstallPs1 = Join-Path $StageRoot 'uninstall.ps1'
    $standardInstall   = '^\$exePath = Join-Path \$PSScriptRoot ''.*''\r?\n\$proc = Start-Process -FilePath \$exePath -ArgumentList @\(.*\) -Wait -PassThru -NoNewWindow\r?\nexit \$proc\.ExitCode\s*$'
    $standardUninstall = '^\$uninstallPath = \[Environment\]::ExpandEnvironmentVariables\(''.*''\)\r?\n\$proc = Start-Process -FilePath \$uninstallPath( -ArgumentList @\(.*\))? -Wait -PassThru -NoNewWindow\r?\nexit \$proc\.ExitCode\s*$'
    $installText   = if (Test-Path -LiteralPath $installPs1)   { [IO.File]::ReadAllText($installPs1) }   else { '' }
    $uninstallText = if (Test-Path -LiteralPath $uninstallPs1) { [IO.File]::ReadAllText($uninstallPs1) } else { '' }
    $quote = { param([string]$s) "'" + (ConvertTo-SingleQuotedContent $s) + "'" }
    if ((Remove-StartProcessFailFast $installText) -match $standardInstall -and (Remove-StartProcessFailFast $uninstallText) -match $standardUninstall) {
        $wrappers = New-ExeWrapperContent -InstallerFileName $installerName `
            -InstallArgs (& $quote ([string]$ManifestData['InstallArgs'])) `
            -UninstallCommand $newUninstallCommand `
            -UninstallArgs $(if ([string]$ManifestData['UninstallArgs']) { & $quote ([string]$ManifestData['UninstallArgs']) } else { '' })
        $installBat = Join-Path $StageRoot 'install.bat'
        $uninstallBat = Join-Path $StageRoot 'uninstall.bat'
        $batCode = { param([string]$path) if (Test-Path -LiteralPath $path) { $m = [regex]::Match([IO.File]::ReadAllText($path), 'EQU 0 exit /b (\S+)'); if ($m.Success) { return $m.Groups[1].Value } }; return '%ERRORLEVEL%' }
        Write-ContentWrappers -OutputPath $StageRoot -InstallPs1Content $wrappers.Install -UninstallPs1Content $wrappers.Uninstall `
            -InstallBatExitCode (& $batCode $installBat) -UninstallBatExitCode (& $batCode $uninstallBat)
    }
    else {
        foreach ($entry in @(@{ Path = $installPs1; Text = $installText }, @{ Path = $uninstallPs1; Text = $uninstallText })) {
            if (-not $entry.Text) { continue }
            $updated = $entry.Text
            foreach ($known in $knownSwitches) {
                if ($known -ine $wanted) { $updated = [regex]::Replace($updated, '(?i)' + [regex]::Escape($known), $wanted) }
            }
            if ($updated -ne $entry.Text) { [IO.File]::WriteAllText($entry.Path, $updated, [System.Text.Encoding]::ASCII) }
        }
        if ($newUninstallCommand -ne $originalUninstallCommand) {
            Write-Log ("Uninstall wrapper is packager-authored; it still targets '{0}' while mode {1} uninstalls through '{2}'." -f $originalUninstallCommand, $Mode, $newUninstallCommand) -Level WARN
        }
    }

    Write-Log ("Install mode applied         : {0} ({1})" -f $Mode, $(if ($wanted) { $wanted } else { 'installer default' }))
    Write-Log ("Install args (mode)          : {0}" -f $ManifestData['InstallArgs'])
    if ($newUninstallCommand) { Write-Log ("Uninstall command (mode)     : {0} {1}" -f $newUninstallCommand, $ManifestData['UninstallArgs']) }
    Write-Log ("Install context (mode)       : {0}" -f $ManifestData['InstallContext'])
    return $true
}

Export-ModuleMember -Function Set-StageManifestInstallMode

# Prefix of the machine-readable line emitted when a same-version existing
# application is skipped. The GUI matches on this to offer an overwrite re-run.
$script:OnExistingConflictMarker = '[APP_PACKAGER_CONFLICT] existing-same-version'

function Get-OnExistingConflictMarker {
    <#
    .SYNOPSIS
        Returns the marker prefix that flags a skipped same-version
        existing application in packager output.
    #>
    return $script:OnExistingConflictMarker
}

function Resolve-OnExistingBehavior {
    <#
    .SYNOPSIS
        Resolves what to do when an application of the same name
        already exists in the site.

    .DESCRIPTION
        Precedence is explicit parameter, then the APP_PACKAGER_ON_EXISTING
        environment variable, then Skip. An unrecognized value throws rather
        than falling back, so a typo cannot silently package under the wrong
        behavior.

    .OUTPUTS
        [pscustomobject] Behavior ('Skip', 'Overwrite' or 'Fail') and Source.
    #>
    param([AllowEmptyString()][string]$Requested = '')

    $valid = @('Skip', 'Overwrite', 'Fail')

    $value = [string]$Requested
    $source = 'parameter'
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [string]$env:APP_PACKAGER_ON_EXISTING
        $source = 'APP_PACKAGER_ON_EXISTING'
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        return [pscustomobject]@{ Behavior = 'Skip'; Source = 'default' }
    }

    $trimmed = $value.Trim()
    $match = @($valid | Where-Object { $_ -eq $trimmed })
    if ($match.Count -ne 1) {
        throw ("OnExisting value '{0}' from {1} is not one of: {2}." -f $trimmed, $source, ($valid -join ', '))
    }
    return [pscustomobject]@{ Behavior = [string]$match[0]; Source = $source }
}

Export-ModuleMember -Function Resolve-OnExistingBehavior, Get-OnExistingConflictMarker

function Get-DeploymentTypeContentOptions {
    <#
    .SYNOPSIS
        Resolves the content options applied to each ConfigMgr deployment type.
    .DESCRIPTION
        APP_PACKAGER_DT_CONTENT_FALLBACK ('Allow' or 'Deny') maps to
        -ContentFallback. APP_PACKAGER_DT_SLOW_NETWORK_MODE ('Download' or
        'DoNothing') maps to -SlowNetworkDeploymentMode. A missing or
        unrecognized value resolves to Allow and Download.
    #>
    [CmdletBinding()]
    param()

    $fallback = $true
    $fallbackRaw = ([string]$env:APP_PACKAGER_DT_CONTENT_FALLBACK).Trim()
    if ($fallbackRaw -ieq 'Deny') { $fallback = $false }
    elseif ($fallbackRaw -and $fallbackRaw -ine 'Allow') {
        Write-Log ("APP_PACKAGER_DT_CONTENT_FALLBACK '{0}' is not Allow or Deny; using Allow." -f $fallbackRaw) -Level WARN
    }

    $mode = 'Download'
    $modeRaw = ([string]$env:APP_PACKAGER_DT_SLOW_NETWORK_MODE).Trim()
    if ($modeRaw -ieq 'DoNothing') { $mode = 'DoNothing' }
    elseif ($modeRaw -and $modeRaw -ine 'Download') {
        Write-Log ("APP_PACKAGER_DT_SLOW_NETWORK_MODE '{0}' is not Download or DoNothing; using Download." -f $modeRaw) -Level WARN
    }

    return [pscustomobject]@{ ContentFallback = $fallback; SlowNetworkDeploymentMode = $mode }
}

function Get-CheckResultSummaryLine {
    <#
    .SYNOPSIS
        The one log line that closes Check Latest or Check ConfigMgr: how many of the checked products need an update.
    .DESCRIPTION
        A product needs an update when its status starts with 'Update available'. This is the rule of Select-OnlyUpdateAvailable,
        which selects those products after Check ConfigMgr. A product whose check failed has the status 'Error'; it is not
        counted as needing an update and is named in the line. Pass -SelectedCount when the check selects the products; leave
        it out when the check selects nothing.
    #>
    param(
        [Parameter(Mandatory)][string]$Check,
        [AllowEmptyCollection()][AllowNull()][object[]]$Rows = @(),
        [AllowNull()][Nullable[int]]$SelectedCount = $null
    )
    $checked = @($Rows | Where-Object { $_ })
    $eligible = @($checked | Where-Object { [string]$_.Status -like 'Update available*' }).Count
    $failed = @($checked | Where-Object { [string]$_.Status -eq 'Error' }).Count
    $products = $(if ($checked.Count -eq 1) { 'product' } else { 'products' })
    $verb = $(if ($eligible -eq 1 -or $checked.Count -eq 1) { 'needs' } else { 'need' })
    $line = '{0} finished: {1} of {2} {3} {4} an update' -f $Check, $eligible, $checked.Count, $products, $verb
    if ($null -ne $SelectedCount) { $line += '; {0} selected' -f $SelectedCount }
    if ($failed -gt 0) { $line += '; {0} {1} failed' -f $failed, $(if ($failed -eq 1) { 'check' } else { 'checks' }) }
    return ($line + '.')
}
Export-ModuleMember -Function Get-CheckResultSummaryLine

# ---------------------------------------------------------------------------
# Cancel for every run (Q-69). The window sets one run state for the run that is active. The state is kept in the
# application domain, so the background runspace and the window thread see the same object. Every child process that
# a run starts is registered by its process id; Cancel and closing the window stop those processes and no others
# (vault: testing-policy, "Process rules on a lab VM").
# ---------------------------------------------------------------------------
function Set-RunState {
    <#
    .SYNOPSIS
        Makes a state the state of the run that is active, or clears it with $null. The state is a synchronized hashtable
        with CancelRequested, Aborted, Done counts and so on.
    #>
    [CmdletBinding()]
    param([AllowNull()]$State)
    [AppDomain]::CurrentDomain.SetData('AppPackager.RunState', $State)
}

function Get-RunState {
    [CmdletBinding()]
    param()
    return [AppDomain]::CurrentDomain.GetData('AppPackager.RunState')
}

function Test-RunCancelRequested {
    [CmdletBinding()]
    param()
    $state = Get-RunState
    return [bool]($state -and [bool]$state['CancelRequested'])
}

function Test-RunAborted {
    # True when a child process of the active run was stopped because of Cancel.
    [CmdletBinding()]
    param()
    $state = Get-RunState
    return [bool]($state -and [bool]$state['Aborted'])
}

function Set-RunAborted {
    [CmdletBinding()]
    param()
    $state = Get-RunState
    if ($state) { $state['Aborted'] = $true }
}

function Get-RunChildRegistry {
    [CmdletBinding()]
    param()
    $registry = [AppDomain]::CurrentDomain.GetData('AppPackager.RunChildren')
    if (-not $registry) {
        $registry = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[int,long]'
        [AppDomain]::CurrentDomain.SetData('AppPackager.RunChildren', $registry)
    }
    return , $registry
}

function Register-RunChild {
    # Records a process that the run started: its id and its start time (the start time tells a reused id apart).
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Diagnostics.Process]$Process)
    $ticks = 0L
    try { $ticks = $Process.StartTime.Ticks } catch { $ticks = 0L }
    $registry = Get-RunChildRegistry
    $registry[[int]$Process.Id] = $ticks
}

function Unregister-RunChild {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ProcessId)
    $ignored = 0L
    [void](Get-RunChildRegistry).TryRemove($ProcessId, [ref]$ignored)
}

function Stop-ProcessTreeById {
    <#
    .SYNOPSIS
        Stops one process and the processes it started, by process id. Returns nothing.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ProcessId)
    try { & taskkill.exe /PID $ProcessId /T /F 2>&1 | Out-Null } catch { <# Stopping the tree is best-effort; Kill follows. #> $null = $_ }
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($p) { try { $p.Kill() } catch { <# The process may have ended already. #> $null = $_ } }
}

function Stop-RunChildren {
    <#
    .SYNOPSIS
        Stops every process that the active run registered, by process id and only when its start time still matches.
        Returns the number of processes that were stopped.
    #>
    [CmdletBinding()]
    param()
    $stopped = 0
    $registry = Get-RunChildRegistry
    foreach ($entry in @($registry.ToArray())) {
        $p = Get-Process -Id $entry.Key -ErrorAction SilentlyContinue
        $same = $false
        if ($p) { try { $same = ($entry.Value -eq 0L) -or ($p.StartTime.Ticks -eq $entry.Value) } catch { $same = $false } }
        if ($same) { Stop-ProcessTreeById -ProcessId $entry.Key; $stopped++ }
        $gone = 0L
        [void]$registry.TryRemove($entry.Key, [ref]$gone)
    }
    return $stopped
}

function Wait-RunChild {
    <#
    .SYNOPSIS
        Waits for a child process of the run, and ends it within a fraction of a second when the run is cancelled.
    .DESCRIPTION
        Returns Exited, TimedOut or Cancelled. On Cancelled the process and the processes it started are stopped by id and
        the run state says Aborted. A process that was stopped from outside because of Cancel (the Cancel button stops the
        registered children at once) also counts as Cancelled when it ends with a failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Diagnostics.Process]$Process,
        [int]$TimeoutSeconds = 30,
        [int]$PollMilliseconds = 100
    )
    $start = [DateTime]::UtcNow
    while ($true) {
        if (Test-RunCancelRequested) {
            if (-not $Process.HasExited) { Stop-ProcessTreeById -ProcessId $Process.Id; [void]$Process.WaitForExit(3000) }
            Set-RunAborted
            return 'Cancelled'
        }
        if ($Process.WaitForExit($PollMilliseconds)) {
            if ((Test-RunCancelRequested) -and $Process.ExitCode -ne 0) { Set-RunAborted; return 'Cancelled' }
            return 'Exited'
        }
        if ($TimeoutSeconds -gt 0 -and ([DateTime]::UtcNow - $start).TotalSeconds -gt $TimeoutSeconds) { return 'TimedOut' }
    }
}

function Get-CancelledSummaryLine {
    # The line that the status and the log show when a run was cancelled: how many products were done.
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Done, [Parameter(Mandatory)][int]$Total)
    return ('Cancelled: {0} of {1} {2} done.' -f $Done, $Total, $(if ($Total -eq 1) { 'product' } else { 'products' }))
}
Export-ModuleMember -Function Set-RunState, Get-RunState, Test-RunCancelRequested, Test-RunAborted, Set-RunAborted, Register-RunChild, Unregister-RunChild, Stop-ProcessTreeById, Stop-RunChildren, Wait-RunChild, Get-CancelledSummaryLine
Export-ModuleMember -Function Get-DeploymentTypeContentOptions

# ---------------------------------------------------------------------------
# Intune Win32 publishing (Graph)
# ---------------------------------------------------------------------------

function Get-IntuneWinEncryptionInfo {
    <#
    .SYNOPSIS
        Reads the metadata and encryption info from a .intunewin file.

    .DESCRIPTION
        A .intunewin file is a zip whose IntuneWinPackage/Metadata/Detection.xml
        carries the fields the Graph commit action needs (fileEncryptionInfo)
        plus the setup file name and unencrypted size. The encrypted payload
        itself is IntuneWinPackage/Contents/IntunePackage.intunewin.

    .OUTPUTS
        [pscustomobject] Name, FileName, SetupFile, UnencryptedContentSize,
        EncryptionKey, MacKey, InitializationVector, Mac, ProfileIdentifier,
        FileDigest, FileDigestAlgorithm.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "IntuneWin file not found: $Path" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'IntuneWinPackage/Metadata/Detection.xml' } | Select-Object -First 1
        if (-not $entry) { throw "Detection.xml not found inside $Path; not a valid .intunewin package." }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { [xml]$doc = $reader.ReadToEnd() } finally { $reader.Dispose() }

        $info = $doc.ApplicationInfo
        $enc = $info.EncryptionInfo
        if (-not $enc) { throw "Detection.xml carries no EncryptionInfo; the package cannot be committed." }

        return [pscustomobject]@{
            Name                   = [string]$info.Name
            FileName               = [string]$info.FileName
            SetupFile              = [string]$info.SetupFile
            UnencryptedContentSize = [long]$info.UnencryptedContentSize
            EncryptionKey          = [string]$enc.EncryptionKey
            MacKey                 = [string]$enc.MacKey
            InitializationVector   = [string]$enc.InitializationVector
            Mac                    = [string]$enc.Mac
            ProfileIdentifier      = [string]$enc.ProfileIdentifier
            FileDigest             = [string]$enc.FileDigest
            FileDigestAlgorithm    = [string]$enc.FileDigestAlgorithm
        }
    }
    finally { $zip.Dispose() }
}

function Export-IntuneWinPayload {
    <#
    .SYNOPSIS
        Extracts the encrypted payload from a .intunewin file.

    .DESCRIPTION
        The bytes uploaded to Azure Storage are the inner encrypted
        IntunePackage.intunewin, not the outer zip.

    .OUTPUTS
        [pscustomobject] Path and Size of the extracted payload.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'IntuneWinPackage/Contents/IntunePackage.intunewin' } | Select-Object -First 1
        if (-not $entry) { throw "Encrypted payload not found inside $Path; not a valid .intunewin package." }
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $Destination, $true)
        return [pscustomobject]@{ Path = $Destination; Size = (Get-Item -LiteralPath $Destination).Length }
    }
    finally { $zip.Dispose() }
}

function Get-MsGraphToken {
    <#
    .SYNOPSIS
        Acquires an app-only Graph token through the client-credentials flow.
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $resp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
    } -ErrorAction Stop
    if (-not $resp.access_token) { throw 'Token endpoint returned no access_token.' }
    return [string]$resp.access_token
}

function Invoke-GraphJson {
    <#
    .SYNOPSIS
        One Graph REST call. Central so retries and tests hang off one seam.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Token,
        $Body = $null
    )

    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) { $params['Body'] = ($Body | ConvertTo-Json -Depth 12) }
    return Invoke-RestMethod @params
}

function Invoke-AzureBlobUpload {
    <#
    .SYNOPSIS
        Uploads a file to an Azure block-blob SAS URI in chunks.

    .DESCRIPTION
        Standard block-blob protocol: PUT each chunk to
        <uri>&comp=block&blockid=<base64 id>, then PUT the ordered block
        list to <uri>&comp=blocklist. 6 MiB chunks keep each request under
        proxy limits without ballooning the block count.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$FilePath,
        [int]$ChunkSizeMB = 6
    )

    $chunkSize = $ChunkSizeMB * 1MB
    $stream = [System.IO.File]::OpenRead($FilePath)
    $blockIds = New-Object System.Collections.Generic.List[string]
    try {
        $buffer = New-Object byte[] $chunkSize
        $index = 0
        while (($read = $stream.Read($buffer, 0, $chunkSize)) -gt 0) {
            $blockId = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('block-{0:D6}' -f $index)))
            $blockIds.Add($blockId)
            # Only a [byte[]] body goes out raw; an if-expression or a range index
            # yields an Object[] that Invoke-RestMethod serializes as decimal text.
            [byte[]]$chunk = New-Object byte[] $read
            [Array]::Copy($buffer, $chunk, $read)
            $blockUri = '{0}&comp=block&blockid={1}' -f $Uri, [uri]::EscapeDataString($blockId)
            Invoke-RestMethod -Method Put -Uri $blockUri -Body $chunk -Headers @{ 'x-ms-blob-type' = 'BlockBlob' } -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
            $index++
        }
        $blockListXml = '<?xml version="1.0" encoding="utf-8"?><BlockList>' + (($blockIds | ForEach-Object { "<Latest>$_</Latest>" }) -join '') + '</BlockList>'
        Invoke-RestMethod -Method Put -Uri ('{0}&comp=blocklist' -f $Uri) -Body $blockListXml -ContentType 'text/plain; charset=iso-8859-1' -ErrorAction Stop | Out-Null
        Write-Log ("Uploaded payload             : {0} block(s)" -f $blockIds.Count)
    }
    finally { $stream.Dispose() }
}

function New-IntuneRegistryValueScriptRule {
    <#
    .SYNOPSIS
        Builds a Graph PowerShell-script detection rule that applies a prefix,
        suffix, or substring comparison to a registry value.

    .DESCRIPTION
        Graph registry rules offer equality and ordering operators only, so a
        BeginsWith/EndsWith/Contains detection keeps its meaning through a
        script that reads the value from the requested registry view and
        exits 0 with output when the comparison holds.
    #>
    param(
        [Parameter(Mandatory)][string]$HiveRoot,
        [Parameter(Mandatory)][string]$KeyPath,
        [Parameter(Mandatory)][string]$ValueName,
        [Parameter(Mandatory)][string]$Expected,
        [Parameter(Mandatory)][ValidateSet('BeginsWith', 'EndsWith', 'Contains')][string]$Operator,
        [bool]$Is64Bit = $true
    )

    $hive = if ($HiveRoot -like 'HKEY_CURRENT_USER*') { 'CurrentUser' } else { 'LocalMachine' }
    $view = if ($Is64Bit) { 'Registry64' } else { 'Registry32' }
    $quote = { param($s) "'" + (ConvertTo-SingleQuotedContent ([string]$s)) + "'" }
    # String.Contains(string, StringComparison) is absent from .NET Framework,
    # which the client-side Windows PowerShell runs the script under.
    $test = switch ($Operator) {
        'BeginsWith' { '$value.StartsWith({0}, [System.StringComparison]::OrdinalIgnoreCase)' -f (& $quote $Expected) }
        'EndsWith'   { '$value.EndsWith({0}, [System.StringComparison]::OrdinalIgnoreCase)' -f (& $quote $Expected) }
        'Contains'   { '$value.IndexOf({0}, [System.StringComparison]::OrdinalIgnoreCase) -ge 0' -f (& $quote $Expected) }
    }
    $lines = @(
        ('$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::{0}, [Microsoft.Win32.RegistryView]::{1})' -f $hive, $view),
        ('$key = $base.OpenSubKey({0})' -f (& $quote $KeyPath)),
        'if ($key) {',
        ('    $value = [string]$key.GetValue({0})' -f (& $quote $ValueName)),
        ('    if ({0}) {{ Write-Output ''Detected''; exit 0 }}' -f $test),
        '}',
        'exit 1'
    )

    return ($lines -join "`r`n")
}

function Get-IntuneScriptRuleSettings {
    <#
    .SYNOPSIS
        Resolves enforceSignatureCheck and runAs32Bit for a script rule.

    .DESCRIPTION
        Signature enforcement only asks the client to verify; it never signs.
        It is set when the detection category produced a valid
        signature, or when policy requires one - and a required signature that
        was not produced stops the publish instead of shipping an unenforced
        rule. Script host bitness comes from the definition, not a constant.

    .OUTPUTS
        [pscustomobject] EnforceSignatureCheck, RunAs32Bit.
    #>
    param([Parameter(Mandatory)]$Manifest)

    $status = ''
    if ($Manifest.PSObject.Properties['ScriptSigning'] -and $Manifest.ScriptSigning -and
        $Manifest.ScriptSigning.PSObject.Properties['Detection'] -and $Manifest.ScriptSigning.Detection) {
        $status = [string]$Manifest.ScriptSigning.Detection.Status
    }
    $signed = @('SignedAndVerified', 'ExistingSignatureValid') -contains $status
    $require = Get-PolicyFlag -Policy (Resolve-CommonSigningPolicy) -Name 'RequireDetection'
    if ($require -and -not $signed) {
        throw ("Valid detection script signatures are required, but the staged detection reports status '{0}'; re-stage with signing enabled before publishing to Intune." -f $(if ($status) { $status } else { 'none' }))
    }

    $runAs32Bit = $false
    if ($Manifest.PSObject.Properties['Execution'] -and $Manifest.Execution -and
        $Manifest.Execution.PSObject.Properties['ScriptHost']) {
        $runAs32Bit = ([string]$Manifest.Execution.ScriptHost -eq 'x86')
    }

    return [pscustomobject]@{
        EnforceSignatureCheck = ($signed -or $require)
        RunAs32Bit            = $runAs32Bit
    }
}

function New-IntuneDetectionScriptRule {
    <#
    .SYNOPSIS
        Builds a Graph PowerShell script detection rule from script content.

    .DESCRIPTION
        A finalized detection names a file whose bytes the signature covers;
        scriptContent carries those exact bytes so the signature survives the
        transport. Text is only encoded directly when no such file exists.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [AllowEmptyString()][string]$ScriptText = '',
        [AllowEmptyString()][string]$ScriptFile = ''
    )

    $settings = Get-IntuneScriptRuleSettings -Manifest $Manifest
    $content = $null
    if (-not [string]::IsNullOrWhiteSpace($ScriptFile) -and (Get-Command -Name Get-SignedScriptRepresentation -ErrorAction SilentlyContinue)) {
        $representation = Get-SignedScriptRepresentation -Path $ScriptFile
        $content = [string]$representation.Base64
    }
    if ([string]::IsNullOrWhiteSpace($content)) {
        $content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($ScriptText))
    }
    # The client verifies the bytes in scriptContent, not the status the
    # manifest recorded at stage time.
    if ($settings.EnforceSignatureCheck) {
        Assert-ScriptSignatureBeforeSend -Bytes ([Convert]::FromBase64String($content)) -Context 'Intune detection script'
    }

    return @{
        '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptRule'
        ruleType              = 'detection'
        enforceSignatureCheck = $settings.EnforceSignatureCheck
        runAs32Bit            = $settings.RunAs32Bit
        scriptContent         = $content
        operationType         = 'notConfigured'
        operator              = 'notConfigured'
    }
}

function ConvertTo-IntuneWin32Rules {
    <#
    .SYNOPSIS
        Maps a stage manifest Detection onto Graph win32LobApp rules.

    .DESCRIPTION
        Field names per the v1.0 win32LobApp*Rule resources. RegistryKeyValue
        compares DisplayVersion-style strings with a version operation when
        the expected value parses as a version, string equality otherwise.
        Operators map one to one where Graph has an equivalent; any other
        operator is refused. A prefix, suffix, or substring operator has no
        native Graph equivalent: it becomes a generated script rule only when
        the profile set Detection.IntuneScriptConversion, because turning a
        validated native detector into a script silently is the behavior the
        catalog policy forbids. Compound detections map clause-per-rule (Graph
        ANDs all detection rules; OR-connected compounds are refused rather
        than silently narrowed).
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [AllowNull()][pscustomobject]$Profile
    )

    $det = $Manifest.Detection
    $scriptConversionAllowed = $false
    if ($det -and $det.PSObject.Properties['IntuneScriptConversion']) {
        $scriptConversionAllowed = ($det.IntuneScriptConversion -eq $true)
    }
    if (-not $scriptConversionAllowed -and $Profile -and $Profile.PSObject.Properties['Detection'] -and $Profile.Detection) {
        $scriptConversionAllowed = ($Profile.Detection.IntuneScriptConversion -eq $true)
    }
    $type = if ($det.Type) { [string]$det.Type } else { 'RegistryKeyValue' }

    $hiveRoot = { param($d) if ($d.PSObject.Properties['Hive'] -and [string]$d.Hive -match '^(CurrentUser|HKCU)$') { 'HKEY_CURRENT_USER\' } else { 'HKEY_LOCAL_MACHINE\' } }
    $operatorMap = @{
        IsEquals      = 'equal'
        NotEquals     = 'notEqual'
        GreaterThan   = 'greaterThan'
        GreaterEquals = 'greaterThanOrEqual'
        LessThan      = 'lessThan'
        LessEquals    = 'lessThanOrEqual'
    }
    $scriptOperators = @('BeginsWith', 'EndsWith', 'Contains')
    $mapOperator = {
        param($name, $context)
        if ($operatorMap.ContainsKey($name)) { return $operatorMap[$name] }
        throw "Detection operator '$name' ($context) has no Intune rule mapping."
    }
    $mapOne = {
        param($d, $dType)
        switch ($dType) {
            'RegistryKeyValue' {
                $expected = if ($d.ExpectedValue) { [string]$d.ExpectedValue } else { [string]$d.DisplayVersion }
                $valName = if ($d.ValueName) { [string]$d.ValueName } else { 'DisplayVersion' }
                # Missing operators default the way the ConfigMgr clause builder does.
                $opName = if ($d.Operator) { [string]$d.Operator } else { 'IsEquals' }
                if ($scriptOperators -contains $opName) {
                    if (-not $scriptConversionAllowed) {
                        throw ("Detection operator '{0}' on registry value '{1}' has no native Intune rule; converting it to a PowerShell detection script is an explicit profile choice (Detection.IntuneScriptConversion). Set it in the workbench, or publish this application to ConfigMgr only." -f $opName, $valName)
                    }
                    $generated = New-IntuneRegistryValueScriptRule -HiveRoot (& $hiveRoot $d) -KeyPath ([string]$d.RegistryKeyRelative) `
                        -ValueName $valName -Expected $expected -Operator $opName -Is64Bit ([bool]$d.Is64Bit)
                    return (New-IntuneDetectionScriptRule -Manifest $Manifest -ScriptText $generated)
                }
                $ver = $null
                $opType = if ([version]::TryParse($expected, [ref]$ver)) { 'version' } else { 'string' }
                $op = & $mapOperator $opName 'registry value'
                @{
                    '@odata.type'        = '#microsoft.graph.win32LobAppRegistryRule'
                    ruleType             = 'detection'
                    check32BitOn64System = (-not [bool]$d.Is64Bit)
                    keyPath              = ((& $hiveRoot $d) + ([string]$d.RegistryKeyRelative))
                    valueName            = $valName
                    operationType        = $opType
                    operator             = $op
                    comparisonValue      = $expected
                }
            }
            'RegistryKey' {
                @{
                    '@odata.type'        = '#microsoft.graph.win32LobAppRegistryRule'
                    ruleType             = 'detection'
                    check32BitOn64System = (-not [bool]$d.Is64Bit)
                    keyPath              = ((& $hiveRoot $d) + ([string]$d.RegistryKeyRelative))
                    valueName            = $null
                    operationType        = 'exists'
                    operator             = 'notConfigured'
                    comparisonValue      = $null
                }
            }
            'File' {
                $propType = [string]$d.PropertyType
                if ($propType -eq 'Version') {
                    $opName = if ($d.Operator) { [string]$d.Operator } else { 'GreaterEquals' }
                    $op = & $mapOperator $opName 'file version'
                    @{
                        '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
                        ruleType             = 'detection'
                        path                 = [string]$d.FilePath
                        fileOrFolderName     = [string]$d.FileName
                        check32BitOn64System = (-not [bool]$d.Is64Bit)
                        operationType        = 'version'
                        operator             = $op
                        comparisonValue      = [string]$d.ExpectedValue
                    }
                }
                else {
                    @{
                        '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
                        ruleType             = 'detection'
                        path                 = [string]$d.FilePath
                        fileOrFolderName     = [string]$d.FileName
                        check32BitOn64System = (-not [bool]$d.Is64Bit)
                        operationType        = 'exists'
                        operator             = 'notConfigured'
                        comparisonValue      = $null
                    }
                }
            }
            'Script' {
                $scriptFile = ''
                if ($d.PSObject.Properties['ScriptFile'] -and -not [string]::IsNullOrWhiteSpace([string]$d.ScriptFile)) {
                    $scriptFile = [string]$d.ScriptFile
                    if (-not [System.IO.Path]::IsPathRooted($scriptFile) -and $Manifest.PSObject.Properties['ContentRoot']) {
                        $scriptFile = Join-Path ([string]$Manifest.ContentRoot) $scriptFile
                    }
                    # Re-encoding the manifest text would only reproduce the signed bytes for
                    # an ASCII or BOM-less UTF-8 original, so a named-but-missing file is fatal
                    # rather than a silent fallback.
                    if (-not (Test-Path -LiteralPath ("FileSystem::" + $scriptFile) -PathType Leaf)) {
                        throw ("The manifest names detection script file '{0}' but it is missing under the content root; re-run the Stage phase so the finalized script is written before publishing." -f [string]$d.ScriptFile)
                    }
                }
                New-IntuneDetectionScriptRule -Manifest $Manifest -ScriptText ([string]$d.ScriptText) -ScriptFile $scriptFile
            }
            default { throw "Detection type '$dType' has no Intune rule mapping." }
        }
    }

    if ($type -eq 'Compound') {
        if ([string]$det.Connector -eq 'Or' -or ($det.PSObject.Properties['GroupSizes'] -and $det.GroupSizes)) {
            throw 'OR-connected compound detections cannot map to Intune rules (Graph ANDs all detection rules); use a Script detection for this app.'
        }
        $rules = @($det.Clauses | ForEach-Object { & $mapOne $_ ([string]$_.Type) })
        if ($rules.Count -gt 1 -and ($rules | Where-Object { $_['@odata.type'] -eq '#microsoft.graph.win32LobAppPowerShellScriptRule' })) {
            throw 'A prefix, suffix, or substring clause maps to a script rule, which this mapping does not combine with other clauses; use a Script detection for this app.'
        }
        return $rules
    }
    return @(& $mapOne $det $type)
}

function Test-IntuneManifestHasRules {
    <#
    .SYNOPSIS
        True when the manifest carries finished Graph rules.
    .DESCRIPTION
        A manifest converted from a ConfigMgr application holds IntuneRules:
        the detection and requirement rules already in Graph form, so the
        publisher sends them as they are.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Manifest)
    return ($Manifest.PSObject.Properties['IntuneRules'] -and @($Manifest.IntuneRules).Count -gt 0)
}

function ConvertFrom-RequirementsJson {
    <#
    .SYNOPSIS
        The requirement specs in the APP_PACKAGER_REQUIREMENTS JSON of a run.

    .DESCRIPTION
        The JSON is { SchemaVersion, Rules: [...] }, which the main window sets
        from the saved deployment conditions of an application. Malformed JSON
        throws instead of publishing without the rules the operator configured.
    #>
    param([AllowEmptyString()][string]$Json = '')

    if ([string]::IsNullOrWhiteSpace($Json)) { return @() }
    try { $parsed = $Json | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "APP_PACKAGER_REQUIREMENTS is not valid JSON: $($_.Exception.Message)" }
    if ($parsed -and $parsed.PSObject.Properties['Rules'] -and $parsed.Rules) { return @($parsed.Rules) }
    return @()
}

function Get-IntuneRequirementDropKey {
    <#
    .SYNOPSIS
        The key that names one requirement rule of one application in Intune.AcceptedRequirementDrops.

    .DESCRIPTION
        The identity tag of the application in Intune (the tag that finds the
        app again on the next publish), a bar and the condition id. It does not
        change with the version.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Manifest, [AllowEmptyString()][string]$ConditionId = '')

    return ('{0}|{1}' -f (Get-IntuneIdentityTag -Manifest $Manifest), $ConditionId)
}

function Get-IntuneRequirementSpecs {
    <#
    .SYNOPSIS
        The requirement specs that the Intune publish of a manifest applies.

    .DESCRIPTION
        The same two sources as the ConfigMgr publish
        (Get-DeploymentTypeRequirementSpecs): the Requirements array of the
        manifest and the specs the run applies beside it (the saved deployment
        conditions). A spec that repeats an earlier one is applied once.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Manifest, [AllowNull()][object[]]$Extra = @())

    $specs = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $sources = @()
    if ($Manifest.PSObject.Properties['Requirements'] -and $Manifest.Requirements) { $sources += @($Manifest.Requirements) }
    if ($Extra) { $sources += @($Extra) }
    foreach ($spec in $sources) {
        if ($null -eq $spec) { continue }
        $value = $(if ($spec -is [System.Collections.IDictionary]) { $spec['Value'] } else { $spec.Value })
        $cultures = $(if ($spec -is [System.Collections.IDictionary]) { $spec['Cultures'] } else { $spec.Cultures })
        $identity = '{0}|{1}|{2}' -f [string]$(if ($spec -is [System.Collections.IDictionary]) { $spec['ConditionId'] } else { $spec.ConditionId }), [string]$value, (@($cultures) -join ',')
        if ($seen.Add($identity)) { $specs.Add($spec) }
    }
    return @($specs.ToArray())
}

function Get-IntuneRequirementScriptPlan {
    <#
    .SYNOPSIS
        Binds one requirement spec to its condition template and writes the check script.

    .DESCRIPTION
        Graph has no rule for a WMI value, an operating system language or a
        custom condition, so each becomes a PowerShell requirement rule that
        prints True or False and passes on the expected output. The binding is
        the one of the ConfigMgr publish: an unknown value or a missing value
        fails here too. A spec that cannot become such a script throws; the
        caller records the message as the gap.

    .OUTPUTS
        [pscustomobject] Template, Script, Expected, IsScriptCondition.
    #>
    param([Parameter(Mandatory)]$Spec, [Parameter(Mandatory)]$Templates)

    $read = { param($Object, $Name) if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { $Object[$Name] } else { $null } } elseif ($Object.PSObject.Properties[$Name]) { $Object.$Name } else { $null } }
    $quote = { param($Text) "'" + ([string]$Text -replace "'", "''") + "'" }

    $id = [string](& $read $Spec 'ConditionId')
    $template = @($Templates.Conditions | Where-Object { [string]$_.Id -eq $id } | Select-Object -First 1)[0]
    if (-not $template) {
        throw ("The requirement names the condition '{0}', which is not in the condition templates." -f $id)
    }
    if ([string]::IsNullOrWhiteSpace([string]$template.GlobalConditionName)) {
        throw ("Condition template '{0}' has no GlobalConditionName." -f $template.Id)
    }

    $ruleType = [string]$template.RuleType
    $kind = [string]$template.Kind
    $rawValue = & $read $Spec 'Value'
    $value = $null
    $cultures = @()
    $boolValue = $false
    switch ($ruleType) {
        'CommonValue' {
            $value = $(if ($null -eq $rawValue) { $null } else { [string]$rawValue })
            $mapping = @()
            if ($template.PSObject.Properties['Values'] -and $template.Values) { $mapping = @($template.Values.PSObject.Properties) }
            if ($mapping.Count -gt 0) {
                $mapped = $template.Values.PSObject.Properties[[string]$value]
                if (-not $mapped) {
                    throw ("Condition '{0}' has no mapping for value '{1}' (known: {2})." -f $id, $value, (($mapping | ForEach-Object { $_.Name }) -join ', '))
                }
                $value = [string]$mapped.Value
            }
            if ([string]::IsNullOrWhiteSpace($value)) { throw ("Condition '{0}' requires a Value." -f $id) }
        }
        'OperatingSystemLanguage' {
            foreach ($culture in @(& $read $Spec 'Cultures')) {
                if ([string]::IsNullOrWhiteSpace([string]$culture)) { continue }
                # A name that is not a language tag is a gap on every host; a newer .NET accepts some unknown names.
                if ([string]$culture -cnotmatch '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$') { throw ("The culture code '{0}' is not valid (for example de-DE)." -f $culture) }
                $cultures += [System.Globalization.CultureInfo]::GetCultureInfo([string]$culture).Name
            }
            if ($cultures.Count -eq 0) { throw ("Condition '{0}' requires at least one culture code (for example de-DE)." -f $id) }
        }
        'Boolean' {
            if ($rawValue -is [bool]) { $boolValue = $rawValue }
            else {
                $text = $(if ($null -eq $rawValue) { '' } else { [string]$rawValue })
                $parsed = $false
                if (-not [bool]::TryParse($text, [ref]$parsed)) { throw ("Condition '{0}' requires a boolean Value (got '{1}')." -f $id, $text) }
                $boolValue = $parsed
            }
        }
        default { throw ("Condition '{0}' has unsupported RuleType '{1}'." -f $id, $ruleType) }
    }

    $checkScript = $null
    $expected = $true
    if ($ruleType -ceq 'CommonValue' -and $kind -ceq 'Wql') {
        if ([string]::IsNullOrWhiteSpace([string]$template.Namespace) -or [string]::IsNullOrWhiteSpace([string]$template.Class) -or [string]::IsNullOrWhiteSpace([string]$template.Property)) {
            throw ("The condition '{0}' names no WMI namespace, class and property." -f $template.Id)
        }
        # A WMI property name goes into the script as code, so it must be a plain name.
        if ([string]$template.Property -cnotmatch '^[A-Za-z0-9_]+$') { throw ("The WMI property name '{0}' is not a plain name." -f $template.Property) }
        # True when some row of the WMI class has the value; the ConfigMgr rule compares with IsEquals.
        $checkScript = @(
            '$ErrorActionPreference = ''SilentlyContinue'''
            ('$values = @(Get-CimInstance -Namespace {0} -ClassName {1} | ForEach-Object {{ [string]$_.{2} }})' -f (& $quote $template.Namespace), (& $quote $template.Class), $template.Property)
            ('Write-Output ($values -contains {0})' -f (& $quote $value))
        ) -join "`r`n"
    }
    elseif ($ruleType -ceq 'OperatingSystemLanguage' -and $kind -ceq 'BuiltIn') {
        # True when the installed UI language of the operating system is one of the cultures.
        $checkScript = @(
            ('$cultures = @({0})' -f ((@($cultures) | ForEach-Object { & $quote $_ }) -join ', '))
            '$name = [System.Globalization.CultureInfo]::InstalledUICulture.Name'
            'Write-Output ($cultures -contains $name)'
        ) -join "`r`n"
    }
    elseif ($ruleType -ceq 'Boolean' -and $kind -ceq 'Script') {
        $checkScript = Get-ConditionTemplateScriptText -Template $template
        $expected = $boolValue
    }
    else {
        throw ("The condition '{0}' ({1} condition with the rule type {2}) has no Intune requirement rule." -f $template.Id, $kind, $ruleType)
    }

    return [pscustomobject]@{
        Template          = $template
        Script            = $checkScript
        Expected          = $expected
        IsScriptCondition = ($kind -ceq 'Script')
    }
}

function ConvertTo-IntuneRequirementRules {
    <#
    .SYNOPSIS
        Turns ConfigMgr requirement specs into Intune requirement rules.

    .DESCRIPTION
        Each spec becomes a win32LobAppPowerShellScriptRule of rule type
        requirement, or a gap: the message that says why no Intune rule exists.
        A gap never throws. With -SignScripts, a script condition follows the
        signing policy of the requirement category like the ConfigMgr global
        condition does; a required signature that is not produced throws,
        because no accepted drop covers it. The scripts that this module writes
        itself (WMI value, language) are never signed.

    .OUTPUTS
        [pscustomobject[]] Spec, Rule (hashtable or $null), Gap, Translated.
    #>
    param(
        [AllowNull()][object[]]$Specs = @(),
        [Parameter(Mandatory)]$Templates,
        [switch]$SignScripts
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($spec in @($Specs)) {
        if ($null -eq $spec) { continue }
        $plan = $null
        $gap = $null
        try { $plan = Get-IntuneRequirementScriptPlan -Spec $spec -Templates $Templates }
        catch { $gap = [string]$_.Exception.Message }
        if ($null -eq $plan) {
            $result.Add([pscustomobject]@{ Spec = $spec; Rule = $null; Gap = $gap; Translated = $false })
            continue
        }

        $template = $plan.Template
        $content = [System.Text.Encoding]::UTF8.GetBytes([string]$plan.Script)
        $enforce = $false
        if ($SignScripts -and $plan.IsScriptCondition) {
            $policy = Resolve-CommonSigningPolicy
            $require = [bool](Get-PolicyFlag -Policy $policy -Name 'RequireRequirements')
            if ($require -or (Get-PolicyFlag -Policy $policy -Name 'SignRequirements')) {
                $identity = Get-ConditionScriptContentIdentity -Template $template
                $content = [byte[]]$identity.Bytes
                $enforce = ($require -and [bool]$identity.Signed)
            }
        }
        $name = $(if ([string]::IsNullOrWhiteSpace([string]$template.GlobalConditionName)) { [string]$template.Id } else { [string]$template.GlobalConditionName })
        if ($name.Length -gt 100) { $name = $name.Substring(0, 100) }
        $rule = @{
            '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptRule'
            ruleType              = 'requirement'
            displayName           = $name
            enforceSignatureCheck = $enforce
            runAs32Bit            = $false
            runAsAccount          = 'system'
            scriptContent         = [Convert]::ToBase64String($content)
            operationType         = 'string'
            operator              = 'equal'
            comparisonValue       = $(if ($plan.Expected) { 'True' } else { 'False' })
        }
        $result.Add([pscustomobject]@{ Spec = $spec; Rule = $rule; Gap = $null; Translated = $true })
    }
    return @($result.ToArray())
}

function Get-IntuneCompatibilityFindings {
    <#
    .SYNOPSIS
        Reports what an Intune publish of this manifest would and would not do.

    .DESCRIPTION
        The review pane and the publisher both read this list, so a blocking
        gap is visible before any Graph mutation instead of surfacing as a
        half-created app. Severity Blocking refuses the target; Review needs an
        operator decision; Info records a capability the adapter does not carry.

        Each ConfigMgr requirement of the application is translated to an Intune
        requirement rule (ConvertTo-IntuneRequirementRules). A requirement with
        no Intune rule is a Blocking finding, RequirementNotTranslatable, unless
        its drop is in AcceptedDrops (the Intune.AcceptedRequirementDrops
        setting); then it is a Review finding, RequirementDropped. Building the
        list never signs a script.

    .PARAMETER Requirements
        Requirement specs that the run applies beside the Requirements array of
        the manifest (the saved deployment conditions).

    .PARAMETER AcceptedDrops
        Keys (Get-IntuneRequirementDropKey) of the rules that may be left out of the app.

    .PARAMETER Templates
        The condition templates that the specs bind to; the saved or built-in ones when omitted.

    .OUTPUTS
        [pscustomobject[]] Severity, Code, Message, DropKey, CanAcceptDrop.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [AllowNull()][object[]]$Requirements = @(),
        [AllowNull()][string[]]$AcceptedDrops = @(),
        [AllowNull()]$Templates = $null
    )

    $findings = New-Object System.Collections.Generic.List[object]
    $add = {
        param($Severity, $Code, $Message, $DropKey)
        $findings.Add([pscustomobject]@{
                Severity      = $Severity
                Code          = $Code
                Message       = $Message
                DropKey       = $(if ($DropKey) { [string]$DropKey } else { $null })
                CanAcceptDrop = [bool]($DropKey -and $Severity -eq 'Blocking')
            })
    }

    # Multi-DT means the same thing here as in Get-ManifestDeploymentTypeSpecs: more than one
    # entry, or a single entry whose name differs from the application. A one-entry manifest
    # named after the app is a plain single deployment type and publishes normally.
    $isMultiDt = $false
    if ($Manifest.PSObject.Properties['DeploymentTypes'] -and @($Manifest.DeploymentTypes).Count -gt 0) {
        $specAppName = [string]$Manifest.AppName
        try {
            $specs = @(Get-ManifestDeploymentTypeSpecs -Manifest $Manifest -NetworkContentPath '.' -AppName $specAppName)
            $isMultiDt = ($specs.Count -gt 1 -or $specs[0].DtName -ne $specAppName)
        }
        catch {
            # A findings pass never throws: an unresolvable variant set is reported as blocking.
            $isMultiDt = $true
        }
    }
    if ($isMultiDt) {
        & $add 'Blocking' 'MultipleDeploymentTypes' 'This application has deployment-type variants. Intune has no equivalent priority and fallback model, so the variant set cannot be published as one Win32 app.'
    }
    if ($Manifest.RequireUserInteraction -eq $true) {
        & $add 'Blocking' 'InteractiveInstall' 'This deployment type allows the user to interact with the installation. Intune runs Win32 installs silently, so an interactive profile cannot be published.'
    }

    # A manifest built from an existing ConfigMgr application carries finished
    # Graph rules (IntuneRules); there is no detection block to map.
    if (Test-IntuneManifestHasRules -Manifest $Manifest) {
        $finishedRules = @($Manifest.IntuneRules | Where-Object { $_ })
        if (@($finishedRules | Where-Object { [string]$_.ruleType -eq 'detection' }).Count -eq 0) {
            & $add 'Blocking' 'DetectionMissing' 'The finished Intune rules hold no detection rule, so Intune could not tell whether the app is installed.'
        }
        # The finished script rules are not signed. A policy that requires signed
        # detection or requirement scripts refuses them before any Graph call.
        $signingPolicy = Resolve-CommonSigningPolicy
        foreach ($finishedRule in $finishedRules) {
            if ([string]$finishedRule.'@odata.type' -ne '#microsoft.graph.win32LobAppPowerShellScriptRule') { continue }
            $requiredFlag = $(if ([string]$finishedRule.ruleType -eq 'detection') { 'RequireDetection' } else { 'RequireRequirements' })
            if (Get-PolicyFlag -Policy $signingPolicy -Name $requiredFlag) {
                & $add 'Blocking' 'ConvertedScriptNotSigned' ("The signing policy requires signed {0} scripts. The converted {0} script is not signed." -f [string]$finishedRule.ruleType)
            }
        }
    }
    else {
        try {
            [void](ConvertTo-IntuneWin32Rules -Manifest $Manifest)
        }
        catch {
            & $add 'Blocking' 'DetectionNotMappable' $_.Exception.Message
        }
    }

    # A manifest with finished Graph rules carries its requirements in them.
    if (-not (Test-IntuneManifestHasRules -Manifest $Manifest)) {
        $requirementSpecs = @(Get-IntuneRequirementSpecs -Manifest $Manifest -Extra $Requirements)
        if ($requirementSpecs.Count -gt 0) {
            if ($null -eq $Templates) { $Templates = Get-ConditionTemplates }
            foreach ($translation in @(ConvertTo-IntuneRequirementRules -Specs $requirementSpecs -Templates $Templates)) {
                $conditionId = [string]$(if ($translation.Spec -is [System.Collections.IDictionary]) { $translation.Spec['ConditionId'] } else { $translation.Spec.ConditionId })
                $dropKey = Get-IntuneRequirementDropKey -Manifest $Manifest -ConditionId $conditionId
                if ($translation.Translated) {
                    & $add 'Info' 'RequirementTranslated' ("The requirement '{0}' is published as a PowerShell requirement rule." -f $conditionId)
                }
                elseif (@($AcceptedDrops) -contains $dropKey) {
                    & $add 'Review' 'RequirementDropped' ("The requirement '{0}' has no Intune rule ({1}) and its drop is accepted for this application, so the published app is not gated by it." -f $conditionId, $translation.Gap)
                }
                else {
                    & $add 'Blocking' 'RequirementNotTranslatable' ("The requirement '{0}' has no Intune rule: {1} The app would install on devices that ConfigMgr skips. Remove the requirement, or accept its drop by adding this entry to the setting Intune.AcceptedRequirementDrops: {2}" -f $conditionId, $translation.Gap, $dropKey) $dropKey
                }
            }
        }
    }

    $architecture = [string]$Manifest.Architecture
    if ($Manifest.PSObject.Properties['IntuneAllowedArchitectures'] -and -not [string]::IsNullOrWhiteSpace([string]$Manifest.IntuneAllowedArchitectures)) {
        $unknownArchitectures = @(([string]$Manifest.IntuneAllowedArchitectures -split ',') | Where-Object { @('x86', 'x64', 'arm64') -notcontains $_.Trim() })
        if ($unknownArchitectures.Count -gt 0) {
            & $add 'Blocking' 'ArchitectureUnsupported' ("Architecture '{0}' has no allowedArchitectures value in the Graph v1.0 Win32 app resource (x86, x64, arm64)." -f ($unknownArchitectures -join ', '))
        }
    }
    elseif ([string]::IsNullOrWhiteSpace($architecture)) {
        & $add 'Review' 'ArchitectureUnresolved' 'The manifest declares no architecture, so the published app keeps the publisher default rather than a value resolved from the installer.'
    }
    elseif (@('x86', 'x64', 'arm64') -notcontains (Get-IntuneAllowedArchitecture -Architecture $architecture)) {
        & $add 'Blocking' 'ArchitectureUnsupported' ("Architecture '{0}' has no allowedArchitectures value in the Graph v1.0 Win32 app resource (x86, x64, arm64)." -f $architecture)
    }

    $timing = $Manifest.Timing
    if ($timing) {
        & $add 'Info' 'TimingNotApplied' 'Estimated and maximum runtime minutes are recorded in the build but not sent to Intune: the Graph v1.0 install-experience resource has no runtime property (Timing.IntuneApplied = false).'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Manifest.SetupFile)) {
        & $add 'Info' 'SetupFileFromPackage' 'The manifest names no setup file, so the setup reference comes from the .intunewin package metadata.'
    }

    return @($findings.ToArray())
}

function Get-IntuneAllowedArchitecture {
    <#
    .SYNOPSIS
        Maps a manifest architecture onto a Graph windowsArchitecture value.

    .DESCRIPTION
        The v1.0 win32LobApp resource documents null, x86, x64 and arm64 for
        allowedArchitectures. Setting allowedArchitectures sets
        applicableArchitectures to none, so only one of the two is sent.
    #>
    param([AllowEmptyString()][string]$Architecture)

    switch -Regex ([string]$Architecture) {
        '^(?i)(x64|amd64|64|64-bit|x86_64)$' { return 'x64' }
        '^(?i)(x86|intel|32|32-bit)$'        { return 'x86' }
        '^(?i)(arm64|aarch64)$'              { return 'arm64' }
        default                              { return '' }
    }
}

function Get-IntuneIdentityTag {
    <#
    .SYNOPSIS
        The notes value that identifies an Intune app as this build's target.

    .DESCRIPTION
        'AppPackager:<ApplicationId>/<ProfileId>'. A manifest staged before the
        workbench carries neither id; it falls back to the packager-default
        profile of an application key derived from the title, which keeps a
        legacy app matchable without claiming an identity it never had.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Manifest)

    $applicationId = [string]$Manifest.ApplicationId
    if ([string]::IsNullOrWhiteSpace($applicationId)) {
        $applicationId = 'legacy:' + (([string]$Manifest.AppName) -replace '[^\w.-]', '-')
    }
    $profileId = [string]$Manifest.ProfileId
    if ([string]::IsNullOrWhiteSpace($profileId)) { $profileId = 'default' }
    return ('AppPackager:{0}/{1}' -f $applicationId, $profileId)
}

function Get-IntuneNotesIdentityTags {
    # The identity tags in an app's notes: each line that is an AppPackager
    # tag and nothing else.
    param([AllowEmptyString()][string]$Notes)
    if ([string]::IsNullOrWhiteSpace($Notes)) { return @() }
    # An application id can carry a space (a custom packager's file name).
    return @($Notes -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^AppPackager:.+/.+$' })
}

function Publish-IntuneWin32App {
    <#
    .SYNOPSIS
        Publishes a packaged .intunewin as an Intune Win32 app through Graph.

    .DESCRIPTION
        The complete v1.0 flow: create the win32LobApp (metadata, commands,
        detection rules from the stage manifest), create a content version
        and file entry, wait for the Azure Storage URI, upload the encrypted
        payload in blocks, commit with the package's fileEncryptionInfo,
        wait for commit, and patch committedContentVersion. A repeat publish
        sends the new metadata in that final patch, after the content is
        committed. Assignment is left to the operator in the Intune console.

        Requires an Entra app registration with application permission
        DeviceManagementApps.ReadWrite.All (admin-consented).

        IconPath sets largeIcon on the app entity as a mimeContent. Left empty
        it falls back to the manifest's Icon name resolved beside the
        .intunewin; an unresolvable icon leaves the property unset.

    .OUTPUTS
        [string] The created mobile app id.
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$IntuneWinPath,
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [AllowEmptyString()][string]$Description = '',
        [string]$MinimumSupportedWindowsRelease = 'Windows10_21H2',
        [ValidateSet('x86', 'x64', 'arm64')][string]$Architecture = 'x64',
        [string]$GraphBase = 'https://graph.microsoft.com/v1.0',
        [int]$PollTimeoutSec = 600,
        [AllowEmptyString()][string]$IconPath = '',
        # Requirement specs the run applies beside the manifest's own (the saved deployment conditions).
        [AllowNull()][object[]]$Requirements = @(),
        # Intune.AcceptedRequirementDrops: the requirements that may be left out of the app.
        [AllowNull()][string[]]$AcceptedRequirementDrops = @(),
        [AllowNull()]$ConditionTemplates = $null
    )

    $meta = Get-IntuneWinEncryptionInfo -Path $IntuneWinPath

    # Compatibility is decided before the first Graph mutation: a blocking gap
    # must not leave a half-created app behind in the tenant.
    $findings = @(Get-IntuneCompatibilityFindings -Manifest $Manifest -Requirements $Requirements -AcceptedDrops $AcceptedRequirementDrops -Templates $ConditionTemplates)
    foreach ($finding in $findings) {
        $level = if ($finding.Severity -eq 'Info') { 'INFO' } else { 'WARN' }
        Write-Log ("Intune compatibility         : [{0}] {1} - {2}" -f $finding.Severity, $finding.Code, $finding.Message) -Level $level
    }
    $blocking = @($findings | Where-Object { $_.Severity -eq 'Blocking' })
    if ($blocking.Count -gt 0) {
        throw ("This application cannot be published to Intune: {0}" -f (($blocking | ForEach-Object { "$($_.Code): $($_.Message)" }) -join ' | '))
    }

    $rules = @()
    if (Test-IntuneManifestHasRules -Manifest $Manifest) { $rules = @($Manifest.IntuneRules) }
    else {
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest $Manifest)
        # The requirement rules follow the detection rules. A rule with no Intune form is left out
        # only when its drop was accepted; the findings above already refused every other one, and
        # this check holds for a caller that passes its own findings.
        $requirementSpecs = @(Get-IntuneRequirementSpecs -Manifest $Manifest -Extra $Requirements)
        if ($requirementSpecs.Count -gt 0) {
            $templates = $(if ($null -ne $ConditionTemplates) { $ConditionTemplates } else { Get-ConditionTemplates })
            foreach ($translation in @(ConvertTo-IntuneRequirementRules -Specs $requirementSpecs -Templates $templates -SignScripts)) {
                $conditionId = [string]$(if ($translation.Spec -is [System.Collections.IDictionary]) { $translation.Spec['ConditionId'] } else { $translation.Spec.ConditionId })
                if ($translation.Translated) { $rules += $translation.Rule }
                elseif (@($AcceptedRequirementDrops) -notcontains (Get-IntuneRequirementDropKey -Manifest $Manifest -ConditionId $conditionId)) {
                    throw ("The requirement '{0}' has no Intune rule: {1}" -f $conditionId, $translation.Gap)
                }
            }
        }
    }

    $installCommand = if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.InstallCommandLine)) { [string]$Manifest.InstallCommandLine } else { 'install.bat' }
    $uninstallCommand = if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.UninstallCommandLine)) { [string]$Manifest.UninstallCommandLine } else { 'uninstall.bat' }
    # The real entry the package runs; the prep tool's own setup reference is
    # only correct for the generic install.bat layout.
    $setupFilePath = if (-not [string]::IsNullOrWhiteSpace([string]$Manifest.SetupFile)) { [string]$Manifest.SetupFile } else { [string]$meta.SetupFile }

    # Display names are not identity: two profiles of one application can carry
    # the same title, and a renamed profile must not adopt a different app.
    $identityTag = Get-IntuneIdentityTag -Manifest $Manifest

    # The definition's architecture wins over the caller's default; an
    # unmappable value was already refused by the compatibility findings.
    $resolvedArchitecture = Get-IntuneAllowedArchitecture -Architecture ([string]$Manifest.Architecture)
    if ([string]::IsNullOrWhiteSpace($resolvedArchitecture)) { $resolvedArchitecture = $Architecture }
    # A converted ConfigMgr application names the architecture set itself.
    if ($Manifest.PSObject.Properties['IntuneAllowedArchitectures'] -and -not [string]::IsNullOrWhiteSpace([string]$Manifest.IntuneAllowedArchitectures)) {
        $resolvedArchitecture = (([string]$Manifest.IntuneAllowedArchitectures -split ',') | ForEach-Object { $_.Trim() }) -join ','
    }

    $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

    # Repeat publishes update the existing app (new content version on the
    # same app identity) instead of stacking duplicates in the console.
    # Identity first: an app renamed in the tenant keeps its tag but no longer matches a
    # display-name filter, and `notes` is not a filterable property, so the Win32 apps are
    # listed and matched here.
    $tagged = @()
    $uri = "$GraphBase/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')"
    while ($uri) {
        $page = Invoke-GraphJson -Method GET -Uri $uri -Token $token
        $tagged += @($page.value | Where-Object {
            $_.'@odata.type' -eq '#microsoft.graph.win32LobApp' -and (Get-IntuneNotesIdentityTags -Notes ([string]$_.notes)) -contains $identityTag
        })
        $uri = [string]$page.'@odata.nextLink'
    }
    if ($tagged.Count -gt 1) {
        throw ("Multiple Intune apps carry the identity tag '{0}' (ids: {1}); resolve the duplicates before publishing." -f $identityTag, (($tagged | ForEach-Object { [string]$_.id }) -join ', '))
    }

    $existingApp = $null
    if ($tagged.Count -eq 1) {
        $existingApp = $tagged[0]
    }
    else {
        $filterName = ([string]$Manifest.AppName) -replace "'", "''"
        $existing = Invoke-GraphJson -Method GET -Uri ("$GraphBase/deviceAppManagement/mobileApps?`$filter=displayName eq '$filterName'") -Token $token
        # An app that carries the tag of another application or profile is
        # that build's target; a shared display name does not make it this one.
        $candidates = @($existing.value | Where-Object {
            # A legacy tag is derived from the title, not from an application
            # identity, so an app that carries only that tag can be adopted.
            $_.'@odata.type' -eq '#microsoft.graph.win32LobApp' -and
            @(Get-IntuneNotesIdentityTags -Notes ([string]$_.notes) | Where-Object { $_ -notlike 'AppPackager:legacy:*' }).Count -eq 0
        })
        if ($candidates.Count -gt 1) {
            throw ("Multiple Intune apps are named '{0}' and none carries this application's identity tag (ids: {1}); rename or tag the intended app before publishing." -f $Manifest.AppName, (($candidates | ForEach-Object { [string]$_.id }) -join ', '))
        }
        $existingApp = @($candidates) | Select-Object -First 1
    }

    $appBody = @{
        '@odata.type'                  = '#microsoft.graph.win32LobApp'
        displayName                    = [string]$Manifest.AppName
        description                    = $(if ($Description) { $Description } else { [string]$Manifest.AppName })
        publisher                      = [string]$Manifest.Publisher
        fileName                       = (Split-Path -Leaf $IntuneWinPath)
        setupFilePath                  = $setupFilePath
        installCommandLine             = $installCommand
        uninstallCommandLine           = $uninstallCommand
        notes                          = $identityTag
        # Graph sets applicableArchitectures to none when allowedArchitectures
        # carries a value, so only the newer property is sent.
        allowedArchitectures           = $resolvedArchitecture
        minimumSupportedWindowsRelease = $MinimumSupportedWindowsRelease
        installExperience              = @{
            # A manifest that installs for the user (per-user installer, HKCU
            # registration) runs in the user context in Intune as well.
            runAsAccount          = $(if ([string]$Manifest.InstallationBehaviorType -eq 'InstallForUser') { 'user' } else { 'system' })
            deviceRestartBehavior = 'basedOnReturnCode'
        }
        rules                          = $rules
        returnCodes                    = @(
            @{ returnCode = 0;    type = 'success' }
            @{ returnCode = 1707; type = 'success' }
            @{ returnCode = 3010; type = 'softReboot' }
            @{ returnCode = 1641; type = 'hardReboot' }
            @{ returnCode = 1618; type = 'retry' }
        )
    }
    # A converted ConfigMgr application keeps its own return codes, restart
    # behavior and minimum hardware values (Graph v1.0 properties).
    if ($Manifest.PSObject.Properties['IntuneReturnCodes'] -and @($Manifest.IntuneReturnCodes).Count -gt 0) {
        $appBody['returnCodes'] = @($Manifest.IntuneReturnCodes | ForEach-Object { @{ returnCode = [int]$_.returnCode; type = [string]$_.type } })
    }
    if ($Manifest.PSObject.Properties['IntuneRestartBehavior'] -and -not [string]::IsNullOrWhiteSpace([string]$Manifest.IntuneRestartBehavior)) {
        $appBody['installExperience']['deviceRestartBehavior'] = [string]$Manifest.IntuneRestartBehavior
    }
    if ($Manifest.PSObject.Properties['IntuneMinimums'] -and $Manifest.IntuneMinimums) {
        $minimumPairs = @()
        if ($Manifest.IntuneMinimums -is [System.Collections.IDictionary]) { $minimumPairs = @($Manifest.IntuneMinimums.GetEnumerator() | ForEach-Object { [pscustomobject]@{ Name = [string]$_.Key; Value = $_.Value } }) }
        else { $minimumPairs = @($Manifest.IntuneMinimums.PSObject.Properties | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.Value } }) }
        foreach ($minimum in $minimumPairs) {
            if (@('MinimumMemoryInMB', 'MinimumNumberOfProcessors', 'MinimumCpuSpeedInMHz', 'MinimumFreeDiskSpaceInMB') -contains $minimum.Name) {
                $appBody[$minimum.Name.Substring(0, 1).ToLowerInvariant() + $minimum.Name.Substring(1)] = [int]$minimum.Value
            }
        }
    }

    # "Any Windows architecture" has a documented value only on
    # applicableArchitectures (a combination of x86 and x64); allowedArchitectures
    # takes one architecture. A converted application that ran on both sends the
    # documented property.
    if ($Manifest.PSObject.Properties['IntuneAllowedArchitectures'] -and -not [string]::IsNullOrWhiteSpace([string]$Manifest.IntuneAllowedArchitectures)) {
        $architectureSet = @($resolvedArchitecture -split ',' | ForEach-Object { $_.Trim() } | Sort-Object -Unique)
        if ($architectureSet.Count -eq 2 -and $architectureSet -contains 'x86' -and $architectureSet -contains 'x64') {
            $appBody.Remove('allowedArchitectures')
            $appBody['applicableArchitectures'] = 'x86,x64'
        }
    }

    # largeIcon is a mimeContent (type + base64 bytes) on the app entity, not a
    # separate upload. An unresolvable icon leaves the property unset rather
    # than failing the publish.
    $resolvedIconPath = $IconPath
    # A custom icon chosen in the workbench outranks the packager's extracted
    # or icon-pack image, which a later pack refresh would otherwise restore.
    if ([string]::IsNullOrWhiteSpace($resolvedIconPath) -and
        ($Manifest.PSObject.Properties.Name -contains 'CustomIcon') -and
        -not [string]::IsNullOrWhiteSpace([string]$Manifest.CustomIcon)) {
        $resolvedIconPath = Join-Path (Split-Path -Path $IntuneWinPath -Parent) ([string]$Manifest.CustomIcon)
    }
    if ([string]::IsNullOrWhiteSpace($resolvedIconPath) -and
        ($Manifest.PSObject.Properties.Name -contains 'Icon') -and
        -not [string]::IsNullOrWhiteSpace([string]$Manifest.Icon)) {
        $resolvedIconPath = Join-Path (Split-Path -Path $IntuneWinPath -Parent) ([string]$Manifest.Icon)
    }
    if (-not [string]::IsNullOrWhiteSpace($resolvedIconPath)) {
        $largeIcon = Get-IconMimeContent -Path $resolvedIconPath
        if ($largeIcon) {
            $appBody['largeIcon'] = $largeIcon
            Write-Log ("Intune app icon              : {0} ({1})" -f (Split-Path -Leaf $resolvedIconPath), $largeIcon.type)
        }
        else {
            Write-Log ("Intune app icon not applied  : {0}" -f $resolvedIconPath) -Level WARN
        }
    }

    if ($existingApp) {
        # Notes can hold operator text. The tag is added on its own first
        # line when the app is adopted, and the notes stay as they are once
        # they carry the tag.
        $existingNotes = [string]$existingApp.notes
        if ((Get-IntuneNotesIdentityTags -Notes $existingNotes) -contains $identityTag) { $appBody['notes'] = $existingNotes }
        else {
            # A legacy tag is replaced by the identity tag; other text stays.
            $kept = @($existingNotes -split "`r?`n" | Where-Object { $_.Trim() -notlike 'AppPackager:legacy:*' }) -join "`r`n"
            if (-not [string]::IsNullOrWhiteSpace($kept)) { $appBody['notes'] = $identityTag + "`r`n" + $kept.Trim() }
        }
        $appId = [string]$existingApp.id
        # The detection rules, commands and setup file describe the new content,
        # so they reach the app in the same request that commits that content.
        # A failed upload leaves the app as it was.
        Write-Log ("Updating Intune Win32 app    : {0} (app id {1})" -f $Manifest.AppName, $appId)
    }
    else {
        Write-Log ("Creating Intune Win32 app    : {0}" -f $Manifest.AppName)
        $app = Invoke-GraphJson -Method POST -Uri "$GraphBase/deviceAppManagement/mobileApps" -Token $token -Body $appBody
        $appId = [string]$app.id
        Write-Log ("Intune app id                : {0}" -f $appId)
    }

    $lobBase = "$GraphBase/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp"
    $content = Invoke-GraphJson -Method POST -Uri "$lobBase/contentVersions" -Token $token -Body @{}
    $contentId = [string]$content.id

    $payload = Export-IntuneWinPayload -Path $IntuneWinPath -Destination ([System.IO.Path]::GetTempFileName())
    try {
        $fileBody = @{
            '@odata.type' = '#microsoft.graph.mobileAppContentFile'
            name          = $meta.FileName
            size          = [long]$meta.UnencryptedContentSize
            sizeEncrypted = [long]$payload.Size
            manifest      = $null
            isDependency  = $false
        }
        $file = Invoke-GraphJson -Method POST -Uri "$lobBase/contentVersions/$contentId/files" -Token $token -Body $fileBody
        $fileId = [string]$file.id
        $fileUri = "$lobBase/contentVersions/$contentId/files/$fileId"

        $deadline = (Get-Date).AddSeconds($PollTimeoutSec)
        do {
            Start-Sleep -Seconds 3
            $file = Invoke-GraphJson -Method GET -Uri $fileUri -Token $token
            if ([string]$file.uploadState -match 'Failed|TimedOut') { throw "Azure storage URI request failed: $($file.uploadState)" }
        } until ($file.azureStorageUri -or (Get-Date) -gt $deadline)
        if (-not $file.azureStorageUri) { throw 'Timed out waiting for the Azure storage upload URI.' }

        Write-Log "Uploading encrypted payload  : $([math]::Round($payload.Size / 1MB, 1)) MB"
        Invoke-AzureBlobUpload -Uri ([string]$file.azureStorageUri) -FilePath $payload.Path

        $commitBody = @{
            fileEncryptionInfo = @{
                encryptionKey        = $meta.EncryptionKey
                macKey               = $meta.MacKey
                initializationVector = $meta.InitializationVector
                mac                  = $meta.Mac
                profileIdentifier    = $meta.ProfileIdentifier
                fileDigest           = $meta.FileDigest
                fileDigestAlgorithm  = $meta.FileDigestAlgorithm
            }
        }
        Invoke-GraphJson -Method POST -Uri "$fileUri/commit" -Token $token -Body $commitBody | Out-Null

        $deadline = (Get-Date).AddSeconds($PollTimeoutSec)
        do {
            Start-Sleep -Seconds 5
            $file = Invoke-GraphJson -Method GET -Uri $fileUri -Token $token
            if ([string]$file.uploadState -match 'commitFileFailed|commitFileTimedOut') { throw "Content commit failed: $($file.uploadState)" }
        } until ([string]$file.uploadState -eq 'commitFileSuccess' -or (Get-Date) -gt $deadline)
        if ([string]$file.uploadState -ne 'commitFileSuccess') { throw "Timed out waiting for content commit (last state: $($file.uploadState))." }

        $finalBody = @{
            '@odata.type'           = '#microsoft.graph.win32LobApp'
            committedContentVersion = $contentId
        }
        if ($existingApp) {
            foreach ($key in @($appBody.Keys)) { $finalBody[$key] = $appBody[$key] }
        }
        Invoke-GraphJson -Method PATCH -Uri "$GraphBase/deviceAppManagement/mobileApps/$appId" -Token $token -Body $finalBody | Out-Null

        Write-Log ("Published to Intune          : {0} (app id {1})" -f $Manifest.AppName, $appId)
        return $appId
    }
    finally {
        Remove-Item -LiteralPath $payload.Path -Force -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function Get-IntuneWinEncryptionInfo, Export-IntuneWinPayload, Get-MsGraphToken, Invoke-GraphJson, Invoke-AzureBlobUpload, ConvertTo-IntuneWin32Rules, Publish-IntuneWin32App, Get-IntuneCompatibilityFindings, ConvertTo-IntuneRequirementRules, ConvertFrom-RequirementsJson, Get-IntuneRequirementSpecs, Get-IntuneRequirementDropKey, Get-IntuneAllowedArchitecture, Get-IntuneIdentityTag, Get-IntuneScriptRuleSettings, New-IntuneDetectionScriptRule
