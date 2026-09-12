#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the packager icon pack decision and verify logic.

.DESCRIPTION
    Parses start-apppackager.ps1, extracts the pure icon-pack functions
    (Read-IconPackManifest, Test-IconPackAppVersion, Get-IconPackChecksum,
    Test-IconPackChecksum, Select-IconPackAsset, Get-IconPackStatusText) from
    the AST, and exercises the manifest read, version compare, checksum
    parse/verify, asset selection, and status-text paths against a scratch
    folder. No WPF assemblies are loaded, no window is shown, and no network
    call is made.

.EXAMPLE
    .\Tests\Invoke-IconPackSmoke.ps1
#>

[CmdletBinding()]
param(
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1')
)

$ErrorActionPreference = 'Stop'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    foreach ($e in $errors) { Write-Host ("PARSE  {0}: {1}" -f (Split-Path -Leaf $ScriptPath), $e.Message) -ForegroundColor Red }
    exit 1
}

$wanted = @(
    'Read-IconPackManifest', 'Test-IconPackAppVersion', 'Get-IconPackChecksum',
    'Test-IconPackChecksum', 'Select-IconPackAsset', 'Get-IconPackStatusText',
    'Get-IconPackRoot', 'Get-IconPackManifestPath'
)
foreach ($name in $wanted) {
    $fn = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true) | Select-Object -First 1
    if (-not $fn) { Write-Host ("MISSING function {0}" -f $name) -ForegroundColor Red; exit 1 }
    . ([scriptblock]::Create($fn.Extent.Text))
}

# Select-IconPackAsset reads the asset names off script scope, as it does in the app.
$script:IconPackAssetName = 'icon-pack.zip'
$script:IconPackSumsName  = 'checksums.txt'

$scratch = Join-Path ([IO.Path]::GetTempPath()) ('apicons-smoke-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $scratch 'Packagers\Icons') -Force | Out-Null

$goodManifest = Join-Path $scratch 'Packagers\Icons\manifest.json'
Set-Content -LiteralPath $goodManifest -Encoding UTF8 -Value @'
{
  "PackVersion": "1.2.0",
  "MinAppVersion": "1.5.0.0",
  "Icons": [
    { "File": "7zip.ico", "Packager": "7zip" },
    { "File": "aimp.png", "Packager": "aimp" }
  ]
}
'@

$emptyManifest = Join-Path $scratch 'empty.json'
Set-Content -LiteralPath $emptyManifest -Encoding UTF8 -Value '{ "PackVersion": "1.0.0", "MinAppVersion": "1.5.0.0", "Icons": [] }'

$badManifest = Join-Path $scratch 'bad.json'
Set-Content -LiteralPath $badManifest -Encoding UTF8 -Value 'not json {'

$payload = Join-Path $scratch 'icon-pack.zip'
Set-Content -LiteralPath $payload -Encoding ASCII -NoNewline -Value 'pack-bytes'
$payloadHash = (Get-FileHash -LiteralPath $payload -Algorithm SHA256).Hash.ToLowerInvariant()

$sums = "$payloadHash *icon-pack.zip`n0000000000000000000000000000000000000000000000000000000000000000  checksums.txt"

$releaseComplete = [pscustomobject]@{
    tag_name = 'v1.0.0'
    assets   = @(
        [pscustomobject]@{ name = 'icon-pack.zip'; browser_download_url = 'https://example.invalid/icon-pack.zip' }
        [pscustomobject]@{ name = 'checksums.txt'; browser_download_url = 'https://example.invalid/checksums.txt' }
    )
}
$releaseEmpty = [pscustomobject]@{ tag_name = 'v1.0.0'; assets = @() }

try {
    $cases = @(
        @{ Name = 'Manifest read returns the pack version'
           Actual = { (Read-IconPackManifest -Path $goodManifest).PackVersion }
           Expected = '1.2.0' }
        @{ Name = 'Manifest read counts the icons'
           Actual = { (Read-IconPackManifest -Path $goodManifest).IconCount }
           Expected = 2 }
        @{ Name = 'Manifest read returns the minimum app version'
           Actual = { (Read-IconPackManifest -Path $goodManifest).MinAppVersion }
           Expected = '1.5.0.0' }
        @{ Name = 'Empty icon array counts zero'
           Actual = { (Read-IconPackManifest -Path $emptyManifest).IconCount }
           Expected = 0 }
        @{ Name = 'Malformed manifest reads as absent'
           Actual = { $null -eq (Read-IconPackManifest -Path $badManifest) }
           Expected = $true }
        @{ Name = 'Missing manifest reads as absent'
           Actual = { $null -eq (Read-IconPackManifest -Path (Join-Path $scratch 'nope.json')) }
           Expected = $true }
        @{ Name = 'Empty manifest path reads as absent'
           Actual = { $null -eq (Read-IconPackManifest -Path '') }
           Expected = $true }

        @{ Name = 'Newer app satisfies the minimum'
           Actual = { Test-IconPackAppVersion -MinAppVersion '1.5.0.0' -CurrentVersion '1.5.0.4' }
           Expected = $true }
        @{ Name = 'Equal version satisfies the minimum'
           Actual = { Test-IconPackAppVersion -MinAppVersion '1.5.0.4' -CurrentVersion '1.5.0.4' }
           Expected = $true }
        @{ Name = 'Older app fails the minimum'
           Actual = { Test-IconPackAppVersion -MinAppVersion '1.6.0.0' -CurrentVersion '1.5.0.4' }
           Expected = $false }
        @{ Name = 'Revision-only minimum is compared'
           Actual = { Test-IconPackAppVersion -MinAppVersion '1.5.0.5' -CurrentVersion '1.5.0.4' }
           Expected = $false }
        @{ Name = 'Absent minimum is satisfied'
           Actual = { Test-IconPackAppVersion -MinAppVersion '' -CurrentVersion '1.5.0.4' }
           Expected = $true }
        @{ Name = 'Unparseable minimum is satisfied'
           Actual = { Test-IconPackAppVersion -MinAppVersion 'nightly' -CurrentVersion '1.5.0.4' }
           Expected = $true }
        @{ Name = 'Unparseable running version is satisfied'
           Actual = { Test-IconPackAppVersion -MinAppVersion '1.5.0.0' -CurrentVersion '' }
           Expected = $true }

        @{ Name = 'Checksum line with the star marker parses'
           Actual = { Get-IconPackChecksum -ChecksumText $sums -FileName 'icon-pack.zip' }
           Expected = $payloadHash }
        @{ Name = 'Two-space checksum line parses'
           Actual = { Get-IconPackChecksum -ChecksumText $sums -FileName 'checksums.txt' }
           Expected = '0000000000000000000000000000000000000000000000000000000000000000' }
        @{ Name = 'Unlisted file has no checksum'
           Actual = { $null -eq (Get-IconPackChecksum -ChecksumText $sums -FileName 'absent.zip') }
           Expected = $true }
        @{ Name = 'Empty checksum text has no checksum'
           Actual = { $null -eq (Get-IconPackChecksum -ChecksumText '' -FileName 'icon-pack.zip') }
           Expected = $true }

        @{ Name = 'Matching hash verifies'
           Actual = { Test-IconPackChecksum -FilePath $payload -ExpectedSha256 $payloadHash }
           Expected = $true }
        @{ Name = 'Hash compare ignores case'
           Actual = { Test-IconPackChecksum -FilePath $payload -ExpectedSha256 $payloadHash.ToUpperInvariant() }
           Expected = $true }
        @{ Name = 'Wrong hash fails verification'
           Actual = { Test-IconPackChecksum -FilePath $payload -ExpectedSha256 ('0' * 64) }
           Expected = $false }
        @{ Name = 'Absent expected hash fails verification'
           Actual = { Test-IconPackChecksum -FilePath $payload -ExpectedSha256 $null }
           Expected = $false }
        @{ Name = 'Missing file fails verification'
           Actual = { Test-IconPackChecksum -FilePath (Join-Path $scratch 'nope.zip') -ExpectedSha256 $payloadHash }
           Expected = $false }

        @{ Name = 'Complete release yields the pack URL'
           Actual = { (Select-IconPackAsset -Release $releaseComplete).PackUrl }
           Expected = 'https://example.invalid/icon-pack.zip' }
        @{ Name = 'Complete release yields the checksum URL'
           Actual = { (Select-IconPackAsset -Release $releaseComplete).SumsUrl }
           Expected = 'https://example.invalid/checksums.txt' }
        @{ Name = 'Release with no assets yields no pack URL'
           Actual = { $null -eq (Select-IconPackAsset -Release $releaseEmpty).PackUrl }
           Expected = $true }
        @{ Name = 'Absent release yields no selection'
           Actual = { $null -eq (Select-IconPackAsset -Release $null) }
           Expected = $true }

        @{ Name = 'Status text reports version and count'
           Actual = { (Get-IconPackStatusText -Manifest (Read-IconPackManifest -Path $goodManifest)) -match 'v1\.2\.0, 2 icons' }
           Expected = $true }
        @{ Name = 'Status text singularizes one icon'
           Actual = { (Get-IconPackStatusText -Manifest ([pscustomobject]@{ PackVersion = '1.0.0'; MinAppVersion = '1.5.0.0'; IconCount = 1 })) -match '1 icon$' }
           Expected = $true }
        @{ Name = 'Status text calls out an empty pack'
           Actual = { (Get-IconPackStatusText -Manifest (Read-IconPackManifest -Path $emptyManifest)) -match 'no icons yet' }
           Expected = $true }
        @{ Name = 'Status text reports a missing pack'
           Actual = { (Get-IconPackStatusText -Manifest $null) -match 'Not installed' }
           Expected = $true }

        @{ Name = 'Pack root sits under Packagers'
           Actual = { (Get-IconPackRoot -AppRoot $scratch) -eq (Join-Path $scratch 'Packagers\Icons') }
           Expected = $true }
        @{ Name = 'Manifest path sits in the pack root'
           Actual = { (Get-IconPackManifestPath -AppRoot $scratch) -eq $goodManifest }
           Expected = $true }
    )

    $failed = 0
    foreach ($case in $cases) {
        $actual = & $case.Actual
        if ($actual -eq $case.Expected) {
            Write-Host ("PASS  {0}" -f $case.Name) -ForegroundColor Green
        } else {
            Write-Host ("FAIL  {0} (expected {1}, got {2})" -f $case.Name, $case.Expected, $actual) -ForegroundColor Red
            $failed++
        }
    }

    Write-Host ("`n{0} passed, {1} failed." -f ($cases.Count - $failed), $failed)
    exit ([int]($failed -gt 0))
}
finally {
    Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
}
