#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the update-check decision logic and the installer helpers.

.DESCRIPTION
    Parses start-apppackager.ps1, extracts the pure update-check functions
    (Get-AppVersion, ConvertFrom-ReleaseTag, Test-UpdateAvailable,
    Test-UpdateCheckDue) from the AST, dot-sources install.ps1 for its checksum
    and state-enumeration helpers, and exercises the throttle, version-compare,
    tag-parse, checksum-parse, and preservation paths. No WPF assemblies are
    loaded, no window is shown, and no network call is made.

.EXAMPLE
    .\Tests\Invoke-UpdateCheckSmoke.ps1
#>

[CmdletBinding()]
param(
    [string]$ScriptPath    = (Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1'),
    [string]$InstallerPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'install.ps1')
)

$ErrorActionPreference = 'Stop'

foreach ($path in @($ScriptPath, $InstallerPath)) {
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        foreach ($e in $errors) { Write-Host ("PARSE  {0}: {1}" -f (Split-Path -Leaf $path), $e.Message) -ForegroundColor Red }
        exit 1
    }
}

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)

$wanted = @('Get-AppVersion', 'ConvertFrom-ReleaseTag', 'Test-UpdateAvailable', 'Test-UpdateCheckDue')
foreach ($name in $wanted) {
    $fn = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true) | Select-Object -First 1
    if (-not $fn) { Write-Host ("MISSING function {0}" -f $name) -ForegroundColor Red; exit 1 }
    . ([scriptblock]::Create($fn.Extent.Text))
}

# Dot-sourcing install.ps1 loads its helpers without performing an install.
. $InstallerPath

$now  = [datetime]::new(2026, 9, 2, 12, 0, 0, [System.DateTimeKind]::Utc)
$sums = "932f2a5d3b92e4437765adfa72aa5871a57df11b999f0d102c09d7b29f5bd37a *AppPackager-1.5.0.3.zip`n0000000000000000000000000000000000000000000000000000000000000000  other.zip"

$scratch = Join-Path ([IO.Path]::GetTempPath()) ('apsmoke-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $scratch 'Packagers') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $scratch 'Logs') -Force | Out-Null
Set-Content -LiteralPath (Join-Path $scratch 'AppPackager.preferences.json') -Value '{}'
Set-Content -LiteralPath (Join-Path $scratch 'Packagers\packager-preferences.json') -Value '{}'
Set-Content -LiteralPath (Join-Path $scratch 'Logs\run.log') -Value 'x'
Set-Content -LiteralPath (Join-Path $scratch 'CHANGELOG.md') -Value 'x'

try {
    $cases = @(
        @{ Name = 'Version parses out of the script header'
           Actual = { [bool]([regex]::IsMatch((Get-AppVersion -ScriptPath $ScriptPath), '^\d+(\.\d+)+$')) }
           Expected = $true }
        @{ Name = 'Missing script yields no version'
           Actual = { $null -eq (Get-AppVersion -ScriptPath (Join-Path $scratch 'nope.ps1')) }
           Expected = $true }

        @{ Name = 'Release tag strips the v prefix'
           Actual = { ConvertFrom-ReleaseTag -Tag 'v1.5.0.4' }
           Expected = '1.5.0.4' }
        @{ Name = 'Bare numeric tag parses'
           Actual = { ConvertFrom-ReleaseTag -Tag '1.5.0.4' }
           Expected = '1.5.0.4' }
        @{ Name = 'Non-version tag is rejected'
           Actual = { $null -eq (ConvertFrom-ReleaseTag -Tag 'nightly') }
           Expected = $true }
        @{ Name = 'Empty tag is rejected'
           Actual = { $null -eq (ConvertFrom-ReleaseTag -Tag '') }
           Expected = $true }

        @{ Name = 'Newer release counts as an update'
           Actual = { Test-UpdateAvailable -CurrentVersion '1.5.0.3' -LatestVersion '1.5.0.4' }
           Expected = $true }
        @{ Name = 'Same version is not an update'
           Actual = { Test-UpdateAvailable -CurrentVersion '1.5.0.4' -LatestVersion '1.5.0.4' }
           Expected = $false }
        @{ Name = 'Older release is not an update'
           Actual = { Test-UpdateAvailable -CurrentVersion '1.5.0.4' -LatestVersion '1.4.0.24' }
           Expected = $false }
        @{ Name = 'Revision-only bump is an update'
           Actual = { Test-UpdateAvailable -CurrentVersion '1.5.0.9' -LatestVersion '1.5.1.0' }
           Expected = $true }
        @{ Name = 'Unparseable latest makes no claim'
           Actual = { Test-UpdateAvailable -CurrentVersion '1.5.0.4' -LatestVersion 'nightly' }
           Expected = $false }
        @{ Name = 'Unparseable current makes no claim'
           Actual = { Test-UpdateAvailable -CurrentVersion '' -LatestVersion '1.5.0.4' }
           Expected = $false }

        @{ Name = 'No prior check is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $null -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'Empty timestamp is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc '' -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'Garbage timestamp is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc 'not-a-date' -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'Check 23 hours ago is throttled'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $now.AddHours(-23).ToString('o') -NowUtc $now -IntervalHours 24 }
           Expected = $false }
        @{ Name = 'Check 25 hours ago is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $now.AddHours(-25).ToString('o') -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'Exactly the interval is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $now.AddHours(-24).ToString('o') -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'Future timestamp is due'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $now.AddHours(6).ToString('o') -NowUtc $now -IntervalHours 24 }
           Expected = $true }
        @{ Name = 'DateTime input is accepted'
           Actual = { Test-UpdateCheckDue -LastCheckUtc $now.AddHours(-1) -NowUtc $now -IntervalHours 24 }
           Expected = $false }

        @{ Name = 'Checksum line with the star marker parses'
           Actual = { Get-ChecksumForFile -ChecksumText $sums -FileName 'AppPackager-1.5.0.3.zip' }
           Expected = '932f2a5d3b92e4437765adfa72aa5871a57df11b999f0d102c09d7b29f5bd37a' }
        @{ Name = 'Two-space checksum line parses'
           Actual = { Get-ChecksumForFile -ChecksumText $sums -FileName 'other.zip' }
           Expected = '0000000000000000000000000000000000000000000000000000000000000000' }
        @{ Name = 'Unlisted file has no checksum'
           Actual = { $null -eq (Get-ChecksumForFile -ChecksumText $sums -FileName 'absent.zip') }
           Expected = $true }

        @{ Name = 'State enumeration finds the JSON and log files'
           Actual = { @(Get-PreservedStateFile -Root $scratch).Count }
           Expected = 3 }
        @{ Name = 'State enumeration excludes shipped content'
           Actual = { @(Get-PreservedStateFile -Root $scratch) -contains 'CHANGELOG.md' }
           Expected = $false }
        @{ Name = 'State enumeration on a missing folder is empty'
           Actual = { @(Get-PreservedStateFile -Root (Join-Path $scratch 'nope')).Count }
           Expected = 0 }
        @{ Name = 'An AppPackager folder is recognized'
           Actual = { Test-AppPackagerFolder -Path $scratch }
           Expected = $false }
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
