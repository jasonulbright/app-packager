#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the first-run setup wizard decision logic.

.DESCRIPTION
    Parses start-apppackager.ps1, extracts the two decision functions
    (Resolve-FirstRunCompleted, Test-FirstRunWizardNeeded) from the AST, and
    exercises the first-run, configured, suppressed, and migration paths. No
    WPF assemblies are loaded and no window is shown.

.EXAMPLE
    .\Tests\Invoke-FirstRunSmoke.ps1
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
    foreach ($e in $errors) { Write-Host ("PARSE  {0}" -f $e.Message) -ForegroundColor Red }
    exit 1
}

$wanted = @('Resolve-FirstRunCompleted', 'Test-FirstRunWizardNeeded')
foreach ($name in $wanted) {
    $fn = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true) | Select-Object -First 1
    if (-not $fn) { Write-Host ("MISSING function {0}" -f $name) -ForegroundColor Red; exit 1 }
    . ([scriptblock]::Create($fn.Extent.Text))
}

$cases = @(
    @{ Name = 'No preferences file shows the wizard'
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ FirstRunCompleted = (Resolve-FirstRunCompleted -StoredValue $null -PreferencesFileExisted $false) }) }
       Expected = $true }
    @{ Name = 'Migration: existing file without the flag skips the wizard'
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ FirstRunCompleted = (Resolve-FirstRunCompleted -StoredValue $null -PreferencesFileExisted $true) }) }
       Expected = $false }
    @{ Name = 'Flag false in an existing file shows the wizard'
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ FirstRunCompleted = (Resolve-FirstRunCompleted -StoredValue $false -PreferencesFileExisted $true) }) }
       Expected = $true }
    @{ Name = "Don't-show-again stamped the flag; wizard stays away"
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ FirstRunCompleted = (Resolve-FirstRunCompleted -StoredValue $true -PreferencesFileExisted $true) }) }
       Expected = $false }
    @{ Name = 'Missing property on the prefs object shows the wizard'
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ SiteCode = 'MCM' }) }
       Expected = $true }
    @{ Name = 'Non-boolean stored value coerces to completed'
       Actual = { Test-FirstRunWizardNeeded -Prefs ([pscustomobject]@{ FirstRunCompleted = (Resolve-FirstRunCompleted -StoredValue 'true' -PreferencesFileExisted $true) }) }
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
