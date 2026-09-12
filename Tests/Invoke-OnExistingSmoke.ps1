#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the existing-application overwrite decision and the
    GUI's conflict-detection parse.

.DESCRIPTION
    Extracts Resolve-OnExistingBehavior from AppPackagerCommon.psm1 and
    Get-ExistingConflictFromOutput from start-apppackager.ps1 via the AST, then
    asserts the precedence chain (parameter, then APP_PACKAGER_ON_EXISTING,
    then Skip), rejection of an unrecognized value, and that the marker the
    engine emits is the one the GUI matches. No CM connection, no WPF
    assemblies, no window.

.EXAMPLE
    .\Tests\Invoke-OnExistingSmoke.ps1
#>

[CmdletBinding()]
param(
    [string]$ModulePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Packagers\AppPackagerCommon.psm1'),
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1')
)

$ErrorActionPreference = 'Stop'

# Dot-sourcing happens at script scope on purpose: a helper function that
# dot-sourced would confine the extracted definitions to its own scope.
foreach ($target in @(
    @{ Path = $ModulePath; Names = @('Resolve-OnExistingBehavior') }
    @{ Path = $ScriptPath; Names = @('Get-ExistingConflictFromOutput') }
)) {
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($target.Path, [ref]$null, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        foreach ($e in $errors) { Write-Host ("PARSE  {0}: {1}" -f (Split-Path -Leaf $target.Path), $e.Message) -ForegroundColor Red }
        exit 1
    }
    foreach ($name in $target.Names) {
        $fn = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
        }, $true) | Select-Object -First 1
        if (-not $fn) { Write-Host ("MISSING function {0} in {1}" -f $name, (Split-Path -Leaf $target.Path)) -ForegroundColor Red; exit 1 }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
}

# The marker literal is defined once at module scope; read it out of the source
# so a rename on either side of the seam fails here rather than in the field.
$moduleText = Get-Content -LiteralPath $ModulePath -Raw
$markerMatch = [regex]::Match($moduleText, "OnExistingConflictMarker\s*=\s*'([^']+)'")
$marker = if ($markerMatch.Success) { $markerMatch.Groups[1].Value } else { '' }

$sampleLog = @"
[2026-09-02 10:00:00] [WARN ] Application already exists    : Audacity (v3.7.9, unchanged)
[2026-09-02 10:00:00] [INFO ] Deployment type(s) validated : Audacity
[2026-09-02 10:00:00] [INFO ] $marker app='Audacity' version='3.7.9'
"@
$cleanLog = "[2026-09-02 10:00:00] [INFO ] Created MECM application     : Audacity"

$env:APP_PACKAGER_ON_EXISTING = $null
Remove-Item Env:\APP_PACKAGER_ON_EXISTING -ErrorAction SilentlyContinue
$defaulted = Resolve-OnExistingBehavior
$paramOnly = Resolve-OnExistingBehavior -Requested 'Overwrite'

$env:APP_PACKAGER_ON_EXISTING = 'Fail'
$fromEnv = Resolve-OnExistingBehavior
$paramWins = Resolve-OnExistingBehavior -Requested 'Skip'
$caseFolded = $null
$env:APP_PACKAGER_ON_EXISTING = 'overwrite'
$caseFolded = Resolve-OnExistingBehavior

$env:APP_PACKAGER_ON_EXISTING = 'Nonsense'
$envRejected = $false
try { Resolve-OnExistingBehavior } catch { $envRejected = ($_.Exception.Message -match 'APP_PACKAGER_ON_EXISTING') }

Remove-Item Env:\APP_PACKAGER_ON_EXISTING -ErrorAction SilentlyContinue
$paramRejected = $false
try { Resolve-OnExistingBehavior -Requested 'Replace' } catch { $paramRejected = ($_.Exception.Message -match 'parameter') }

$parsed = Get-ExistingConflictFromOutput -Output $sampleLog
$parsedClean = Get-ExistingConflictFromOutput -Output $cleanLog

$cases = @(
    @{ Name = 'No parameter and no env resolves to Skip';       Actual = { $defaulted.Behavior };   Expected = 'Skip' }
    @{ Name = 'Default reports its source';                     Actual = { $defaulted.Source };     Expected = 'default' }
    @{ Name = 'Parameter alone is honored';                     Actual = { $paramOnly.Behavior };   Expected = 'Overwrite' }
    @{ Name = 'Env var is honored when no parameter';           Actual = { $fromEnv.Behavior };     Expected = 'Fail' }
    @{ Name = 'Env var reports its source';                     Actual = { $fromEnv.Source };       Expected = 'APP_PACKAGER_ON_EXISTING' }
    @{ Name = 'Parameter outranks the env var';                 Actual = { $paramWins.Behavior };   Expected = 'Skip' }
    @{ Name = 'Env value is matched case-insensitively';        Actual = { $caseFolded.Behavior };  Expected = 'Overwrite' }
    @{ Name = 'Unrecognized env value throws';                  Actual = { $envRejected };          Expected = $true }
    @{ Name = 'Unrecognized parameter value throws';            Actual = { $paramRejected };        Expected = $true }
    @{ Name = 'Marker literal is defined in the module';        Actual = { [bool]$marker };         Expected = $true }
    @{ Name = 'GUI parses the app name out of the marker';      Actual = { $parsed.AppName };       Expected = 'Audacity' }
    @{ Name = 'GUI parses the version out of the marker';       Actual = { $parsed.Version };       Expected = '3.7.9' }
    @{ Name = 'Clean output reports no conflict';               Actual = { $null -eq $parsedClean }; Expected = $true }
    @{ Name = 'Empty output reports no conflict';               Actual = { $null -eq (Get-ExistingConflictFromOutput -Output '') }; Expected = $true }
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
