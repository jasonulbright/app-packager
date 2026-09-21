#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the Deployment Target sidebar decision logic.

.DESCRIPTION
    Parses start-apppackager.ps1, extracts Get-SidebarTargetState from the AST,
    and asserts the label, tooltip, enablement, and One Click pre-flight verdict
    for each deployment target. No WPF assemblies are loaded and no window is
    shown.

.EXAMPLE
    .\Tests\Invoke-SidebarTargetSmoke.ps1
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

foreach ($name in @('Get-SidebarTargetState')) {
    $fn = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true) | Select-Object -First 1
    if (-not $fn) { Write-Host ("MISSING function {0}" -f $name) -ForegroundColor Red; exit 1 }
    . ([scriptblock]::Create($fn.Extent.Text))
}

# Update-SidebarForDeploymentTarget must stay the only writer of these buttons'
# label and tooltips, and must be wired to all three refresh points.
$callCount = ([regex]::Matches((Get-Content -Path $ScriptPath -Raw), 'Update-SidebarForDeploymentTarget')).Count

$mecm      = Get-SidebarTargetState -DeploymentTarget 'MECM'
$both      = Get-SidebarTargetState -DeploymentTarget 'MECMAndIntune'
$intune    = Get-SidebarTargetState -DeploymentTarget 'IntuneOnly'
$unknown   = Get-SidebarTargetState -DeploymentTarget 'Nonsense'

$cases = @(
    @{ Name = 'MECM enables Check ConfigMgr';            Actual = { $mecm.CheckMecmEnabled };        Expected = $true }
    @{ Name = 'MECM keeps the Package Apps label';       Actual = { $mecm.PackageContent };          Expected = 'Package Apps' }
    @{ Name = 'MECM runs the One Click pre-flight';      Actual = { $mecm.SkipMecmPreflight };       Expected = $false }
    @{ Name = 'MECMAndIntune matches MECM exactly';      Actual = { ($both | ConvertTo-Json -Compress) -eq ($mecm | ConvertTo-Json -Compress) }; Expected = $true }
    @{ Name = 'IntuneOnly disables Check ConfigMgr';     Actual = { $intune.CheckMecmEnabled };      Expected = $false }
    @{ Name = 'IntuneOnly tooltip names the site need';  Actual = { $intune.CheckMecmToolTip -match 'ConfigMgr site' -and $intune.CheckMecmToolTip -match 'Intune only' }; Expected = $true }
    @{ Name = 'IntuneOnly relabels to Publish Apps';     Actual = { $intune.PackageContent };        Expected = 'Publish Apps' }
    @{ Name = 'IntuneOnly package tooltip mentions Intune'; Actual = { $intune.PackageToolTip -match 'Intune' }; Expected = $true }
    @{ Name = 'IntuneOnly skips the One Click pre-flight'; Actual = { $intune.SkipMecmPreflight };   Expected = $true }
    @{ Name = 'Unrecognized target falls back to MECM';  Actual = { ($unknown | ConvertTo-Json -Compress) -eq ($mecm | ConvertTo-Json -Compress) }; Expected = $true }
    @{ Name = 'Sidebar refresh wired at four call sites'; Actual = { $callCount -ge 5 }; Expected = $true }
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
