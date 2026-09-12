#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath,
    [string]$InstallerPath = (Join-Path $env:TEMP 'TeamViewerHost.exe'),
    [string]$ExpectedVersion = '15.81.6.0',
    [string]$StringTableExePath = (Join-Path $env:SystemRoot 'System32\notepad.exe')
)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path $PSScriptRoot '..\Packagers\package-teamviewerhost.ps1' }

# Lifts Get-TeamViewerHostExeVersion out of the packager and runs it against an
# EXE whose version resource string table is empty and one where it is filled.
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path -LiteralPath $ScriptPath).Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
$fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-TeamViewerHostExeVersion' }, $false)
if (-not $fn) { throw "Cannot locate Get-TeamViewerHostExeVersion in $ScriptPath." }
. ([scriptblock]::Create($fn.Extent.Text))

if (-not (Test-Path -LiteralPath $InstallerPath)) { throw "Installer payload not found: $InstallerPath" }
if (-not (Test-Path -LiteralPath $StringTableExePath)) { throw "String-table sample not found: $StringTableExePath" }

$script:Failed = 0

function Test-Assert {
    param([Parameter(Mandatory)][string]$Label, [bool]$Condition, [string]$Detail = '')

    if ($Condition) { Write-Output "PASS: $Label" }
    else {
        $script:Failed++
        Write-Output ("FAIL: $Label" + $(if ($Detail) { " - $Detail" } else { '' }))
    }
}

Write-Output "Packager     : $((Resolve-Path -LiteralPath $ScriptPath).Path)"
Write-Output "PowerShell   : $($PSVersionTable.PSVersion)"

$vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($InstallerPath)
Write-Output ""
Write-Output "--- $InstallerPath ---"
Write-Output ("    string table ProductVersion: '{0}'" -f $vi.ProductVersion)
Write-Output ("    string table FileVersion   : '{0}'" -f $vi.FileVersion)
Write-Output ("    fixed block Product parts  : {0}.{1}.{2}.{3}" -f $vi.ProductMajorPart, $vi.ProductMinorPart, $vi.ProductBuildPart, $vi.ProductPrivatePart)

$hostVersion = Get-TeamViewerHostExeVersion -ExePath $InstallerPath
Write-Output "    returned                   : $hostVersion"
Test-Assert -Label "empty string table falls back to the fixed block" -Condition ($hostVersion -eq $ExpectedVersion) -Detail "got '$hostVersion', expected '$ExpectedVersion'"

$sampleVersion = Get-TeamViewerHostExeVersion -ExePath $StringTableExePath
Write-Output ""
Write-Output "--- $StringTableExePath ---"
Write-Output "    returned                   : $sampleVersion"
Test-Assert -Label "populated string table returns a dotted version" -Condition ($sampleVersion -match '^\d+(\.\d+)+') -Detail "got '$sampleVersion'"

Write-Output ""
if ($script:Failed) { throw "$script:Failed check(s) failed." }
Write-Output "All TeamViewer Host version checks passed."
