#Requires -Version 5.1
[CmdletBinding()]
param([string]$ScriptPath)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1' }

# Drives the conflict dialog's real button registrations with .NET events in
# place of WPF, in a normal script scope. No application launch.
Add-Type -TypeDefinition @'
using System;
public class ConflictTestButton {
    public event EventHandler Click;
    public void RaiseClick() { if (Click != null) Click(this, EventArgs.Empty); }
}
'@

$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path -LiteralPath $ScriptPath).Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
$fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-ExistingConflictDialog' }, $false)
$source = $fn.Extent.Text
$start = $source.IndexOf('$script:ConflictDialogChoice = ''Skip''')
$end = $source.IndexOf('[void]$dlg.ShowDialog()', $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Cannot locate conflict dialog registrations.' }
$registrations = [scriptblock]::Create($source.Substring($start, $end - $start))
$returnLine = [scriptblock]::Create(([regex]::Match($source.Substring($end), 'return @\{[^
]*\}')).Value)

function Test-ConflictButton {
    param([string]$Button, [string]$Expected)
    $buttons = @{ btnSkip = (New-Object ConflictTestButton); btnOverwrite = (New-Object ConflictTestButton); btnCancel = (New-Object ConflictTestButton) }
    $closed = @{ Value = $false }
    $dlg = [pscustomobject]@{ Buttons = $buttons; Closed = $closed }
    $dlg | Add-Member -MemberType ScriptMethod -Name FindName -Value { param($name) $this.Buttons[$name] }
    $dlg | Add-Member -MemberType ScriptMethod -Name Close -Value { $this.Closed.Value = $true }
    $chkAll = [pscustomobject]@{ IsChecked = $true }
    . $registrations
    $buttons[$Button].RaiseClick()
    $result = & $returnLine
    if (-not $closed.Value) { throw "$Button did not close the dialog." }
    if ($result.Choice -ne $Expected) { throw "$Button returned '$($result.Choice)', expected '$Expected'." }
    if (-not $result.ApplyToAll) { throw "$Button lost the apply-to-all flag." }
    Write-Output "PASS: $Button -> $Expected"
}

$failed = 0
foreach ($case in @(@{ Button = 'btnSkip'; Expected = 'Skip' }, @{ Button = 'btnOverwrite'; Expected = 'Overwrite' }, @{ Button = 'btnCancel'; Expected = 'Cancel' })) {
    try { Test-ConflictButton @case } catch { $failed++; Write-Output "FAIL: $($case.Button): $($_.Exception.Message)" }
}
if ($failed) { throw "$failed conflict dialog scenario(s) failed." }
