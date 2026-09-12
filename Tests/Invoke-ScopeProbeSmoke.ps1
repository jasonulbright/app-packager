#Requires -Version 5.1
[CmdletBinding()]
param([string]$ScriptPath)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1' }

# Extracts the launch-scope block from the application script and runs it in a
# child scope with a sample function, then checks that a GetNewClosure handler
# can call the function. Run this file with "& path" (child scope), not -File.
$text = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $ScriptPath).Path)
$start = $text.IndexOf('$script:ScopeProbe = $true')
$end = $text.IndexOf("Remove-Variable -Name ScopeProbe", $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Cannot locate the scope block.' }
$end = $text.IndexOf("`n", $end) + 1
$block = $text.Substring($start, $end - $start).Replace('$PSCommandPath', "'$($MyInvocation.MyCommand.Path.Replace("'", "''"))'")

function Get-SmokeScriptState { 'script function reached' }

$scriptScopeIsGlobal = $null
$script:ScopeProbe = $true
$scriptScopeIsGlobal = Test-Path -LiteralPath 'variable:global:ScopeProbe'
Remove-Variable -Name ScopeProbe -Scope Script -ErrorAction SilentlyContinue
if ($scriptScopeIsGlobal) { throw 'Run this smoke with "& <path>" so the script scope is not global.' }

$before = try { & { Get-SmokeScriptState }.GetNewClosure() } catch { 'unreachable' }
. ([scriptblock]::Create($block))
$after = try { & { Get-SmokeScriptState }.GetNewClosure() } catch { 'unreachable' }
Write-Output ("PS {0}: before={1}; after={2}" -f $PSVersionTable.PSVersion, $before, $after)
if ($before -ne 'unreachable') { throw 'Test setup did not reproduce the hidden function.' }
if ($after -ne 'script function reached') { throw 'Scope block did not expose the function.' }
Write-Output 'PASS: closure can call the script function after the scope block.'
