#Requires -Version 5.1
[CmdletBinding()]
param([string]$ScriptPath)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1' }

# Exercise the actual event registrations in a script scope, with .NET events
# standing in for WPF. No application launch, downloads, or preference writes.
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
public class FirstRunTestButton {
    public event EventHandler Click;
    public void RaiseClick() { if (Click != null) Click(this, EventArgs.Empty); }
}
public class FirstRunTestDialog {
    public event CancelEventHandler Closing;
    public bool Closed;
    public void Close() {
        var args = new CancelEventArgs();
        if (Closing != null) Closing(this, args);
        Closed = !args.Cancel;
    }
}
'@

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    (Resolve-Path -LiteralPath $ScriptPath).Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
$wizard = $ast.Find({ param($n)
    $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $n.Name -eq 'Show-FirstRunWizard'
}, $false)
$source = $wizard.Extent.Text
$start = $source.IndexOf('$script:FirstRunDlgSaved = $false')
$end = $source.IndexOf('[void]$dlg.ShowDialog()', $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Cannot locate wizard event registrations.' }
$registrations = [scriptblock]::Create($source.Substring($start, $end - $start))

function Save-Preferences {
    param($Prefs)
    if ($script:SaveFails) { throw 'Simulated preference write failure' }
    $script:SavedJson = $Prefs | ConvertTo-Json -Depth 5
    $script:SaveCount++
}
function Invoke-RefreshGrid { $script:RefreshCount++ }
function Update-SidebarForDeploymentTarget { $script:SidebarCount++ }
function Show-ThemedMessage {
    param($Owner, $Title, $Message, $Buttons, $Icon)
    $script:SaveError = $Message
}
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-WizardCallback {
    param([string]$Action, [bool]$Suppress, [string]$Target = 'MECM')
    $script:Prefs = [pscustomobject]@{
        SiteCode = ''; ProviderMachineName = ''; FileShareRoot = ''; DownloadRoot = ''
        FirstRunCompleted = $false
        Intune = [pscustomobject]@{
            TenantId = ''; ClientId = ''; ClientSecretProtected = ''
            DeploymentTarget = 'MECM'; PublishToIntune = $false
        }
    }
    $script:SaveCount = 0
    $script:RefreshCount = 0
    $script:SidebarCount = 0
    $script:SaveError = ''
    $script:SavedJson = ''
    $script:SaveFails = $Action -eq 'Retry'
    $dlg = New-Object FirstRunTestDialog
    $btnWizSave = New-Object FirstRunTestButton
    $btnWizSkip = New-Object FirstRunTestButton
    $cboTarget = [pscustomobject]@{ SelectedItem = [pscustomobject]@{ Tag = $Target } }
    $txtWizSC = [pscustomobject]@{ Text = ' TEST ' }
    $txtWizProvider = [pscustomobject]@{ Text = ' provider.example ' }
    $txtWizFS = [pscustomobject]@{ Text = ' \\server\content ' }
    $txtWizDL = [pscustomobject]@{ Text = ' C:\temp\packages ' }
    $txtWizTenant = [pscustomobject]@{ Text = ' tenant ' }
    $txtWizClient = [pscustomobject]@{ Text = ' client ' }
    $pwdWizSecret = [pscustomobject]@{ SecurePassword = (New-Object Security.SecureString) }
    $chkWizDontAsk = [pscustomobject]@{ IsChecked = $Suppress }

    . $registrations
    if ($Action -eq 'Retry') {
        $btnWizSave.RaiseClick()
        Assert-True ($script:SaveError -eq 'Simulated preference write failure') 'Write failure was not surfaced.'
        Assert-True (-not $dlg.Closed -and -not $script:FirstRunDlgSaved) 'Failed save closed the wizard.'
        $script:SaveFails = $false
        $script:SaveError = ''
    }
    if ($Action -eq 'Skip') { $btnWizSkip.RaiseClick() }
    elseif ($Action -eq 'Close') { $dlg.Close() }
    else { $btnWizSave.RaiseClick() }

    Assert-True $dlg.Closed ('Wizard did not close: ' + $script:SaveError)
    Assert-True ([string]::IsNullOrEmpty($script:SaveError)) ('Save failed: ' + $script:SaveError)
    if ($Action -in @('Save', 'Retry')) {
        Assert-True ($script:SaveCount -eq 1) 'Save must persist exactly once, including when suppression is checked.'
        Assert-True $script:FirstRunDlgSaved 'Saved state did not reach the parent script.'
        Assert-True ($script:RefreshCount -eq 1 -and $script:SidebarCount -eq 1) 'Post-save UI helpers were not called.'
        $saved = $script:SavedJson | ConvertFrom-Json
        Assert-True $saved.FirstRunCompleted 'Saved preferences did not suppress setup.'
        Assert-True ($saved.Intune.DeploymentTarget -eq $Target) 'Deployment target was not saved.'
        if ($Target -ne 'IntuneOnly') { Assert-True ($saved.SiteCode -eq 'TEST') 'MECM fields were not trimmed and saved.' }
        if ($Target -ne 'MECM') { Assert-True ($saved.Intune.TenantId -eq 'tenant' -and $saved.Intune.PublishToIntune) 'Intune fields were not saved.' }
    } else {
        Assert-True ($script:SaveCount -eq [int]$Suppress) 'Skip/close suppression was not persisted correctly.'
        Assert-True (-not $script:FirstRunDlgSaved) 'Skipping incorrectly marked setup as saved.'
        Assert-True ($script:RefreshCount -eq 0 -and $script:SidebarCount -eq 0) 'Skipping unexpectedly refreshed the UI.'
    }
    Write-Output "PASS: $Action / suppress=$Suppress / target=$Target"
}

$failed = 0
foreach ($scenario in @(
    @{ Action = 'Save'; Suppress = $false; Target = 'MECM' },
    @{ Action = 'Save'; Suppress = $true; Target = 'MECMAndIntune' },
    @{ Action = 'Save'; Suppress = $false; Target = 'IntuneOnly' },
    @{ Action = 'Skip'; Suppress = $false },
    @{ Action = 'Skip'; Suppress = $true },
    @{ Action = 'Close'; Suppress = $true },
    @{ Action = 'Retry'; Suppress = $false }
)) {
    try { Test-WizardCallback @scenario }
    catch { $failed++; Write-Output "FAIL: $($scenario.Action): $($_.Exception.Message)" }
}
if ($failed) { throw "$failed first-run callback scenario(s) failed." }
