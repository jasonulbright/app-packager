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
function ConvertTo-DeploymentTargetName { param($Destinations) if ($Destinations.ConfigMgr) { 'MECM' } elseif ($Destinations.Intune) { 'IntuneOnly' } elseif ($Destinations.WSUS) { 'WSUSOnly' } else { 'MECM' } }
function Invoke-RefreshGrid { $script:RefreshCount++ }
function Update-SidebarForSystems { $script:SidebarCount++ }
function Show-ThemedMessage {
    param($Owner, $Title, $Message, $Buttons, $Icon)
    $script:SaveError = $Message
}
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-WizardCallback {
    param([string]$Action, [bool]$Suppress, [string[]]$Systems = @('ConfigMgr'))
    $script:Prefs = [pscustomobject]@{
        SiteCode = ''; ProviderMachineName = ''; FileShareRoot = ''; DownloadRoot = ''
        FirstRunCompleted = $false
        Systems = [pscustomobject]@{ ConfigMgr = $true; Intune = $false; Wsus = $false }
        AppFlow = [pscustomobject]@{ DefaultDestinations = [pscustomobject]@{ ConfigMgr = $true; WSUS = $false; Intune = $false } }
        Intune = [pscustomobject]@{
            TenantId = ''; ClientId = ''; ClientSecretProtected = ''
            DeploymentTarget = 'MECM'; PublishToIntune = $false
        }
        Wsus = [pscustomobject]@{ ServerName = ''; PortNumber = 8530; UseSsl = $false }
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
    $chkWizMecm = [pscustomobject]@{ IsChecked = ($Systems -contains 'ConfigMgr') }
    $chkWizIntune = [pscustomobject]@{ IsChecked = ($Systems -contains 'Intune') }
    $chkWizWsus = [pscustomobject]@{ IsChecked = ($Systems -contains 'WSUS') }
    $Target = if ($Systems -contains 'ConfigMgr') { 'MECM' } elseif ($Systems -contains 'Intune') { 'IntuneOnly' } elseif ($Systems -contains 'WSUS') { 'WSUSOnly' } else { 'MECM' }
    $txtWizSC = [pscustomobject]@{ Text = ' TEST ' }
    $txtWizProvider = [pscustomobject]@{ Text = ' provider.example ' }
    $txtWizFS = [pscustomobject]@{ Text = ' \\server\content ' }
    $txtWizDL = [pscustomobject]@{ Text = ' C:\temp\packages ' }
    $txtWizTenant = [pscustomobject]@{ Text = ' tenant ' }
    $txtWizClient = [pscustomobject]@{ Text = ' client ' }
    $pwdWizSecret = [pscustomobject]@{ SecurePassword = (New-Object Security.SecureString) }
    $txtWizWsusServer = [pscustomobject]@{ Text = ' wsus.example ' }
    $txtWizWsusPort = [pscustomobject]@{ Text = ' 8531 ' }
    $chkWizWsusSsl = [pscustomobject]@{ IsChecked = $true }
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
        Assert-True ($saved.Intune.DeploymentTarget -eq $Target) 'The One Click destination was not derived from the selected systems.'
        Assert-True (($saved.AppFlow.DefaultDestinations.ConfigMgr -eq ($Systems -contains 'ConfigMgr')) -and ($saved.AppFlow.DefaultDestinations.WSUS -eq ($Systems -contains 'WSUS')) -and ($saved.AppFlow.DefaultDestinations.Intune -eq ($Systems -contains 'Intune'))) 'The default One Click destinations were not derived from the selected systems.'
        Assert-True (($saved.Systems.ConfigMgr -eq ($Systems -contains 'ConfigMgr')) -and ($saved.Systems.Intune -eq ($Systems -contains 'Intune')) -and ($saved.Systems.Wsus -eq ($Systems -contains 'WSUS'))) 'The selected systems were not saved.'
        Assert-True ([bool]$saved.Intune.PublishToIntune -eq ($Systems -contains 'Intune')) 'The Intune publish toggle does not follow the Intune box.'
        if ($Systems -contains 'ConfigMgr') { Assert-True ($saved.SiteCode -eq 'TEST') 'ConfigMgr fields were not trimmed and saved.' }
        else { Assert-True ($saved.SiteCode -eq '') 'Setup without ConfigMgr saved ConfigMgr fields.' }
        if ($Systems -contains 'Intune') { Assert-True ($saved.Intune.TenantId -eq 'tenant' -and $saved.Intune.ClientId -eq 'client') 'Intune fields were not saved.' }
        else { Assert-True ($saved.Intune.TenantId -eq '') 'Setup without Intune saved Intune fields.' }
        if ($Systems -contains 'WSUS') {
            Assert-True ($saved.Wsus.ServerName -eq 'wsus.example' -and $saved.Wsus.PortNumber -eq 8531 -and $saved.Wsus.UseSsl) 'WSUS fields were not trimmed and saved.'
        }
        else { Assert-True ($saved.Wsus.ServerName -eq '') 'Setup without WSUS saved WSUS fields.' }
    } else {
        Assert-True ($script:SaveCount -eq [int]$Suppress) 'Skip/close suppression was not persisted correctly.'
        Assert-True (-not $script:FirstRunDlgSaved) 'Skipping incorrectly marked setup as saved.'
        Assert-True ($script:RefreshCount -eq 0 -and $script:SidebarCount -eq 0) 'Skipping unexpectedly refreshed the UI.'
    }
    Write-Output ("PASS: {0} / suppress={1} / systems={2}" -f $Action, $Suppress, ($Systems -join ' + '))
}

$failed = 0
foreach ($scenario in @(
    @{ Action = 'Save'; Suppress = $false; Systems = @('ConfigMgr') },
    @{ Action = 'Save'; Suppress = $true; Systems = @('ConfigMgr', 'Intune') },
    @{ Action = 'Save'; Suppress = $false; Systems = @('Intune') },
    @{ Action = 'Save'; Suppress = $false; Systems = @('ConfigMgr', 'WSUS') },
    @{ Action = 'Save'; Suppress = $true; Systems = @('WSUS') },
    @{ Action = 'Save'; Suppress = $false; Systems = @('ConfigMgr', 'Intune', 'WSUS') },
    @{ Action = 'Save'; Suppress = $false; Systems = @('Intune', 'WSUS') },
    @{ Action = 'Skip'; Suppress = $false },
    @{ Action = 'Skip'; Suppress = $true },
    @{ Action = 'Close'; Suppress = $true },
    @{ Action = 'Retry'; Suppress = $false }
)) {
    try { Test-WizardCallback @scenario }
    catch { $failed++; Write-Output "FAIL: $($scenario.Action): $($_.Exception.Message)" }
}

# The wizard's WSUS port follows its SSL box until the operator types a port,
# on real WPF controls. Setting IsChecked without a click is what a keyboard
# or UI Automation toggle does.
try {
    Add-Type -AssemblyName PresentationFramework
    $portStart = $source.IndexOf('$wizPortFollowsSsl')
    $portEnd = $source.IndexOf('$currentTarget', [Math]::Max(0, $portStart))
    if ($portStart -lt 0 -or $portEnd -lt 0) { throw 'Cannot locate the wizard port-follow registrations.' }
    $portFollow = [scriptblock]::Create($source.Substring($portStart, $portEnd - $portStart))
    $script:Prefs = [pscustomobject]@{ Wsus = [pscustomobject]@{ PortNumber = 8530 } }
    $txtWizWsusPort = New-Object System.Windows.Controls.TextBox
    $txtWizWsusPort.Text = '8530'
    $chkWizWsusSsl = New-Object System.Windows.Controls.CheckBox
    . $portFollow
    $chkWizWsusSsl.IsChecked = $true
    Assert-True ($txtWizWsusPort.Text -eq '8531') 'Ticking Use SSL did not move the default port to 8531.'
    $chkWizWsusSsl.IsChecked = $false
    Assert-True ($txtWizWsusPort.Text -eq '8530') 'Clearing Use SSL did not move the port back to 8530.'
    $script:Prefs.Wsus.PortNumber = 443
    $txtWizWsusPort = New-Object System.Windows.Controls.TextBox
    $txtWizWsusPort.Text = '443'
    $chkWizWsusSsl = New-Object System.Windows.Controls.CheckBox
    . $portFollow
    $chkWizWsusSsl.IsChecked = $true
    Assert-True ($txtWizWsusPort.Text -eq '443') 'Use SSL replaced a port the operator chose.'
    Write-Output 'PASS: WSUS port follows Use SSL'
}
catch { $failed++; Write-Output "FAIL: WSUS port follow: $($_.Exception.Message)" }

# The log line after a saved setup names every default One Click destination.
try {
    Import-Module (Join-Path (Split-Path (Resolve-Path -LiteralPath $ScriptPath).Path -Parent) 'Packagers\AppPackagerOneClick.psd1') -Force
    $doneStart = $source.IndexOf('[void]$dlg.ShowDialog()')
    $completion = [scriptblock]::Create($source.Substring($doneStart, $source.LastIndexOf('}') - $doneStart).Replace('[void]$dlg.ShowDialog()', ''))
    function Add-LogLine { param([string]$Message) $script:SetupLog = $Message }
    foreach ($case in @(
        @{ Defaults = [pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $false }; Expected = 'ConfigMgr and WSUS' },
        @{ Defaults = [pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $true }; Expected = 'Intune' }
    )) {
        $script:Prefs = [pscustomobject]@{ AppFlow = [pscustomobject]@{ DefaultDestinations = $case.Defaults }; Intune = [pscustomobject]@{ DeploymentTarget = 'MECMAndWSUS' } }
        $script:FirstRunDlgSaved = $true
        $script:SetupLog = ''
        . $completion
        Assert-True ($script:SetupLog -like ('Setup complete. One Click publishes to {0};*' -f $case.Expected)) ('The setup log line reads: ' + $script:SetupLog)
    }
    Write-Output 'PASS: setup log names the default destinations'
}
catch { $failed++; Write-Output "FAIL: setup log line: $($_.Exception.Message)" }
if ($failed) { throw "$failed first-run callback scenario(s) failed." }
