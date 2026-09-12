#Requires -Version 5.1
# Options -> Script signing round trip: defaults are all off, a commit
# reaches the preferences file and reloads, and a cancelled dialog leaves
# nothing behind. No certificate store is written and no dialog opens.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('New-ScriptSigningPanel', 'Read-Preferences', 'Resolve-FirstRunCompleted',
                    'Get-WorkbenchCommand', 'Get-WorkbenchSigningPolicy', 'Get-WorkbenchSigningPolicyJson', 'Get-WorkbenchSigningPolicyDigest')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw ('function not found: ' + $name) }
    # Global, the way the shell publishes its functions: a GetNewClosure
    # handler runs in a module whose scope chain ends at global, so a
    # script-scoped copy would be invisible to the panel's own handlers.
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$fixtureDir = Join-Path $env:TEMP ('ap-signing-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
$script:fixturePath = Join-Path $fixtureDir 'preferences.json'
function Get-PreferencesPath { $script:fixturePath }

$checks = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:checks += 1
    Write-Host ('  ok  ' + $Message)
}

try {
    $script:Prefs = Read-Preferences
    $signing = $script:Prefs.ScriptSigning
    & $assert ($null -ne $signing) 'preferences carry a root-level ScriptSigning block'
    foreach ($flag in @('SignDetection', 'SignRequirements', 'SignDeployment', 'RequireDetection', 'RequireRequirements', 'RequireDeployment', 'TimestampRequired')) {
        if ([bool]$signing.$flag) { throw ('FAIL: ' + $flag + ' does not default to off') }
    }
    & $assert ($true) 'every signing switch defaults to off'
    & $assert ([string]$signing.StoreLocation -eq 'CurrentUser' -and [string]$signing.HashAlgorithm -eq 'SHA256') 'the default identity is CurrentUser with SHA-256'

    # Cancel: the panel is built and edited, then discarded without commit.
    $cancelPanel = New-ScriptSigningPanel
    $cancelPanel.Element.FindName('chkSignDeployment').IsChecked = $true
    $cancelPanel.Element.FindName('txtSignTimestamp').Text = 'http://discarded.example'
    & $assert (-not [bool]$script:Prefs.ScriptSigning.SignDeployment) 'a partial edit does not reach preferences before commit'

    $before = Get-WorkbenchSigningPolicyDigest
    $script:Prefs | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:fixturePath -Encoding UTF8
    $reloadedAfterCancel = Read-Preferences
    & $assert (-not [bool]$reloadedAfterCancel.ScriptSigning.SignDeployment) 'Cancel persists nothing'
    & $assert ([string]$reloadedAfterCancel.ScriptSigning.TimestampServer -eq '') 'Cancel discards the timestamp server edit'

    # OK: the panel commits into preferences, which are then written.
    $panel = New-ScriptSigningPanel
    & $assert ([string]$panel.Name -eq 'Script Signing') 'the panel registers under its own Options section'
    $panel.Element.FindName('chkSignDetection').IsChecked = $true
    $panel.Element.FindName('chkRequireDetection').IsChecked = $true
    $panel.Element.FindName('chkSignDeployment').IsChecked = $true
    $panel.Element.FindName('txtSignTimestamp').Text = ' http://timestamp.example '
    $panel.Element.FindName('chkSignTimestampRequired').IsChecked = $true
    $panel.Element.FindName('cboSignStore').SelectedItem = 'LocalMachine'
    & $panel.Commit
    $script:Prefs | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:fixturePath -Encoding UTF8

    $reloaded = Read-Preferences
    & $assert ([bool]$reloaded.ScriptSigning.SignDetection) 'sign detection survives save and reload'
    & $assert ([bool]$reloaded.ScriptSigning.RequireDetection) 'require detection survives save and reload'
    & $assert ([bool]$reloaded.ScriptSigning.SignDeployment) 'sign deployment survives save and reload'
    & $assert (-not [bool]$reloaded.ScriptSigning.SignRequirements) 'an untouched switch stays off'
    & $assert ([string]$reloaded.ScriptSigning.TimestampServer -eq 'http://timestamp.example') 'the timestamp server is trimmed and stored'
    & $assert ([bool]$reloaded.ScriptSigning.TimestampRequired) 'the timestamp requirement survives'
    & $assert ([string]$reloaded.ScriptSigning.StoreLocation -eq 'LocalMachine') 'the store selection survives'
    & $assert ([string]$reloaded.ScriptSigning.CertificateThumbprint -eq '') 'no certificate is selected without an explicit choice'

    $after = Get-WorkbenchSigningPolicyDigest
    & $assert ($after -ne $before -and $after.Length -eq 64) 'the policy digest changes with the policy'

    # A stored thumbprint that is not 40 hex characters is not adopted.
    $tampered = $reloaded | ConvertTo-Json -Depth 6 | ConvertFrom-Json
    $tampered.ScriptSigning.CertificateThumbprint = 'not-a-thumbprint'
    $tampered | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:fixturePath -Encoding UTF8
    & $assert ([string](Read-Preferences).ScriptSigning.CertificateThumbprint -eq '') 'a malformed thumbprint is rejected on load'

    Write-Host ''
    Write-Host ('PASS: Script signing options round trip, {0} checks' -f $script:checks)
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}
