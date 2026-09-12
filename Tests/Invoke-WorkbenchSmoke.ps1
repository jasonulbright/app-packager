#Requires -Version 5.1
# Headless UI probe for the Application Workbench. Builds the real window
# off-screen, drives section switching, field edits, reset, Save and Save
# as against a temporary data root, then draft recovery and the unsaved
# switch prompt with the prompt stubbed.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerWorkbench.psm1') -Force -DisableNameChecking

$dataRoot = Join-Path $env:TEMP ('ap-workbench-' + [guid]::NewGuid().ToString('N'))
$env:APP_PACKAGER_WORKBENCH_ROOT = $dataRoot
[void](New-Item -ItemType Directory -Path $dataRoot -Force)

try {
    # Every function this file defines is reflected into this scope, the same
    # way the shell publishes them for its closure handlers.
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $published = New-Object System.Collections.Generic.List[string]
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
        Set-Item -Path ('function:global:' + $fn.Name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
        $published.Add($fn.Name)
    }

    $script:WorkbenchXamlPathOverride = Join-Path $root 'WorkbenchWindow.xaml'
    $PackagersRoot = Join-Path $root 'Packagers'
    $script:Prefs = [pscustomobject]@{
        EstimatedRuntimeMins = 15
        MaximumRuntimeMins   = 30
        DownloadRoot         = Join-Path $dataRoot 'dl'
        DeploymentConditions = [pscustomobject]@{ Apps = [pscustomobject]@{} }
        CommandOverrides     = [pscustomobject]@{ Apps = [pscustomobject]@{} }
        Intune               = [pscustomobject]@{ DeploymentTarget = 'MECM' }
        ScriptSigning        = [pscustomobject]@{
            SignDetection = $false; SignRequirements = $false; SignDeployment = $false
            RequireDetection = $false; RequireRequirements = $false; RequireDeployment = $false
            CertificateThumbprint = ''; StoreLocation = 'CurrentUser'
            TimestampServer = ''; TimestampRequired = $false; HashAlgorithm = 'SHA256'
        }
    }
    function Add-LogLine { param([string]$Message) }
    function Show-ThemedMessage { param($Owner, $Title, $Message, $Buttons, $Icon) return $script:StubMessageAnswer }
    function Set-DialogChromeFromOwner { param($Dialog, $Owner) }
    function Install-TitleBarDragFallback { param($Window) }
    $script:StubMessageAnswer = 'Yes'

    $probeOwner = New-Object System.Windows.Window
    $probeOwner.Width = 10; $probeOwner.Height = 10
    $probeOwner.WindowStartupLocation = 'Manual'
    $probeOwner.Left = -32000; $probeOwner.Top = -32000

    $assert = {
        param([bool]$Condition, [string]$Message)
        if (-not $Condition) { throw ('FAIL: ' + $Message) }
    }
    $checks = 0
    $ok = { param([string]$m) $script:PassedChecks += 1; Write-Host ('  ok  ' + $m) }
    $script:PassedChecks = 0

    # --- Pass 1: sections, edits, reset, Save as -------------------------
    $script:SavedProfileId = ''
    $script:SavedApplicationId = ''
    Show-ApplicationWorkbench -Owner $probeOwner -Probe {
        param($p)
        $find = $p.Find

        & $assert ($p.Sections.Items.Count -eq 7) 'the section list carries all seven sections'
        & $ok 'seven sections'

        for ($i = 0; $i -lt 7; $i++) {
            $p.Sections.SelectedIndex = $i
            $visible = @($p.Panels | Where-Object { $_.Visibility -eq [System.Windows.Visibility]::Visible })
            & $assert ($visible.Count -eq 1 -and $visible[0] -eq $p.Panels[$i]) ('section ' + $i + ' shows exactly its own panel')
        }
        & $ok 'section switching shows one panel at a time'

        $txtDisplayName = & $find 'txtDisplayName'
        $lblDisplayNameSrc = & $find 'lblDisplayNameSrc'
        & $assert ($lblDisplayNameSrc.Text -like 'Inherited:*') 'an untouched field reports the inherited value'
        $txtDisplayName.Text = 'Probe Display Name'
        & $p.Commit
        & $p.Refresh
        & $assert ($lblDisplayNameSrc.Text -like 'Custom*') 'an edited field reports Custom with the inherited value'
        & $assert ($p.State.Dirty) 'editing marks the profile dirty'
        & $ok 'inherited and custom indicators'

        (& $find 'btnDisplayNameReset').RaiseEvent(
            (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
        & $assert ([string]$txtDisplayName.Text -eq '') 'per-field reset clears the editor value'
        & $assert ($lblDisplayNameSrc.Text -like 'Inherited:*') 'per-field reset returns the field to inherit'
        & $ok 'per-field reset'

        $txtScript = & $find 'txtScript'
        $gutter = & $find 'txtScriptGutter'
        (& $find 'cboInstallMode').SelectedItem = 'Custom'
        $txtScript.Text = "Write-Output 'one'`nWrite-Output 'two'`nWrite-Output 'three'"
        & $assert ((@($gutter.Text -split "`r?`n")).Count -eq 3) 'the gutter carries one label per script line'
        & $ok 'line-number gutter tracks the script'

        $broken = @(Get-WorkbenchParseDiagnostics -Text 'if (')
        & $assert ($broken.Count -gt 0) 'a broken script produces a parse diagnostic'
        & $assert (@(Get-WorkbenchParseDiagnostics -Text 'Write-Output 1').Count -eq 0) 'a valid script produces no diagnostic'
        & $assert (-not (Test-WorkbenchDetectionScriptOutput -Text 'exit 0')) 'a detector that only exits is rejected'
        & $assert (Test-WorkbenchDetectionScriptOutput -Text 'Write-Output "found"') 'a detector writing to STDOUT is accepted'
        & $ok 'parse diagnostics and the detection output contract'

        $txtDisplayName.Text = 'Probe Managed Title'
        (& $find 'chkEstimatedDefault').IsChecked = $false
        (& $find 'txtEstimatedMinutes').Text = '20'
        $script:StubNameAnswer = 'Probe Managed'
        $script:WorkbenchNamePromptOverride = { param($t, $v) $script:StubNameAnswer }
        $saved = & $p.Save 'Probe Managed'
        & $assert ([bool]$saved) 'Save as writes the profile'
        & $assert (-not $p.State.Dirty) 'a save clears the dirty state'
        & $assert ($p.State.ProfileId -ne 'default') 'a save of the default profile creates a named profile'
        $script:SavedProfileId = $p.State.ProfileId
        $script:SavedApplicationId = [string]$p.State.Application.ApplicationId
        & $ok 'Save as creates a named profile'

        $findings = & $p.Validate
        & $assert (@($findings).Count -gt 0) 'validation returns findings'
        foreach ($chip in @('chipContent', 'chipMecm', 'chipIntune')) {
            & $assert ((& $find $chip).Text -notlike '*Not validated*') ($chip + ' reports a state after validation')
        }
        $isolated = Get-WorkbenchChipState -Validated $true -Findings @(
            [pscustomobject]@{ Severity = 'Blocking'; Code = 'VARIANT-INTUNE'; Message = 'x' })
        & $assert ($isolated.Intune -eq 'Unsupported') 'an Intune blocker marks Intune unsupported'
        & $assert ($isolated.Mecm -eq 'Ready') 'an Intune blocker never blocks MECM'
        & $assert ($isolated.Content -eq 'Ready') 'an Intune blocker never blocks the content build'
        & $ok 'three independent status chips'
    }

    & $assert (-not [string]::IsNullOrWhiteSpace($script:SavedProfileId)) 'the probe recorded the saved profile id'
    $stored = Get-Profile -ApplicationId $script:SavedApplicationId -ProfileId $script:SavedProfileId
    & $assert ([string]$stored.Name -eq 'Probe Managed') 'the stored profile keeps its name'
    & $assert ([int]$stored.Revision -ge 1) 'the stored profile carries a revision'
    & $assert ([string]$stored.Application.DisplayName -eq 'Probe Managed Title') 'the display-name override survived the save'
    & $assert ([int]$stored.Timing.EstimatedMinutes -eq 20) 'the estimated-duration override survived the save'
    & $ok 'saved profile round-trips through the definition model'

    # --- Pass 2: draft recovery ------------------------------------------
    $draft = Get-Profile -ApplicationId $script:SavedApplicationId -ProfileId $script:SavedProfileId
    $draftData = $draft | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $draftData.Application.DisplayName = 'Recovered From Draft'
    [void](Save-Draft -ApplicationId $script:SavedApplicationId -ProfileId $script:SavedProfileId -Draft $draftData)
    $script:StubMessageAnswer = 'Yes'
    Show-ApplicationWorkbench -Owner $probeOwner -PreselectApplicationId $script:SavedApplicationId -Probe {
        param($p)
        & $assert ($p.State.ProfileId -eq $script:SavedProfileId) 'the active profile opens by default'
        & $assert ([string](& $p.Find 'txtDisplayName').Text -eq 'Recovered From Draft') 'the recovered draft populates the editor'
    }
    & $ok 'draft recovery'

    # --- Pass 3: the unsaved-switch prompt, with the prompt stubbed -------
    $script:UnsavedPromptCalls = 0
    $script:WorkbenchUnsavedPromptOverride = {
        param($message)
        $script:UnsavedPromptCalls += 1
        return $script:StubUnsavedAnswer
    }
    $script:StubMessageAnswer = 'No'
    $script:StubUnsavedAnswer = 'Cancel'
    Show-ApplicationWorkbench -Owner $probeOwner -PreselectApplicationId $script:SavedApplicationId -Probe {
        param($p)
        (& $p.Find 'txtPublisher').Text = 'Probe Publisher'
        & $assert ($p.State.Dirty) 'the edit is pending'
        $before = [string]$p.State.ProfileId
        $allowed = & $p.ConfirmDiscard 'unit probe'
        & $assert (-not $allowed) 'Cancel refuses the switch'
        & $assert ($p.State.Dirty) 'Cancel keeps the pending edit'
        & $assert ([string]$p.State.ProfileId -eq $before) 'Cancel leaves the profile selection alone'

        $script:StubUnsavedAnswer = 'Discard'
        $allowed = & $p.ConfirmDiscard 'unit probe'
        & $assert ([bool]$allowed) 'Discard allows the switch'
        & $assert (-not $p.State.Dirty) 'Discard clears the pending edit'
        & $assert ($script:UnsavedPromptCalls -eq 2) 'the prompt ran once per attempt'
    }
    & $ok 'unsaved-switch prompt path'

    Write-Host ''
    Write-Host ('PASS: Application Workbench probe, {0} checks' -f $script:PassedChecks)
}
finally {
    foreach ($name in @($published)) {
        Remove-Item -Path ('function:global:' + $name) -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $dataRoot -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\APP_PACKAGER_WORKBENCH_ROOT -ErrorAction SilentlyContinue
}
