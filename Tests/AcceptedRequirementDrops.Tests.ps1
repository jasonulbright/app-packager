#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Intune.AcceptedRequirementDrops is a saved list with an Options field, reaches the publisher, and the review dialog stores it.

.DESCRIPTION
    The setting holds the requirements that have no Intune rule and may be left out of an Intune
    app; each entry is the identity tag of the app in Intune, a bar and the condition id. The
    review before a publish to Intune shows the findings of the staged apps and, when a
    requirement blocks, offers to accept its drop.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-IntunePublishConfigForContext', 'Test-DeploymentTargetPublishesToWsus', 'Test-DeploymentTargetPublishesToIntune', 'Get-OptionsSectionDefinition',
        'Invoke-PackagerIntunePublish', 'Show-IntuneReviewDialog', 'Get-IntuneReviewFinding', 'Confirm-IntuneRequirementReview') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Read-PrefsWithIntune {
        param([hashtable]$Intune)
        $script:PrefsFile = Join-Path $TestDrive ('prefs-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ Intune = [pscustomobject]$Intune } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        Read-Preferences
    }
    # Stand-ins for the launcher functions that the review calls.
    function Find-NewestStageManifestForPackager { param($PackagerPath, $DownloadRoot) }
    function Add-LogLine { param($Message) $script:Log += $Message }
    function Save-Preferences { param($Prefs) }
    function Get-UserErrorText { param($ErrorRecord, $Fallback) [string]$ErrorRecord.Exception.Message }
}

Describe 'Intune.AcceptedRequirementDrops preference' {
    It 'defaults to an empty list when no file exists' {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        @((Read-Preferences).Intune.AcceptedRequirementDrops).Count | Should -Be 0
    }

    It 'reads the entries from the file' {
        $prefs = Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; AcceptedRequirementDrops = @('AppPackager:a/default|x', 'AppPackager:b/default|y') }
        @($prefs.Intune.AcceptedRequirementDrops) | Should -Be @('AppPackager:a/default|x', 'AppPackager:b/default|y')
    }

    It 'reads one entry that is stored as text' {
        $prefs = Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; AcceptedRequirementDrops = 'AppPackager:a/default|x' }
        @($prefs.Intune.AcceptedRequirementDrops) | Should -Be @('AppPackager:a/default|x')
    }

    It 'drops blank and repeated entries' {
        $prefs = Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; AcceptedRequirementDrops = @('  ', 'k|x', 'k|x', '') }
        @($prefs.Intune.AcceptedRequirementDrops) | Should -Be @('k|x')
    }

    It 'keeps the empty list in a file from before the key' {
        $prefs = Read-PrefsWithIntune @{ DeploymentTarget = 'ConfigMgr'; TenantId = 'contoso.onmicrosoft.com' }
        @($prefs.Intune.AcceptedRequirementDrops).Count | Should -Be 0
    }

    It 'reaches the publish configuration' {
        try { $secret = ConvertTo-SecureString -String 'placeholder' -AsPlainText -Force | ConvertFrom-SecureString }
        catch { Set-ItResult -Skipped -Because 'DPAPI is not available in this session.'; return }
        $prefs = [pscustomobject]@{ Intune = [pscustomobject]@{
                DeploymentTarget = 'IntuneOnly'; PublishToIntune = $true; TenantId = 'contoso.onmicrosoft.com'
                ClientId = '11111111-2222-3333-4444-555555555555'; ClientSecretProtected = $secret; MinimumWindowsRelease = 'Windows10_21H2'
                AcceptedRequirementDrops = @('k|x') } }
        $config = Get-IntunePublishConfigForContext -Prefs $prefs -Target 'IntuneOnly'
        @($config.AcceptedRequirementDrops) | Should -Be @('k|x')
    }
}

Describe 'Options field of Intune.AcceptedRequirementDrops' {
    BeforeAll {
        $section = Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'intune' }
        $script:Field = $section.Fields | Where-Object { $_.Id -eq 'int.drops' }
    }

    It 'has a field in Options, Connections, Intune' {
        $script:Field.Pref | Should -Be 'Intune.AcceptedRequirementDrops'
        $script:Field.Title | Should -Be 'Accepted requirement drops'
        $script:Field.Kind | Should -Be 'Text'
    }

    It 'shows the entries on one line and reads them back' {
        (& $script:Field.ToText @('a|x', 'b|y')) | Should -Be 'a|x; b|y'
        $one = & $script:Field.FromText 'a|x; b|y'
        ($one -join ',') | Should -Be 'a|x,b|y'
        $two = & $script:Field.FromText "a|x`r`nb|y;a|x"
        ($two -join ',') | Should -Be 'a|x,b|y'
    }

    It 'reads an empty field as an empty list and not as text' {
        $value = & $script:Field.FromText '  '
        ($null -ne $value -and $value -is [System.Array]) | Should -BeTrue
        @($value).Count | Should -Be 0
    }
}

Describe 'The publish to Intune hands the requirements to the publisher' {
    BeforeEach {
        $script:Log = @()
        Mock Read-StageManifest { [pscustomobject]@{ AppName = 'Contoso Tool'; Architecture = 'x64' } }
        Mock Publish-IntuneWin32App { 'app-1' }
    }

    It 'passes the requirement specs of the run and the accepted drops' {
        $note = [pscustomobject]@{ ManifestPath = 'C:\x\stage-manifest.json'; LocalPath = 'C:\x\a.intunewin' }
        $config = @{ TenantId = 't'; ClientId = 'c'; ClientSecret = 's'; AcceptedRequirementDrops = @('k|x') }
        $json = '{"SchemaVersion":1,"Rules":[{"ConditionId":"cpu-arch","Value":"x64"}]}'

        $result = Invoke-PackagerIntunePublish -IntuneNote $note -Config $config -RequirementsJson $json

        $result.Ok | Should -BeTrue
        Should -Invoke Publish-IntuneWin32App -Times 1 -Exactly -ParameterFilter { @($Requirements).Count -eq 1 -and $Requirements[0].ConditionId -eq 'cpu-arch' -and @($AcceptedRequirementDrops) -contains 'k|x' }
    }

    It 'reports requirements JSON that is not valid and publishes nothing' {
        $note = [pscustomobject]@{ ManifestPath = 'C:\x\stage-manifest.json'; LocalPath = 'C:\x\a.intunewin' }
        $config = @{ TenantId = 't'; ClientId = 'c'; ClientSecret = 's' }

        $result = Invoke-PackagerIntunePublish -IntuneNote $note -Config $config -RequirementsJson '{ not json'

        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'not valid JSON'
        Should -Invoke Publish-IntuneWin32App -Times 0 -Exactly
    }
}

Describe 'The review before a publish to Intune' {
    BeforeEach {
        $script:Log = @()
        $script:Asked = 0
        $script:Answer = 'Continue'
        $script:IntuneReviewPromptOverride = {
            param($Message, $CanAccept)
            $script:Asked++; $script:AskedMessage = $Message; $script:AskedCanAccept = $CanAccept
            $script:Answer
        }
        $script:Prefs = [pscustomobject]@{ DownloadRoot = 'C:\stage'; Intune = [pscustomobject]@{ AcceptedRequirementDrops = @() } }
        $script:Row = [pscustomobject]@{ Script = 'package-contoso.ps1'; FullPath = 'C:\p\package-contoso.ps1'; Application = 'Contoso Tool' }
        $script:Key = 'AppPackager:legacy:Contoso-Tool/default|no-such-condition'
        $script:Saved = 0
        Mock Find-NewestStageManifestForPackager { 'C:\stage\stage-manifest.json' }
        Mock Save-Preferences { $script:Saved++ }
        $script:Requirement = 'no-such-condition'
        Mock Read-StageManifest {
            [pscustomobject]@{
                AppName = 'Contoso Tool'; Architecture = 'x64'; SetupFile = 'install.bat'
                Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Contoso'; Is64Bit = $true }
                Requirements = @([pscustomobject]@{ ConditionId = $script:Requirement; Value = 'x64' })
            }
        }
    }

    AfterAll { $script:IntuneReviewPromptOverride = $null }

    It 'asks nothing when every finding is information only' {
        $script:Requirement = 'cpu-arch'

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeTrue
        $script:Asked | Should -Be 0
    }

    It 'asks nothing for an application that has no staged build' {
        Mock Find-NewestStageManifestForPackager { $null }

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeTrue
        $script:Asked | Should -Be 0
    }

    It 'shows a blocking requirement with the key to accept' {
        $script:Answer = 'Cancel'

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeFalse

        $script:Asked | Should -Be 1
        $script:AskedCanAccept | Should -BeTrue
        $script:AskedMessage | Should -Match 'RequirementNotTranslatable'
        $script:AskedMessage | Should -Match ([regex]::Escape($script:Key))
        $script:Saved | Should -Be 0
    }

    It 'stores the drop when the answer is Accept' {
        $script:Answer = 'Accept'

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeTrue

        @($script:Prefs.Intune.AcceptedRequirementDrops) | Should -Be @($script:Key)
        $script:Saved | Should -Be 1
    }

    It 'keeps the entries that exist and stores a key once' {
        $script:Prefs.Intune.AcceptedRequirementDrops = @('AppPackager:other/default|x', $script:Key.ToUpperInvariant())
        $script:Requirement = 'second-missing'
        $script:Answer = 'Accept'

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row, $script:Row)) | Should -BeTrue

        @($script:Prefs.Intune.AcceptedRequirementDrops).Count | Should -Be 3
        @($script:Prefs.Intune.AcceptedRequirementDrops)[0] | Should -Be 'AppPackager:other/default|x'
    }

    It 'continues without storing when the answer is Continue' {
        $script:Answer = 'Continue'

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeTrue

        @($script:Prefs.Intune.AcceptedRequirementDrops).Count | Should -Be 0
        $script:Saved | Should -Be 0
    }

    It 'shows an accepted drop as a review finding without the accept button' {
        $script:Prefs.Intune.AcceptedRequirementDrops = @($script:Key)

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeTrue

        $script:Asked | Should -Be 1
        $script:AskedCanAccept | Should -BeFalse
        $script:AskedMessage | Should -Match 'RequirementDropped'
    }

    It 'cancels the run when the accepted drops cannot be stored' {
        $script:Answer = 'Accept'
        Mock Save-Preferences { throw 'The file is locked.' }

        (Confirm-IntuneRequirementReview -Owner 'owner' -Rows @($script:Row)) | Should -BeFalse

        ($script:Log -join "`n") | Should -Match 'were not stored'
    }
}
