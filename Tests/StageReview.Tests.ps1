#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The review before staging: SkipConfigReview is a saved preference (default off) with an Options field, the
    review shows for the applications that have install options, Cancel skips the application, Continue sends
    the changed values to this run only, and a batch run or One Click never asks.
#>

BeforeDiscovery {
    $packagers = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')).Path 'Packagers'
    $script:ReviewedNames = @('package-ssms.ps1', 'package-dbeaver.ps1', 'package-adobereader.ps1', 'package-beyondcompare5.ps1', 'package-teamviewerhost.ps1') +
        @(Get-ChildItem -LiteralPath $packagers -Filter 'package-m365*.ps1' | ForEach-Object { $_.Name }) +
        @(Get-ChildItem -LiteralPath $packagers -Filter 'package-citrixworkspace-*.ps1' | ForEach-Object { $_.Name })
}

BeforeAll {
    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    function Get-LauncherFunction {
        param([string]$Name)
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $Name." }
        $node
    }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-OptionsSectionDefinition', 'Get-PackagerMetadata',
        'Get-StageReviewSection', 'Test-StageReviewNeeded', 'Get-PackagerPreferenceMirror', 'Get-StageReviewRunOptions',
        'Confirm-StageReview', 'Get-StageReviewDefinition', 'New-PackagerRunEnvironment', 'Invoke-PackagerStage',
        'Set-PackagerEnvironment', 'Set-ProcessStartInfoArgumentList', 'ConvertTo-ProcessArgument') {
        . ([scriptblock]::Create((Get-LauncherFunction $name).Extent.Text))
    }
    function Get-PreferencesPath { $script:PrefsFile }
    function Add-LogLine { param([string]$Message) $script:LogLines += $Message }
    function Show-StageReviewDialog {
        param($Owner, [string]$Application, [string]$SectionId, $Prefs)
        $script:Asked += $Application
        $answer = $script:Answers[$Application]
        if ($null -eq $answer) { return @{ Continue = $true; RunOptionsJson = '' } }
        return $answer
    }
    function Assert-PackagerStageIntegrity { param($Result, $PackagerPath, $DownloadRoot) }
    function Invoke-ProcessWithStreaming {
        param($StartInfo, $OutLog, $ErrLog, $StructuredLog, $LogTextBox)
        $list = @()
        if ($StartInfo.ArgumentList -and $StartInfo.ArgumentList.Count -gt 0) { $list = @($StartInfo.ArgumentList) } else { $list = @(([string]$StartInfo.Arguments) -split ' ') }
        $script:Child = @{ Arguments = $list; Env = $StartInfo.EnvironmentVariables['APP_PACKAGER_RUN_OPTIONS'] }
        $json = $script:Child.Env
        if ($json) {
            $files = ($json | ConvertFrom-Json).Files
            $script:Child.Files = @{}
            foreach ($p in $files.PSObject.Properties) { $script:Child.Files[$p.Name] = @{ Path = [string]$p.Value; Text = (Get-Content -LiteralPath $p.Value -Raw) } }
        }
        [pscustomobject]@{ ExitCode = 0; OutLog = $OutLog }
    }

    function New-DefaultPrefs {
        $script:PrefsFile = Join-Path $TestDrive 'missing.json'
        Read-Preferences
    }
    function Copy-Prefs { param($Prefs) ConvertTo-Json -InputObject $Prefs -Depth 8 | ConvertFrom-Json }
    $script:PackagersRoot = Join-Path $script:RepoRoot 'Packagers'
    # The packagers that have install options, and the Options section of each.
    $script:Reviewed = @{
        'package-ssms.ps1' = 'ssms'; 'package-dbeaver.ps1' = 'dbeaver'; 'package-adobereader.ps1' = 'adobe'
        'package-beyondcompare5.ps1' = 'bc5'; 'package-teamviewerhost.ps1' = 'teamviewer'
    }
    foreach ($file in Get-ChildItem -LiteralPath $script:PackagersRoot -Filter 'package-m365*.ps1') { $script:Reviewed[$file.Name] = 'm365' }
    foreach ($file in Get-ChildItem -LiteralPath $script:PackagersRoot -Filter 'package-citrixworkspace-*.ps1') { $script:Reviewed[$file.Name] = 'cwa' }
}

Describe 'SkipConfigReview' {
    It 'defaults to off when no file exists' {
        (New-DefaultPrefs).SkipConfigReview | Should -BeFalse
    }

    It 'reads <Value> from the file' -ForEach @(@{ Value = $true }, @{ Value = $false }) {
        $script:PrefsFile = Join-Path $TestDrive ('p-' + [guid]::NewGuid().ToString('N') + '.json')
        ([pscustomobject]@{ SkipConfigReview = $Value } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).SkipConfigReview | Should -Be $Value
    }

    It 'keeps off in a file from before the key' {
        $script:PrefsFile = Join-Path $TestDrive 'old.json'
        ([pscustomobject]@{ SiteCode = 'ABC' } | ConvertTo-Json) | Set-Content -LiteralPath $script:PrefsFile -Encoding ASCII
        (Read-Preferences).SkipConfigReview | Should -BeFalse
    }

    It 'has a field in Options, Packaging, Folders' {
        $section = Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'folders' }
        $section.Parent | Should -Be 'packaging'
        $field = $section.Fields | Where-Object { $_.Id -eq 'pkg.skipReview' }
        $field.Pref | Should -Be 'SkipConfigReview'
        $field.Kind | Should -Be 'Check'
    }
}

Describe 'Which applications have a review' {
    It 'maps <_> to a section that Options defines' -ForEach $script:ReviewedNames {
        $section = Get-StageReviewSection -Script $_
        $section | Should -Be $script:Reviewed[$_]
        @(Get-OptionsSectionDefinition | Where-Object { $_.Id -eq $section -and $_.Parent -eq 'applications' }).Count | Should -Be 1
    }

    It 'finds every packager file that it maps by name' {
        foreach ($name in 'package-ssms.ps1', 'package-dbeaver.ps1', 'package-adobereader.ps1', 'package-beyondcompare5.ps1', 'package-teamviewerhost.ps1') {
            Test-Path -LiteralPath (Join-Path $script:PackagersRoot $name) | Should -BeTrue -Because $name
        }
    }

    It 'maps the six Microsoft 365 packagers and the four Citrix Workspace packagers' {
        @($script:Reviewed.Values | Where-Object { $_ -eq 'm365' }).Count | Should -Be 6
        @($script:Reviewed.Values | Where-Object { $_ -eq 'cwa' }).Count | Should -Be 4
    }

    It 'gives no review to <_>' -ForEach 'package-7zip.ps1', 'package-teamviewer.ps1', 'package-firefox.ps1', '' {
        Get-StageReviewSection -Script $_ | Should -Be ''
    }

    It 'writes every field of a reviewed section to a preference that reaches the packager' {
        $mirror = Get-PackagerPreferenceMirror -Prefs (New-DefaultPrefs) -Folder $TestDrive
        $reachable = @($mirror.Keys) + 'M365Channel', 'M365DeployMode'
        $unreachable = foreach ($id in 'ssms', 'dbeaver', 'adobe', 'bc5', 'm365', 'cwa', 'teamviewer') {
            $section = Get-StageReviewDefinition -Definition (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq $id }) -Prefs (New-DefaultPrefs)
            foreach ($field in @($section.Fields | Where-Object { $_.Pref })) {
                $pref = [string]$field.Pref
                if ($pref -match '^(cwa|tv):') { continue }
                if ($reachable -notcontains ($pref -split '\.')[0]) { '{0}: {1}' -f $field.Id, $pref }
            }
        }
        @($unreachable) | Should -BeNullOrEmpty
    }

    It 'drops the imported configuration.xml rows from the Microsoft 365 review and keeps the generated settings' {
        $prefs = New-DefaultPrefs
        $section = Get-StageReviewDefinition -Definition (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'm365' }) -Prefs $prefs
        $ids = @($section.Fields | ForEach-Object { [string]$_.Id })
        foreach ($gone in 'm365.links', 'm365.useCustom', 'm365.file', 'm365.preview') { $ids | Should -Not -Contain $gone }
        foreach ($kept in 'm365.channel', 'm365.deploy', 'm365.languages', 'm365.exclude', 'm365.shared') { $ids | Should -Contain $kept }
        $ids | Should -Not -Contain 'm365.customNote'
        $prefs.M365CustomConfig.Enabled = $true
        $noted = Get-StageReviewDefinition -Definition (Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'm365' }) -Prefs $prefs
        @($noted.Fields | ForEach-Object { [string]$_.Id }) | Should -Contain 'm365.customNote'
    }

    It 'leaves the other sections as they are' {
        $definition = Get-OptionsSectionDefinition | Where-Object { $_.Id -eq 'ssms' }
        [object]::ReferenceEquals((Get-StageReviewDefinition -Definition $definition -Prefs (New-DefaultPrefs)), $definition) | Should -BeTrue
    }
}

Describe 'When the review shows' {
    It 'follows the section, the skip setting and the header: <Case>' -ForEach @(
        @{ Case = 'install options, review on'; Script = 'package-ssms.ps1'; Skip = $false; Header = $false; Expected = $true }
        @{ Case = 'install options, skipped'; Script = 'package-ssms.ps1'; Skip = $true; Header = $false; Expected = $false }
        @{ Case = 'install options, skipped, header forces it'; Script = 'package-ssms.ps1'; Skip = $true; Header = $true; Expected = $true }
        @{ Case = 'no install options, review on'; Script = 'package-7zip.ps1'; Skip = $false; Header = $false; Expected = $false }
        @{ Case = 'no install options, header forces it'; Script = 'package-7zip.ps1'; Skip = $true; Header = $true; Expected = $false }
        @{ Case = 'Citrix, review on'; Script = 'package-citrixworkspace-ltsr-x64.ps1'; Skip = $false; Header = $false; Expected = $true }
    ) {
        $prefs = [pscustomobject]@{ SkipConfigReview = $Skip }
        $row = [pscustomobject]@{ Script = $Script; ReviewBeforeStage = $Header }
        Test-StageReviewNeeded -Row $row -Prefs $prefs | Should -Be $Expected
    }

    It 'treats a row and a preference file without the new keys as review on, skip off' {
        Test-StageReviewNeeded -Row ([pscustomobject]@{ Script = 'package-dbeaver.ps1' }) -Prefs ([pscustomobject]@{}) | Should -BeTrue
    }

    It 'reads ReviewBeforeStage from a packager header' -ForEach @(
        @{ Line = '# ReviewBeforeStage: Yes'; Expected = $true }
        @{ Line = '# ReviewBeforeStage: No'; Expected = $false }
        @{ Line = '# Notes: nothing'; Expected = $false }
    ) {
        $path = Join-Path $TestDrive ('package-x-' + [guid]::NewGuid().ToString('N') + '.ps1')
        ("<#`r`n# Vendor: Contoso`r`n# App: Widget`r`n$Line`r`n#>`r`n") | Set-Content -LiteralPath $path -Encoding ASCII
        (Get-PackagerMetadata -Path $path).ReviewBeforeStage | Should -Be $Expected
    }


}

Describe 'Confirm-StageReview' {
    BeforeEach {
        $script:Asked = @(); $script:LogLines = @(); $script:Answers = @{}
        $script:Prefs = New-DefaultPrefs
        $script:window = $null
    }

    It 'asks for each application that has install options, in order, and skips the others' {
        $rows = @(
            [pscustomobject]@{ Application = 'SSMS'; Script = 'package-ssms.ps1' }
            [pscustomobject]@{ Application = '7-Zip'; Script = 'package-7zip.ps1' }
            [pscustomobject]@{ Application = 'DBeaver'; Script = 'package-dbeaver.ps1' }
        )
        $review = Confirm-StageReview -Rows $rows
        $script:Asked | Should -Be @('SSMS', 'DBeaver')
        @($review.Rows).Count | Should -Be 3
        $review.RunOptionsByApp.Count | Should -Be 0
    }

    It 'skips an application when the review is cancelled and logs it' {
        $script:Answers['DBeaver'] = @{ Continue = $false; RunOptionsJson = '' }
        $rows = @(
            [pscustomobject]@{ Application = 'SSMS'; Script = 'package-ssms.ps1' }
            [pscustomobject]@{ Application = 'DBeaver'; Script = 'package-dbeaver.ps1' }
        )
        $review = Confirm-StageReview -Rows $rows
        @($review.Rows | ForEach-Object { $_.Application }) | Should -Be @('SSMS')
        @($script:LogLines | Where-Object { $_ -match '^Skipped DBeaver: the review before staging was cancelled' }).Count | Should -Be 1
    }

    It 'records the changed values for this application only' {
        $script:Answers['SSMS'] = @{ Continue = $true; RunOptionsJson = '{"Preferences":{"SSMSInstallOptions":{"UIMode":"Passive"}},"Files":{}}' }
        $rows = @(
            [pscustomobject]@{ Application = 'SSMS'; Script = 'package-ssms.ps1' }
            [pscustomobject]@{ Application = 'DBeaver'; Script = 'package-dbeaver.ps1' }
        )
        $review = Confirm-StageReview -Rows $rows
        $review.RunOptionsByApp.Keys | Should -Be @('package-ssms')
        $review.RunOptionsByApp['package-ssms'] | Should -Match 'Passive'
        @($script:LogLines | Where-Object { $_ -match '^Install options of SSMS for this run' }).Count | Should -Be 1
    }

    It 'asks nothing when the review is skipped, and still asks for a header that forces it' {
        $script:Prefs.SkipConfigReview = $true
        $rows = @(
            [pscustomobject]@{ Application = 'SSMS'; Script = 'package-ssms.ps1'; ReviewBeforeStage = $false }
            [pscustomobject]@{ Application = 'DBeaver'; Script = 'package-dbeaver.ps1'; ReviewBeforeStage = $true }
        )
        $review = Confirm-StageReview -Rows $rows
        $script:Asked | Should -Be @('DBeaver')
        @($review.Rows).Count | Should -Be 2
    }

    It 'never changes the saved preferences' {
        $before = ConvertTo-Json -InputObject $script:Prefs -Depth 8
        $script:Answers['SSMS'] = @{ Continue = $true; RunOptionsJson = '{"Preferences":{"SSMSInstallOptions":{"UIMode":"Passive"}},"Files":{}}' }
        [void](Confirm-StageReview -Rows @([pscustomobject]@{ Application = 'SSMS'; Script = 'package-ssms.ps1' }))
        (ConvertTo-Json -InputObject $script:Prefs -Depth 8) | Should -BeExactly $before
    }
}

Describe 'Unattended runs never ask' {
    It 'calls the review only from the Stage button and the Workbench stage' {
        $calls = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Confirm-StageReview' }, $true))
        $calls.Count | Should -Be 2
        $owners = foreach ($call in $calls) {
            $parent = $call.Parent
            while ($parent -and $parent -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $parent = $parent.Parent }
            if ($parent) { $parent.Name } else { '(script)' }
        }
        @($owners | Sort-Object) | Should -Be @('(script)', 'Invoke-WorkbenchRun')
    }

    It 'shows the dialog only from Confirm-StageReview' {
        $calls = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in 'Show-StageReviewDialog', 'New-StageReviewDialog' }, $true))
        $owners = foreach ($call in $calls) {
            $parent = $call.Parent
            while ($parent -and $parent -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $parent = $parent.Parent }
            $parent.Name
        }
        @($owners | Sort-Object) | Should -Be @('Confirm-StageReview', 'Show-StageReviewDialog')
    }

    It 'does not mention the review in the batch dispatcher, the One Click run or the background pipeline' {
        foreach ($name in 'Invoke-MultiAppPipeline', 'Invoke-PackagerBatch', 'Invoke-OneClickRun') {
            $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            if ($node) { $node.Extent.Text | Should -Not -Match 'Confirm-StageReview|Show-StageReviewDialog' -Because $name }
        }
        $text = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
        $batch = [regex]::Match($text, '(?s)# Batch-mode dispatcher.*?# Data model').Value
        $batch | Should -Not -BeNullOrEmpty
        $batch | Should -Not -Match 'StageReview'
    }

    It 'passes the run options to the Stage step of the pipeline and to no other step' {
        $text = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
        ([regex]::Matches($text, '-RunOptionsJson \$\(if \(\$Ctx\.RunOptionsByApp\)')).Count | Should -Be 1
    }
}

Describe 'Get-StageReviewRunOptions' {
    BeforeEach { $script:Saved = New-DefaultPrefs }

    It 'returns nothing when no value changed' {
        Get-StageReviewRunOptions -Saved $script:Saved -Edited (Copy-Prefs $script:Saved) -Folder $TestDrive | Should -BeExactly ''
    }

    It 'counts an empty list field and a single language text as unchanged' {
        $edited = Copy-Prefs $script:Saved
        $edited.SSMSInstallOptions.Workloads = @($null)
        $edited.M365Options.Languages = 'en-us'
        Get-StageReviewRunOptions -Saved $script:Saved -Edited $edited -Folder $TestDrive | Should -BeExactly ''
    }

    It 'returns the whole changed block of a preference' {
        $edited = Copy-Prefs $script:Saved
        $edited.SSMSInstallOptions.UIMode = 'Passive'
        $run = Get-StageReviewRunOptions -Saved $script:Saved -Edited $edited -Folder $TestDrive | ConvertFrom-Json
        @($run.Preferences.PSObject.Properties.Name) | Should -Be @('SSMSInstallOptions')
        $run.Preferences.SSMSInstallOptions.UIMode | Should -Be 'Passive'
        $run.Preferences.SSMSInstallOptions.NoCache | Should -Be $script:Saved.SSMSInstallOptions.NoCache
        $script:Saved.SSMSInstallOptions.UIMode | Should -Be 'Quiet'
    }

    It 'returns the Microsoft 365 channel and deploy mode' {
        $edited = Copy-Prefs $script:Saved
        $edited.M365Channel = 'Current'
        $edited.M365DeployMode = 'Online'
        $run = Get-StageReviewRunOptions -Saved $script:Saved -Edited $edited -Folder $TestDrive | ConvertFrom-Json
        $run.Preferences.M365Channel | Should -Be 'Current'
        $run.Preferences.M365DeployMode | Should -Be 'Online'
    }

    It 'returns a changed settings file under its own name' {
        $saved = @{ cwa = [pscustomobject]@{ Store = [pscustomobject]@{ Url = '' } }; tv = [pscustomobject]@{ ApiToken = '' } }
        $edited = @{ cwa = [pscustomobject]@{ Store = [pscustomobject]@{ Url = 'https://store.example.com' } }; tv = [pscustomobject]@{ ApiToken = '' } }
        $run = Get-StageReviewRunOptions -Saved $script:Saved -Edited (Copy-Prefs $script:Saved) -SavedAux $saved -EditedAux $edited -Folder $TestDrive | ConvertFrom-Json
        @($run.Files.PSObject.Properties.Name) | Should -Be @('citrix-workspace-switches.json')
        $run.Files.'citrix-workspace-switches.json'.Store.Url | Should -Be 'https://store.example.com'
    }
}

Describe 'Run options reach the Stage child' {
    BeforeEach { $script:Child = $null }

    It 'sends nothing when the review changed nothing' {
        $null = Invoke-PackagerStage -PackagerPath (Join-Path $TestDrive 'package-ssms.ps1') -LogFolder (Join-Path $TestDrive 'logs') -M365Channel 'MonthlyEnterprise'
        $script:Child.Env | Should -BeNullOrEmpty
    }

    It 'writes the changed settings file for the run, names it in the environment and removes it afterwards' {
        $json = '{"Preferences":{"M365Channel":"Current","M365DeployMode":"Online","SSMSInstallOptions":{"UIMode":"Passive"}},"Files":{"citrix-workspace-switches.json":{"Store":{"Url":"https://store.example.com"}}}}'
        $null = Invoke-PackagerStage -PackagerPath (Join-Path $TestDrive 'package-m365apps-x64.ps1') -LogFolder (Join-Path $TestDrive 'logs') -M365Channel 'MonthlyEnterprise' -M365DeployMode 'Managed' -RunOptionsJson $json
        $runEnv = $script:Child.Env | ConvertFrom-Json
        $runEnv.Preferences.SSMSInstallOptions.UIMode | Should -Be 'Passive'
        $path = [string]$runEnv.Files.'citrix-workspace-switches.json'
        $script:Child.Files['citrix-workspace-switches.json'].Text | Should -Match 'store.example.com'
        Test-Path -LiteralPath $path | Should -BeFalse
        Test-Path -LiteralPath (Split-Path $path) | Should -BeFalse
        $arguments = @($script:Child.Arguments)
        $arguments[$arguments.IndexOf('-M365Channel') + 1] | Should -Be 'Current'
        $arguments[$arguments.IndexOf('-M365DeployMode') + 1] | Should -Be 'Online'
    }

    It 'keeps the channel and deploy mode of the call when the run does not change them' {
        $json = '{"Preferences":{"SSMSInstallOptions":{"UIMode":"Passive"}},"Files":{}}'
        $null = Invoke-PackagerStage -PackagerPath (Join-Path $TestDrive 'package-m365apps-x64.ps1') -LogFolder (Join-Path $TestDrive 'logs') -M365Channel 'MonthlyEnterprise' -M365DeployMode 'Managed' -RunOptionsJson $json
        $arguments = @($script:Child.Arguments)
        $arguments[$arguments.IndexOf('-M365Channel') + 1] | Should -Be 'MonthlyEnterprise'
    }
}

Describe 'The mirror of the packager preferences' {
    It 'holds the keys that Save-Preferences writes for the packagers' {
        $mirror = Get-PackagerPreferenceMirror -Prefs (New-DefaultPrefs) -Folder $TestDrive
        (@($mirror.Keys) -join ',') | Should -Be 'CompanyName,M365ExcludeApps,SSMSInstallOptions,DBeaverInstallOptions,AdobeReaderInstallOptions,BeyondCompareKeyFile,BeyondCompareOptions,M365Options,M365UseCustomConfig,M365CustomConfigPath,LocalSourceFolders'
    }

    It 'turns the imported file on only while the box is selected and the file exists' {
        $prefs = New-DefaultPrefs
        $prefs.M365CustomConfig.Enabled = $true
        (Get-PackagerPreferenceMirror -Prefs $prefs -Folder $TestDrive).M365UseCustomConfig | Should -BeFalse
        Set-Content -LiteralPath (Join-Path $TestDrive 'm365-configuration.user.xml') -Value '<Configuration/>' -Encoding ASCII
        $mirror = Get-PackagerPreferenceMirror -Prefs $prefs -Folder $TestDrive
        $mirror.M365UseCustomConfig | Should -BeTrue
        $mirror.M365CustomConfigPath | Should -Be (Join-Path $TestDrive 'm365-configuration.user.xml')
    }
}

Describe 'The packagers read the run copy of a settings file' {
    It 'reads <Name> through Resolve-PackagerRunFile' -ForEach @(
        @{ Name = 'package-citrixworkspace-cr.ps1' }, @{ Name = 'package-citrixworkspace-ltsr-arm64.ps1' }
        @{ Name = 'package-citrixworkspace-ltsr-x64.ps1' }, @{ Name = 'package-citrixworkspace-ltsr-x86.ps1' }
        @{ Name = 'package-teamviewerhost.ps1' }
    ) {
        $text = [IO.File]::ReadAllText((Join-Path $script:PackagersRoot $Name))
        $text | Should -Match '(?m)^\$(SwitchesFile|ConfigFile)\s+= Resolve-PackagerRunFile -Name "[a-z-]+\.json" -Default \(Join-Path \$PSScriptRoot "[a-z-]+\.json"\)'
    }
}
