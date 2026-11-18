#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The supersedence-only pass of One Click: the pass for one application, and how the run uses it.

.DESCRIPTION
    The behavior follows the specification of App Packager Desktop (docs/specs/supersedence.md, section "Failure and retry"):
    a publish whose supersedence step failed stays published with a warning; the version stays pending; the next run does
    not publish it again, it runs the lookup and the supersedence step for that application and nothing else. The pass
    runs against mocked ConfigMgr and stage functions. The run loop is read as source, as the other One Click tests do.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerOneClick.psd1') -Force

    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-OneClickSupersedencePass' }, $false)
    if (-not $fn) { throw 'function Invoke-OneClickSupersedencePass not found' }
    $script:PassText = $fn.Extent.Text
    . ([scriptblock]::Create($script:PassText))
    $errorText = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-UserErrorText' }, $false)
    . ([scriptblock]::Create($errorText.Extent.Text))

    # The function of the shell that the pass calls; the modules give the others.
    function Find-NewestStageManifestForPackager { param([string]$PackagerPath, [string]$DownloadRoot) $null }

    function New-PlanResult {
        param([string]$Mode = 'Replace', [string]$Marker = 'APD-Line:catalog:package-a|default|x64', [string]$Refusal = $null, [string[]]$Log = @('Supersedence mode           : Replace (site default)'))
        [pscustomobject]@{ ConfigMgrManifest = $null; IntuneManifest = $null; Marker = $Marker; Refusal = $Refusal; Log = @($Log); Mode = $Mode; ConfigMgrRule = 'Minimum' }
    }
    function Set-PassMocks {
        param($Plan = (New-PlanResult), [bool]$Connected = $true, $Step = $null)
        if (-not $Step) { $Step = [pscustomobject]@{ Status = 'Applied'; Message = 'App A 2.0 supersedes App A 1.0'; Warnings = @(); Failure = $null } }
        $script:StepResult = $Step
        $script:PlanResult = $Plan
        $script:Connected = $Connected
        Mock Find-NewestStageManifestForPackager { Join-Path $TestDrive 'stage-manifest.json' }
        Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like '*stage-manifest.json' }
        Mock Read-StageManifest { [pscustomobject]@{ AppName = 'App A'; SoftwareVersion = '2.0'; ApplicationId = 'catalog:package-a' } }
        Mock Get-SupersedenceSettings { New-SupersedenceSettings }
        Mock Get-SupersedenceCommandLine { $null }
        Mock Get-SupersedencePlan { $script:PlanResult }
        Mock Connect-CMSite { $script:Connected }
        Mock Set-Location { }
        Mock Invoke-SupersedenceStep { $script:StepResult }
    }
    function Invoke-TestPass {
        Invoke-OneClickSupersedencePass -ApplicationName 'App A 2.0' -Version '2.0' -PackagerPath 'C:\x\package-a.ps1' -DownloadRoot 'C:\dl' -SiteCode 'MCM' -ProviderMachineName 'cm01'
    }
}

Describe 'The supersedence-only pass for one application' {
    It 'runs the supersedence step for the application that is on the site and returns its result' {
        Set-PassMocks
        $r = Invoke-TestPass
        $r.Status | Should -Be 'Applied'
        $r.Message | Should -Be 'App A 2.0 supersedes App A 1.0'
        Should -Invoke Invoke-SupersedenceStep -Times 1 -Exactly -ParameterFilter {
            $NewApplicationName -eq 'App A 2.0' -and $NewVersion -eq '2.0' -and $Marker -eq 'APD-Line:catalog:package-a|default|x64' -and $Mode -eq 'Replace' -and $RequirePublished
        }
    }

    It 'passes the title of the line of the packager when the manifest has no application identifier' {
        Set-PassMocks
        [void](Invoke-TestPass)
        Should -Invoke Get-SupersedencePlan -Times 1 -Exactly -ParameterFilter { $FallbackApplicationId -eq 'catalog:package-a' }
    }

    It 'resolves the mode like a publish and returns the lines of the plan' {
        Set-PassMocks -Plan (New-PlanResult -Mode 'ReplaceAndUninstall' -Log @('Version rule (ConfigMgr)    : Exact (title override)', 'Supersedence mode           : Replace and uninstall (title override)'))
        $r = Invoke-TestPass
        $r.Lines.Count | Should -Be 2
        Should -Invoke Invoke-SupersedenceStep -Times 1 -Exactly -ParameterFilter { $Mode -eq 'ReplaceAndUninstall' }
    }

    It 'does nothing when the mode is Off and asks nothing of the site' {
        Set-PassMocks -Plan (New-PlanResult -Mode 'Off')
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotRun'
        Should -Invoke Connect-CMSite -Times 0 -Exactly
        Should -Invoke Invoke-SupersedenceStep -Times 0 -Exactly
    }

    It 'does nothing for a title without an identity marker' {
        Set-PassMocks -Plan (New-PlanResult -Marker $null)
        (Invoke-TestPass).Status | Should -Be 'NotRun'
        Should -Invoke Invoke-SupersedenceStep -Times 0 -Exactly
    }

    It 'reports a refused plan, such as the blocked combination, and writes nothing' {
        Set-PassMocks -Plan (New-PlanResult -Refusal 'Replace and uninstall needs the version rule Exact.')
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotApplied'
        $r.Message | Should -Match 'needs the version rule Exact'
        Should -Invoke Connect-CMSite -Times 0 -Exactly
        Should -Invoke Invoke-SupersedenceStep -Times 0 -Exactly
    }

    It 'reports a missing stage manifest and asks nothing of the site' {
        Set-PassMocks
        Mock Find-NewestStageManifestForPackager { $null }
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotApplied'
        $r.Message | Should -Match 'stage-manifest.json not found'
        Should -Invoke Invoke-SupersedenceStep -Times 0 -Exactly
    }

    It 'reports a site that cannot be reached' {
        Set-PassMocks -Connected $false
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotApplied'
        $r.Message | Should -Match 'site MCM failed'
        Should -Invoke Invoke-SupersedenceStep -Times 0 -Exactly
    }

    It 'passes the result of a step that is not applied, with the reason, and keeps the warnings' {
        Set-PassMocks -Step ([pscustomobject]@{ Status = 'NotApplied'; Message = "The application 'App A 2.0' is not on the site. Publish it first."; Warnings = @('w1'); Failure = 'x' })
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotApplied'
        $r.Message | Should -Match 'is not on the site'
        $r.Warnings | Should -Be @('w1')
    }

    It 'passes the result of a line without an older version' {
        Set-PassMocks -Step ([pscustomobject]@{ Status = 'NoOlderVersion'; Message = 'No older version found'; Warnings = @(); Failure = $null })
        (Invoke-TestPass).Status | Should -Be 'NoOlderVersion'
    }

    It 'never throws and goes back to the location it started in' {
        Set-PassMocks
        Mock Invoke-SupersedenceStep { throw [System.InvalidOperationException]::new('provider gone') }
        $r = Invoke-TestPass
        $r.Status | Should -Be 'NotApplied'
        # An error that the tool did not write itself is raw exception text: the window gets the fallback sentence.
        $r.Message | Should -Be 'the supersedence step failed. See the log file.'
        Should -Invoke Set-Location -Times 1 -Exactly -ParameterFilter { $Path -eq 'MCM:' }
        Should -Invoke Set-Location -Times 2 -Exactly
    }

    It 'stages nothing and publishes nothing' {
        $script:PassText | Should -Not -Match 'Invoke-PackagerStage|Invoke-PackagerPackage|Invoke-PackagerWsusPostStep|Invoke-PackagerIntunePublish|New-CMApplication|Set-CMApplication\b'
        $script:PassText | Should -Match '-RequirePublished'
    }
}

Describe 'How the One Click run uses the pass' {
    BeforeAll {
        $branch = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] -and $n.Extent.Text -match "'OneClick'" -and $n.Extent.Text -match "'FullRun'" }, $true)
        $text = $(if ($branch) { $branch.Extent.Text } else { $script:Ast.Extent.Text })
        $start = $text.IndexOf("'OneClick' {")
        $end = $text.IndexOf("'FullRun' {", $start)
        $script:OneClick = $text.Substring($start, $end - $start)
    }

    It 'runs the pass only for a version that the history holds as published, pending, unforced and of the same build' {
        $script:OneClick | Should -Match '(?s)-not \$Ctx\.ForceFlag -and -not \$buildStale -and \$sameBuild -and \$rec -and \$rec\.Version -eq \[string\]\$latest\) \{\s*if \(\$d -eq ''ConfigMgr'' -and \$rec\.SupersedencePending\) \{\s*\$supersedencePass = \$rec'
    }

    It 'does not publish the version again and runs before the stage and the other destinations' {
        $pass = $script:OneClick.IndexOf('Invoke-OneClickSupersedencePass')
        $pass | Should -BeGreaterThan $script:OneClick.IndexOf('$supersedencePass = $rec')
        $pass | Should -BeLessThan $script:OneClick.IndexOf('# 3. Stage')
        $script:OneClick | Should -Match '(?s)\$supersedencePass = \$rec\s*\$skipReasons \+= [^\r\n]*supersedence pending[^\r\n]*\s*continue'
    }

    It 'ends the pending state only when the pass applied or found no older version' {
        $script:OneClick | Should -Match "(?s)\{ \`$_ -in @\('Applied', 'NoOlderVersion'\) \} \{[^}]*Clear-OneClickSupersedencePending"
        ([regex]::Matches($script:OneClick, 'Clear-OneClickSupersedencePending')).Count | Should -Be 1
    }

    It 'records a publish whose supersedence step failed as published with the pending state, and not as failed' {
        $script:OneClick | Should -Match 'Get-OneClickSupersedenceFailure -Text \$pkg\.StdOut -ExitCode \$pkg\.ExitCode'
        $script:OneClick | Should -Match '-SupersedenceNotApplied:\(\[bool\]\$supersedenceFailure\) -SupersedenceReason'
        $script:OneClick | Should -Match "published \{0\}, supersedence not applied"
        $script:OneClick | Should -Match "Published, supersedence not applied: ' \+ \`$supersedenceFailure"
    }

    It 'finishes a row that only ran the pass with the outcome Superseded, counted on its own' {
        $script:OneClick | Should -Match "\`$passOutcome = 'Superseded'"
        $script:OneClick | Should -Match "\`$counts\['Superseded'\]\+\+"
    }

    It 'keeps the pass out of the other runs' {
        $text = $script:Ast.Extent.Text
        ([regex]::Matches($text, 'Invoke-OneClickSupersedencePass -ApplicationName')).Count | Should -Be 1
    }
}
