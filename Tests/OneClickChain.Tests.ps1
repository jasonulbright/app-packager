#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The One Click chain (Q-124): the run and its plan always cover the version check, Check ConfigMgr, Check CVEs, stage and
    package; the plan never shows an empty version because Check Latest did not run. Also the Architecture column and the
    Options title page without a supersedence note.

.DESCRIPTION
    The window is read as source, as the other shell tests do. The CVE step runs against stubs.
#>

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Shell = Get-Content -LiteralPath (Join-Path $script:Root 'start-apppackager.ps1') -Raw
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseInput($script:Shell, [ref]$null, [ref]$null)
    foreach ($name in 'Invoke-OneClickCveStep', 'Get-CveSkipReason', 'ConvertTo-CvePairKey', 'ConvertTo-CveVersionKey', 'Set-CveCellsPair') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        $fn | Should -Not -BeNullOrEmpty -Because "function $name must exist"
        Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim().TrimStart('{').TrimEnd('}')))
    }
}

AfterAll {
    foreach ($name in 'Invoke-OneClickCveStep', 'Get-CveSkipReason', 'ConvertTo-CvePairKey', 'ConvertTo-CveVersionKey', 'Set-CveCellsPair', 'Get-ProductVulnerabilities', 'Get-VulnerabilitySummary', 'Get-CveGridCells', 'ConvertTo-VulnerabilityResultText', 'Update-PackagerHistory', 'Get-KevCatalog') {
        Remove-Item -Path ('function:global:' + $name) -ErrorAction SilentlyContinue
    }
}

Describe 'The One Click button' {
    It 'always runs the whole chain and reads no Action preference' {
        $m = [regex]::Match($script:Shell, '(?s)\$btnFullRun\.Add_Click\(\{.*?\n\}\)')
        $m.Success | Should -BeTrue
        $m.Value | Should -Match "\`$action = 'StageAndPackage'"
        $m.Value | Should -Not -Match 'AppFlow\.Action'
    }
    It 'checks the latest version of every application without a version before the plan opens' {
        $m = [regex]::Match($script:Shell, '(?s)\$btnFullRun\.Add_Click\(\{.*?\n\}\)')
        $m.Value | Should -Match 'IsNullOrWhiteSpace\(\[string\]\$_\.LatestVersion\)'
        $m.Value | Should -Match 'Invoke-MultiAppPipeline -Operation CheckLatest -Rows \$unchecked'
        $m.Value | Should -Match 'AfterDone\s+= \$openPlan'
        $m.Value | Should -Match 'Show-OneClickPlanDialog'
    }
    It 'has a done hook that the pipeline runs for the step after a run' {
        $script:Shell | Should -Match "ContainsKey\('AfterDone'\)"
    }
}

Describe 'The One Click run' {
    It 'reads the ConfigMgr site, also with Force on' {
        $script:Shell | Should -Match "-and 'ConfigMgr' -in \`$wanted -and \`$Ctx\.AdminUiFound"
    }
    It 'runs the CVE step before the stage, and the context carries the NVD settings' {
        $iCve = $script:Shell.IndexOf('Invoke-OneClickCveStep -Row $row')
        $iStage = $script:Shell.IndexOf("& `$setStep ('staging ' + `$app)")
        $iCve | Should -BeGreaterThan 0
        $iStage | Should -BeGreaterThan $iCve
        $script:Shell | Should -Match 'NvdApiKey\s+= Get-CveNvdApiKey\s+CacheTtlMinutes\s+= Get-CveCacheTtlMinutes\s+KevCatalog\s+= \$null'
    }
    It 'has no Action choice in Options, One Click' {
        $script:Shell | Should -Not -Match 'cboAction'
        $script:Shell | Should -Not -Match 'Action on update'
    }
}

Describe 'Invoke-OneClickCveStep' {
    BeforeEach {
        $script:Row = [pscustomobject]@{ Script = 'package-x.ps1'; Cpe = 'cpe:2.3:a:v:x'; CurrentVersion = '1.0'; LatestVersion = '2.0'; CvePair = ''
            CveCount = ''; CveCvss = ''; CveKev = ''; CveEpss = ''; CveChecked = ''; CveFixed = ''; CveNewest = ''; CveUpdated = '' }
        $script:State = @{ Aborted = $false }
        $script:Ctx = @{ KevCatalog = @{}; NvdApiKey = ''; CacheTtlMinutes = 60; ForceRefresh = $false; CveSession = @{} }
        function global:Get-KevCatalog { @{} }
        function global:Update-PackagerHistory { param($PackagerName, $Event, $Data) $script:History = @($PackagerName, $Event) }
        function global:Get-VulnerabilitySummary { param($Result) @{ Count = 2 } }
        function global:Get-CveGridCells { param($Summary, $Cpe) [pscustomobject]@{ Count = '2'; Cvss = '7.5'; Kev = ''; Epss = ''; Checked = 'now'; Fixed = ''; Newest = ''; Updated = ''; Text = '2 CVEs' } }
        function global:ConvertTo-VulnerabilityResultText { param($Result) '2 CVEs' }
        $script:History = $null
    }
    It 'skips a row without a CPE and says why' {
        $script:Row.Cpe = ''
        $text = Invoke-OneClickCveStep -Row $script:Row -PackagerName 'package-x' -Context $script:Ctx -State $script:State
        $text | Should -Match 'skipped: CPE missing'
        $script:Row.CveCount | Should -Be 'CPE missing'
    }
    It 'fills the CVE cells and records the summary after a lookup' {
        function global:Get-ProductVulnerabilities { param($Cpe, $CurrentVersion, $LatestVersion, $ApiKey, $CacheTtlMinutes, $KevCatalog, $Force, $VendorAdvisoryUrl) [pscustomobject]@{ Error = ''; Source = 'NVD' } }
        $text = Invoke-OneClickCveStep -Row $script:Row -PackagerName 'package-x' -Context $script:Ctx -State $script:State
        $text | Should -Match 'Check CVEs: 2 CVEs'
        $script:Row.CveCount | Should -Be '2'
        $script:Row.CveCvss | Should -Be '7.5'
        $script:History[1] | Should -Be 'VulnChecked'
    }
    It 'logs a failed lookup and never throws' {
        function global:Get-ProductVulnerabilities { throw 'NVD down' }
        $text = Invoke-OneClickCveStep -Row $script:Row -PackagerName 'package-x' -Context $script:Ctx -State $script:State
        $text | Should -Match 'Check CVEs failed'
        $script:History | Should -BeNullOrEmpty
    }
}

Describe 'The Architecture column' {
    It 'has the header Architecture and no Package Source column' {
        $script:Shell | Should -Match "Id = 'source'; Header = 'Architecture'"
        $script:Shell | Should -Not -Match "Header = 'Package Source'"
        $script:Shell | Should -Match "Id = 'scope'; Header = 'Install Scope'"
    }
}

Describe 'The title page of Options' {
    It 'has no supersedence note, column or check box' {
        $m = [regex]::Match($script:Shell, "(?s)Id = 'title'; Title = 'Application Title'.*?\n        \}\)")
        $m.Success | Should -BeTrue
        $m.Value | Should -Not -Match '(?i)supersed'
    }
}
