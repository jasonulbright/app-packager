#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Tableau Prep Builder: the download path and the point release.

.DESCRIPTION
    The Prep release page names the release train only (2026.3); the Desktop and Reader page lists the point release
    (2026.3.1). The packager takes the patch from the pages and settles it with a probe of the download address. A probe that
    cannot tell (403) never changes the version and never reads as "does not exist". The fixtures are small copies of the
    shape of both pages. No request goes to Tableau: the probe is a script block.
#>

BeforeAll {
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-tableauprep.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($assign in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Parent -is [System.Management.Automation.Language.NamedBlockAst] }, $false)) {
        if ($assign.Left.Extent.Text -in @('$DownloadUrlTemplate', '$ReleasePageUrl', '$TrainPageUrl')) { . ([scriptblock]::Create($assign.Extent.Text)) }
    }
    foreach ($name in 'Get-TableauTrainPatches', 'ConvertTo-TableauProbeResult', 'Resolve-TableauPrepVersion') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:Text = Get-Content -LiteralPath $script:PackagerPath -Raw

    # The Prep page: the train only, with older trains below it.
    $script:PrepPage = '<h1>Tableau Prep release notes</h1><p>Current release: 2026.3</p><ul><li><a href="/p/2026.2">2026.2</a></li><li>2026.1</li><li>2025.3</li></ul>'
    # The Desktop and Reader page: the point releases.
    $script:TrainPage = '<h1>Release notes</h1><ul><li>2026.3.1</li><li>2026.2.3</li><li>2026.2.2</li><li>2026.1.4</li></ul>'

    function New-Probe {
        # A probe from a table: version -> Exists, Missing or Unknown; any version that the table lacks is Missing.
        param([hashtable]$Table, [string]$Default = 'Missing')
        $table = $Table; $default = $Default
        return { param($v) $global:Asked += $v; if ($table.ContainsKey($v)) { $table[$v] } else { $default } }.GetNewClosure()
    }
}

Describe 'Tableau Prep download address' {
    It 'uses the esdalt path with the dotted release in the folder and the hyphenated one in the file name' {
        ($DownloadUrlTemplate -f '2026-3-1', '2026.3.1') | Should -Be 'https://downloads.tableau.com/esdalt/tableau_prep/2026.3.1/TableauPrep-2026-3-1.exe'
    }

    It 'no longer builds the tssoftware address that no stage ever proved' {
        $script:Text | Should -Not -Match 'downloads\.tableau\.com/tssoftware/TableauPrep'
    }

    It 'reads the Prep page and the Desktop and Reader page' {
        $ReleasePageUrl | Should -Be 'https://www.tableau.com/support/releases/prep'
        $TrainPageUrl | Should -Be 'https://www.tableau.com/support/releases'
    }
}

Describe 'Tableau Prep point release' {
    BeforeEach { $global:Asked = @() }

    It 'lists the point releases of one train only' {
        Get-TableauTrainPatches -Html $script:TrainPage -Train '2026.2' | Should -Be @(2, 3)
        Get-TableauTrainPatches -Html $script:PrepPage -Train '2026.3' | Should -Be @(0)
        Get-TableauTrainPatches -Html 'nothing' -Train '2026.3' | Should -BeNullOrEmpty
    }

    It 'takes the patch of the Desktop page for the train of the Prep page' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml $script:TrainPage -Probe (New-Probe @{ '2026.3.1' = 'Exists' })
        $r.Version | Should -Be '2026.3.1'
        $r.State | Should -Be 'Exists'
    }

    It 'never takes a patch of an older train that the Desktop page lists' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml '<li>2026.2.3</li>' -Probe (New-Probe @{ '2026.3.0' = 'Exists' })
        $r.Version | Should -Be '2026.3.0'
    }

    It 'steps down when the probe says that the listed patch does not exist' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml $script:TrainPage -Probe (New-Probe @{ '2026.3.1' = 'Missing'; '2026.3.0' = 'Exists' })
        $r.Version | Should -Be '2026.3.0'
        $r.Notes -join "`n" | Should -Match 'does not exist on the download host; trying 2026\.3\.0'
    }

    It 'finds a point release that no page lists yet by stepping up from the known patch' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml '' -Probe (New-Probe @{ '2026.3.0' = 'Exists'; '2026.3.1' = 'Exists'; '2026.3.2' = 'Exists' })
        $r.Version | Should -Be '2026.3.2'
        $global:Asked | Should -Be @('2026.3.0', '2026.3.1', '2026.3.2', '2026.3.3')
    }

    It 'stops stepping up at the first patch that does not exist' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml '' -Probe (New-Probe @{ '2026.3.0' = 'Exists'; '2026.3.1' = 'Exists'; '2026.3.3' = 'Exists' })
        $r.Version | Should -Be '2026.3.1'
    }

    It 'keeps the version of the pages when the host cannot tell (403), and says so' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml $script:TrainPage -Probe (New-Probe @{} -Default 'Unknown')
        $r.Version | Should -Be '2026.3.1'
        $r.State | Should -Be 'Unknown'
        $r.Notes -join "`n" | Should -Match 'cannot tell whether Prep 2026\.3\.1 exists'
        $global:Asked | Should -Be @('2026.3.1')
    }

    It 'never steps down and never reads a 403 as "does not exist"' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml $script:TrainPage -Probe (New-Probe @{ '2026.3.1' = 'Unknown'; '2026.3.0' = 'Exists' })
        $r.Version | Should -Be '2026.3.1'
    }

    It 'with no Desktop page and a host that cannot tell, stays at the train and says it is not confirmed' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml '' -Probe (New-Probe @{} -Default 'Unknown')
        $r.Version | Should -Be '2026.3.0'
        $r.State | Should -Be 'Unknown'
        $r.Notes -join "`n" | Should -Match 'no page lists a point release'
    }

    It 'reports Missing when no patch of the train exists' {
        $r = Resolve-TableauPrepVersion -PrepHtml $script:PrepPage -TrainHtml $script:TrainPage -Probe (New-Probe @{})
        $r.State | Should -Be 'Missing'
        $r.Notes -join "`n" | Should -Match 'No Prep installer of the train 2026\.3 exists'
    }

    It 'returns nothing for a page without a release' {
        Resolve-TableauPrepVersion -PrepHtml 'no releases' -TrainHtml '' -Probe (New-Probe @{}) | Should -BeNullOrEmpty
    }

    It 'takes the point release of the Prep page itself when it lists one' {
        $r = Resolve-TableauPrepVersion -PrepHtml '<p>2026.3.2</p>' -TrainHtml '<p>2026.3.1</p>' -Probe (New-Probe @{ '2026.3.2' = 'Exists' })
        $r.Version | Should -Be '2026.3.2'
    }
}

Describe 'Tableau installer probe answer' {
    It 'reads a file answer as Exists' {
        ConvertTo-TableauProbeResult -Answer '200|application/octet-stream|https://downloads.tableau.com/esdalt/tableau_prep/2026.3.1/TableauPrep-2026-3-1.exe' | Should -Be 'Exists'
        ConvertTo-TableauProbeResult -Answer '206|application/x-msdownload|https://downloads.tableau.com/x.exe' | Should -Be 'Exists'
    }

    It 'reads 404 and 410 as Missing' {
        ConvertTo-TableauProbeResult -Answer '404|text/html|https://downloads.tableau.com/x.exe' | Should -Be 'Missing'
        ConvertTo-TableauProbeResult -Answer '410||https://downloads.tableau.com/x.exe' | Should -Be 'Missing'
    }

    It 'reads 403, a server error, no answer and a sign-in page as Unknown, never as Missing' {
        ConvertTo-TableauProbeResult -Answer '403|text/html|https://downloads.tableau.com/x.exe' | Should -Be 'Unknown'
        ConvertTo-TableauProbeResult -Answer '503||https://downloads.tableau.com/x.exe' | Should -Be 'Unknown'
        ConvertTo-TableauProbeResult -Answer '000||' | Should -Be 'Unknown'
        ConvertTo-TableauProbeResult -Answer '' | Should -Be 'Unknown'
        ConvertTo-TableauProbeResult -Answer '200|text/html; charset=utf-8|https://downloads.tableau.com/x.exe' | Should -Be 'Unknown'
        ConvertTo-TableauProbeResult -Answer '200|application/octet-stream|https://sso.example.test/login' | Should -Be 'Unknown'
    }
}

Describe 'Tableau Prep download' {
    It 'downloads through the resumable download with the browser headers' {
        $script:Text | Should -Match 'Invoke-ResumableDownload -Url \$release\.DownloadUrl -OutFile \$localExe -ExtraCurlArgs \(Get-TableauBrowserCurlArgs\)'
    }

    It 'says in the help what a 403 means and that the download resumes' {
        $script:Text | Should -Match 'a 403 never counts as "does not exist"'
        $script:Text | Should -Match 'continues where\s+it stopped'
    }
}
