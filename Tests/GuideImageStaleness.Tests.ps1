#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# H-5: a guide image can be older than the last commit that changed the screen it shows. The map tools\guide-capture\GuideImages.sources.json names the
# sources of each image; tools\guide-capture\Get-StaleGuideImages.ps1 compares the last commit of the image with the last commit of its sources.
# The repository test reports the stale images as a warning and as a CI step summary and does not fail, unless GUIDE_STALE_MODE is 'fail'
# (GUIDE_STALE_MAX images may then be stale; default 0). A shallow clone cannot answer the question: the test reports that and passes.
# The fixture tests run the script against small git repositories.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:Script = Join-Path $script:Root 'tools\guide-capture\Get-StaleGuideImages.ps1'
    $script:PwshPath = (Get-Process -Id $PID).Path

    function script:Invoke-Script {
        param([string]$Repository, [string]$Map, [switch]$FailOnStale, [switch]$OnlyProblems)
        $arguments = @('-NoProfile', '-File', $script:Script, '-Root', $Repository, '-Format', 'Json')
        if ($Map) { $arguments += @('-MapPath', $Map) }
        if (-not $OnlyProblems) { $arguments += '-All' }
        if ($FailOnStale) { $arguments += '-FailOnStale' }
        $text = (& $script:PwshPath @arguments 2>&1 | Out-String)
        $exit = $LASTEXITCODE
        $report = $null
        if ($text.Trim().StartsWith('{')) { $report = $text | ConvertFrom-Json }
        return [pscustomobject]@{ Report = $report; Exit = $exit; Text = $text }
    }

    function script:Invoke-Git {
        param([string]$Repository, [string[]]$Arguments)
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { $output = & git -C $Repository -c user.name=t -c user.email=t@example.com -c commit.gpgsign=false @Arguments 2>&1 } finally { $ErrorActionPreference = $previous }
        if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed: $output" }
    }

    # Writes the files and commits them one hour after the previous commit of the fixture.
    function script:Add-Commit {
        param([string]$Repository, [string]$Message, [hashtable]$Files)
        foreach ($name in $Files.Keys) {
            $full = Join-Path $Repository $name
            New-Item -ItemType Directory -Force -Path (Split-Path $full -Parent) | Out-Null
            Set-Content -LiteralPath $full -Value $Files[$name] -Encoding Ascii
        }
        $script:Clock++
        $date = ([DateTimeOffset]'2026-11-01T00:00:00+00:00').AddHours($script:Clock).ToString('yyyy-MM-ddTHH:mm:ssK')
        $env:GIT_AUTHOR_DATE = $date; $env:GIT_COMMITTER_DATE = $date
        try { script:Invoke-Git $Repository @('add', '-A'); script:Invoke-Git $Repository @('commit', '-q', '-m', $Message) }
        finally { Remove-Item Env:GIT_AUTHOR_DATE, Env:GIT_COMMITTER_DATE -ErrorAction SilentlyContinue }
    }

    $script:MapText = '{ "ignoreSubjects": ["^Release "], "shared": ["Lib/Shared/"], "rules": [ { "ids": ["alpha-*"], "sources": ["Main.xaml", "Lib/Alpha*"] }, { "ids": ["beta"], "sources": ["Beta.ps1"] } ] }'

    function script:New-Fixture {
        $path = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $path | Out-Null
        $script:Clock = 0
        script:Invoke-Git $path @('init', '-q', '-b', 'main')
        script:Add-Commit $path 'sources' @{ 'Main.xaml' = 'm'; 'Lib/Alpha.ps1' = 'a'; 'Beta.ps1' = 'b'; 'Lib/Shared/Common.ps1' = 'c'; 'map.json' = $script:MapText }
        return $path
    }

    function script:Get-Row { param($Result, [string]$Id) @($Result.Report.Images | Where-Object { $_.Id -eq $Id })[0] }
}

Describe 'The staleness script compares the last commit of an image with the last commit of its sources' {
    It 'reports an image that is newer than its sources as fresh' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        $result = script:Invoke-Script $repo (Join-Path $repo 'map.json')
        @($result.Report.Images | ForEach-Object Status) | Should -Be @('Fresh', 'Fresh')
    }

    It 'reports an image as stale after a later change to one of its sources, and names the change' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        script:Add-Commit $repo 'change the alpha view' @{ 'Main.xaml' = 'm2' }
        $result = script:Invoke-Script $repo (Join-Path $repo 'map.json')
        (script:Get-Row $result 'alpha-one').Status | Should -Be 'Stale'
        (script:Get-Row $result 'alpha-one').Reason | Should -Be 'change the alpha view'
        @((script:Get-Row $result 'alpha-one').ChangedFiles) | Should -Be @('Main.xaml')
        (script:Get-Row $result 'beta').Status | Should -Be 'Fresh'
    }

    It 'treats an image and its sources in the same commit as fresh, and a retake as fresh again' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'change and retake' @{ 'Main.xaml' = 'm2'; 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        (script:Get-Row (script:Invoke-Script $repo (Join-Path $repo 'map.json')) 'alpha-one').Status | Should -Be 'Fresh'
        script:Add-Commit $repo 'change again' @{ 'Lib/Alpha.ps1' = 'a2' }
        (script:Get-Row (script:Invoke-Script $repo (Join-Path $repo 'map.json')) 'alpha-one').Status | Should -Be 'Stale'
        script:Add-Commit $repo 'retake' @{ 'docs/images/alpha-one.png' = '2' }
        (script:Get-Row (script:Invoke-Script $repo (Join-Path $repo 'map.json')) 'alpha-one').Status | Should -Be 'Fresh'
    }

    It 'makes every image stale after a change to a shared source, and not after a change to another screen' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        script:Add-Commit $repo 'another screen' @{ 'Other.ps1' = 'o' }
        @((script:Invoke-Script $repo (Join-Path $repo 'map.json')).Report.Images | ForEach-Object Status) | Should -Be @('Fresh', 'Fresh')
        script:Add-Commit $repo 'shared change' @{ 'Lib/Shared/Common.ps1' = 'c2' }
        @((script:Invoke-Script $repo (Join-Path $repo 'map.json')).Report.Images | ForEach-Object Status) | Should -Be @('Stale', 'Stale')
    }

    It 'ignores a commit whose subject matches ignoreSubjects' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        script:Add-Commit $repo 'Release 2026.11.1.0001' @{ 'Beta.ps1' = 'b2' }
        (script:Get-Row (script:Invoke-Script $repo (Join-Path $repo 'map.json')) 'beta').Status | Should -Be 'Fresh'
    }

    It 'reports an image with no rule as Unmapped and an uncommitted image as Unknown' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/no-rule.png' = '1' }
        Set-Content -LiteralPath (Join-Path $repo 'docs/images/beta.png') -Value 'x' -Encoding Ascii
        $result = script:Invoke-Script $repo (Join-Path $repo 'map.json')
        (script:Get-Row $result 'no-rule').Status | Should -Be 'Unmapped'
        (script:Get-Row $result 'beta').Status | Should -Be 'Unknown'
    }

    It 'reports Unknown, not Stale, in a shallow clone' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        script:Add-Commit $repo 'change' @{ 'Beta.ps1' = 'b2' }
        $clone = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        script:Invoke-Git $TestDrive @('clone', '-q', '--depth', '1', ([uri]$repo).AbsoluteUri, $clone)
        $result = script:Invoke-Script $clone (Join-Path $clone 'map.json')
        @($result.Report.Images | ForEach-Object Status | Sort-Object -Unique) | Should -Be @('Unknown')
        (script:Get-Row $result 'beta').Reason | Should -Match 'Shallow'
    }

    It 'exits with 1 under -FailOnStale only for a stale or unmapped image' {
        $repo = script:New-Fixture
        script:Add-Commit $repo 'images' @{ 'docs/images/alpha-one.png' = '1'; 'docs/images/beta.png' = '1' }
        (script:Invoke-Script $repo (Join-Path $repo 'map.json') -FailOnStale).Exit | Should -Be 0
        script:Add-Commit $repo 'change' @{ 'Beta.ps1' = 'b2' }
        (script:Invoke-Script $repo (Join-Path $repo 'map.json') -FailOnStale).Exit | Should -Be 1
        (script:Invoke-Script $repo (Join-Path $repo 'map.json')).Exit | Should -Be 0
    }
}

Describe 'The guide images of this repository' {
    It 'have a rule in the map and the stale ones are reported as a warning' {
        $result = script:Invoke-Script $script:Root -OnlyProblems
        $result.Report | Should -Not -BeNullOrEmpty -Because $result.Text
        $stale = @($result.Report.Images | Where-Object { $_.Status -eq 'Stale' })
        $unknown = @($result.Report.Images | Where-Object { $_.Status -eq 'Unknown' })
        $unmapped = @($result.Report.Images | Where-Object { $_.Status -eq 'Unmapped' })

        Write-Host $result.Report.Summary
        foreach ($image in $stale) { Write-Warning ('STALE {0}: image {1}, source {2}, {3}' -f $image.Id, $image.ImageDate, $image.SourceDate, $image.Reason) }
        foreach ($image in $unknown) { Write-Warning ('UNKNOWN {0}: {1}' -f $image.Id, $image.Reason) }
        if ($env:GITHUB_STEP_SUMMARY) {
            $lines = @('### Guide image staleness', '', $result.Report.Summary)
            if ($stale.Count + $unknown.Count -gt 0) {
                $lines += @('', '| Image | Status | Image date | Source date | Last source change |', '|---|---|---|---|---|')
                foreach ($image in @($stale) + @($unknown)) { $lines += ('| {0} | {1} | {2} | {3} | {4} |' -f $image.Id, $image.Status, $image.ImageDate, $image.SourceDate, ([string]$image.Reason).Replace('|', '/')) }
            }
            Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Value ($lines + '')
        }

        $unmapped.Count | Should -Be 0 -Because ('tools\guide-capture\GuideImages.sources.json has no rule for: ' + (($unmapped | ForEach-Object Id) -join ', '))
        if ($env:GUIDE_STALE_MODE -eq 'fail') {
            $allowed = 0
            if ($env:GUIDE_STALE_MAX) { $allowed = [int]$env:GUIDE_STALE_MAX }
            $unknown.Count | Should -Be 0 -Because 'the history is incomplete, so staleness cannot be checked (a shallow clone)'
            $stale.Count | Should -BeLessOrEqual $allowed -Because ('these guide images are older than the last change to their screen: ' + (($stale | ForEach-Object Id) -join ', '))
        }
    }
}
