#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5.x guard: the user-facing text of this repository follows the ConfigMgr lexicon
    (lexicon\configmgr-lexicon.json). Windows PowerShell 5.1 and PowerShell 7.

.DESCRIPTION
    Scans README.md, docs\**\*.md, the [Unreleased] part of CHANGELOG.md, XAML text attributes, message,
    dialog, status, log, warning and error strings and comment-based help in .ps1 and .psm1 files.
    It skips code identifiers, cmdlet and class names, URLs, code blocks and inline code.

    Two modes:
      Enforce      (default in the suite repository): a hit with severity error or warning fails the test.
      Report-only  (set SUITE_LEXICON_REPORT_ONLY=1, or no lexicon folder with the file 'enforce.marker'):
                   the test lists the hits and passes. Every other repository runs in this mode until its
                   fix pull request lands. Then the repository deletes the line that sets report-only.

    The lexicon is found in 'lexicon' (the suite repository) or in 'Lib\Lexicon' (a vendored copy).
    The allow-list is lexicon-allowlist.json next to the lexicon. Each entry needs a reason.
#>

BeforeDiscovery {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
}

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:LexDir = $null
    foreach ($candidate in 'lexicon', 'Lib\Lexicon') {
        $p = Join-Path $script:Root $candidate
        if (Test-Path -LiteralPath (Join-Path $p 'configmgr-lexicon.json')) { $script:LexDir = $p; break }
    }
    if ($script:LexDir) { Import-Module (Join-Path $script:LexDir 'Lexicon.psm1') -Force }
    $script:ReportOnly = ($env:SUITE_LEXICON_REPORT_ONLY -eq '1') -or ($script:LexDir -and -not (Test-Path -LiteralPath (Join-Path $script:LexDir 'enforce.marker')))
}

Describe 'ConfigMgr lexicon data' {
    It 'is found in this repository' {
        $script:LexDir | Should -Not -BeNullOrEmpty -Because 'lexicon\configmgr-lexicon.json or Lib\Lexicon\configmgr-lexicon.json must exist'
    }

    It 'has the documented fields on every entry' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        $names = 'id', 'category', 'scope', 'kind', 'preferred', 'flagged', 'allowWhen', 'context', 'replacement', 'source', 'status', 'severity', 'notes'
        foreach ($e in $lex.entries) {
            foreach ($n in $names) { ($e.PSObject.Properties.Name -contains $n) | Should -BeTrue -Because ('{0} has no field {1}' -f $e.id, $n) }
            $e.status | Should -BeIn 'verified', 'unverified' -Because ('{0} status' -f $e.id)
            $e.severity | Should -BeIn 'error', 'warning', 'info' -Because ('{0} severity' -f $e.id)
            @($e.source).Count | Should -BeGreaterThan 0 -Because ('{0} needs a source' -f $e.id)
        }
        @($lex.entries.id | Select-Object -Unique).Count | Should -Be @($lex.entries).Count -Because 'ids are unique'
    }

    It 'has a source that is a Microsoft Learn URL or a code or owner reference for every entry' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        foreach ($e in $lex.entries) {
            foreach ($s in @($e.source)) {
                $s | Should -Match '^(https://learn\.microsoft\.com/|code:|vault |owner )' -Because ('{0} source {1}' -f $e.id, $s)
            }
        }
    }

    It 'has an allow-list in which every entry carries a reason' {
        $allow = @(Import-LexiconAllowList -Path (Join-Path $script:LexDir 'lexicon-allowlist.json'))
        foreach ($a in $allow) { [string]$a.reason | Should -Not -BeNullOrEmpty -Because ('allow entry for {0}' -f $a.file) }
    }

    It 'finds a seeded bad phrase (the checker works)' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        $item = [pscustomobject]@{ File = 'x.md'; Line = 1; Kind = 'markdown'; Text = 'Read the facts of one device from the site.' }
        $hit = @(Find-LexiconHit -Item $item -Lexicon $lex)
        ($hit.RuleId -contains 'LX-001') | Should -BeTrue
        ($hit.RuleId -contains 'LX-002') | Should -BeTrue
    }

    It 'lets a pull-distribution point pass and flags pull the data' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        $ok = [pscustomobject]@{ File = 'x.md'; Line = 1; Kind = 'markdown'; Text = 'A pull-distribution point gets content from a source distribution point.' }
        $bad = [pscustomobject]@{ File = 'x.md'; Line = 2; Kind = 'markdown'; Text = 'The tool can pull the data of the device.' }
        @(Find-LexiconHit -Item $ok -Lexicon $lex | Where-Object RuleId -eq 'LX-008').Count | Should -Be 0
        @(Find-LexiconHit -Item $bad -Lexicon $lex | Where-Object RuleId -eq 'LX-008').Count | Should -Be 1
    }
}

Describe 'Text extraction' {
    It 'finds a sentence kept in a variable, a hashtable value and a method argument' {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('lexext-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tmp | Out-Null
        try {
            $code = @(
                '$script:StatusText = "Ready. Click Scan."',
                '$h = @{ Status = "Ready. Click Refresh now." }',
                '$sb.Append("Ready. Click the button.") | Out-Null',
                '$x = Get-Thing -Name "SomeName"',
                '$sql = "SELECT Name FROM v_R_System WHERE ResourceID = 1."'
            )
            Set-Content -LiteralPath (Join-Path $tmp 'a.ps1') -Value $code -Encoding ascii
            $items = @(Get-LexiconTextItem -Root $tmp)
            @($items | Where-Object { $_.Text -like '*Click Scan*' }).Count | Should -Be 1
            @($items | Where-Object { $_.Text -like '*Click Refresh*' }).Count | Should -Be 1
            @($items | Where-Object { $_.Text -like '*Click the button*' }).Count | Should -Be 1
            @($items | Where-Object { $_.Text -like '*SELECT*' -or $_.Text -like '*SomeName*' }).Count | Should -Be 0
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Module manifests' {
    It 'reads the Description of a .psd1 and flags a seeded bad phrase in it (Q-82)' {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('lexpsd-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path (Join-Path $tmp 'Module') | Out-Null
        try {
            $manifest = @(
                '@{',
                "    RootModule        = 'Seed.psm1'",
                "    Description       = 'Reads the site facts and the O''Brien rules of one device.'",
                '    Author            = "Facts Author"',
                '}'
            )
            Set-Content -LiteralPath (Join-Path $tmp 'Module\Seed.psd1') -Value $manifest -Encoding ascii
            $items = @(Get-LexiconTextItem -Root $tmp)
            $m = @($items | Where-Object { $_.Kind -eq 'manifest' })
            $m.Count | Should -Be 1
            $m[0].File | Should -Be 'Module/Seed.psd1'
            $m[0].Line | Should -Be 3
            $m[0].Text | Should -Be "Reads the site facts and the O'Brien rules of one device."
            $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
            @(Find-LexiconHit -Item $m[0] -Lexicon $lex).Count | Should -BeGreaterThan 0
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }

    It 'does not read the manifests under Lib, Tests or Catalog-less vendored folders' {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('lexpsd-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path (Join-Path $tmp 'Lib\Vendored') | Out-Null
        try {
            Set-Content -LiteralPath (Join-Path $tmp 'Lib\Vendored\V.psd1') -Value @('@{', "    Description = 'Site facts.'", '}') -Encoding ascii
            @(Get-LexiconTextItem -Root $tmp | Where-Object { $_.Kind -eq 'manifest' }).Count | Should -Be 0
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Banned names' {
    It 'finds a seeded banned name in a tracked-file line and keeps format detection apart' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('lexban-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tmp | Out-Null
        try {
            $name = 'Chocol' + 'atey'
            Set-Content -LiteralPath (Join-Path $tmp 'a.md') -Value ('Install it with ' + $name + ' today.'), 'The tool reads a .nupkg file.', 'Nothing here.' -Encoding ascii
            $hits = @(Get-BannedNameHit -Root $tmp -Lexicon $lex)
            @($hits | Where-Object { $_.RuleId -eq 'LX-040' -and $_.Line -eq 1 }).Count | Should -Be 1
            @($hits | Where-Object { $_.RuleId -eq 'LX-041' -and $_.Line -eq 2 }).Count | Should -Be 1
            @($hits | Where-Object { $_.Line -eq 3 }).Count | Should -Be 0
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }

    It 'has no banned name in a tracked file (report-only until enforce.marker exists)' {
        $lex = Import-LexiconData -Path (Join-Path $script:LexDir 'configmgr-lexicon.json')
        $allow = @(Import-LexiconAllowList -Path (Join-Path $script:LexDir 'lexicon-allowlist.json'))
        $hits = @(Get-BannedNameHit -Root $script:Root -Lexicon $lex -AllowList $allow | Where-Object RuleId -eq 'LX-040')
        $lines = @($hits | Select-Object -First 40 | ForEach-Object { '{0}:{1} [{2}] {3}' -f $_.File, $_.Line, $_.RuleId, $_.Text })
        if ($script:ReportOnly) {
            Write-Host ('LEXICON REPORT-ONLY: {0} banned-name hits' -f $hits.Count)
            $lines | ForEach-Object { Write-Host $_ }
            Set-ItResult -Skipped -Because ('report-only mode: {0} banned-name hits' -f $hits.Count) | Out-Null
        }
        else {
            $lines | Should -BeNullOrEmpty -Because 'banned names must not appear in a tracked file'
        }
    }
}

Describe 'User-facing text follows the ConfigMgr lexicon' {
    It 'has no error or warning hit (<mode>)' -ForEach @(@{ mode = 'see SUITE_LEXICON_REPORT_ONLY' }) {
        $hits = @(Invoke-LexiconCheck -Root $script:Root -LexiconPath (Join-Path $script:LexDir 'configmgr-lexicon.json') -AllowListPath (Join-Path $script:LexDir 'lexicon-allowlist.json') -RepoType PowerShell)
        $blocking = @($hits | Where-Object { $_.Severity -in 'error', 'warning' })
        $lines = @($blocking | Sort-Object RuleId, File, Line | Select-Object -First 40 | ForEach-Object { '{0}:{1} [{2}] "{3}" -> {4}' -f $_.File, $_.Line, $_.RuleId, $_.Phrase, $_.Replacement })
        if ($script:ReportOnly) {
            Write-Host ('LEXICON REPORT-ONLY: {0} blocking hits, {1} info hits' -f $blocking.Count, (@($hits).Count - $blocking.Count))
            $lines | ForEach-Object { Write-Host $_ }
            Set-ItResult -Skipped -Because ('report-only mode: {0} hits' -f $blocking.Count) | Out-Null
        }
        else {
            $lines | Should -BeNullOrEmpty -Because 'user-facing text must use the ConfigMgr vocabulary of lexicon\configmgr-lexicon.json'
        }
    }
}
