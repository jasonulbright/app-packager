BeforeAll {
    $script:Source = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\Packagers\package-ccleaner.ps1'))
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($script:Source, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    # The uninstall wrapper is the here-string assigned to $uninstallContent.
    $assignment = $ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left.Extent.Text -eq '$uninstallContent' }, $true)
    $script:Uninstall = $assignment.Right.Expression.Value
}

Describe 'CCleaner packager' {
    It 'sends a browser user agent to the version history page' {
        $script:Source | Should -Match "curl\.exe [^\r\n]*-A 'Mozilla/5\.0' \`$VersionHistoryUrl"
    }

    It 'names the uninstall key with the major version from version 7' {
        $script:Source | Should -Match '\$arpName = if \(\$major -ge 7\) \{ "CCleaner \$major" \} else \{ "CCleaner" \}'
        $script:Source | Should -Match 'DisplayName\s+= \$arpName'
    }

    It 'has an uninstall wrapper that parses' {
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($script:Uninstall, [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'chooses the uninstaller arguments for <Name>' -ForEach @(
        @{ Name = 'version 7'; Command = '"C:\Program Files\Common Files\Piriform\Icarus\piriform-ccl\icarus.exe" /manual_update /uninstall:piriform-ccl'
           Expected = @('/manual_update /uninstall:piriform-ccl', '/silent') }
        @{ Name = 'version 6'; Command = '"C:\Program Files\CCleaner\uninst.exe"'; Expected = @('/S') }
        @{ Name = 'an unquoted command with arguments'; Command = 'C:\Program Files\Common Files\Piriform\Icarus\piriform-ccl\icarus.exe /manual_update /uninstall:piriform-ccl'
           Expected = @('/manual_update /uninstall:piriform-ccl', '/silent') }
    ) {
        # The statements the wrapper runs on the UninstallString, taken from the wrapper text.
        $start = $script:Uninstall.IndexOf('if ($cmd -match')
        $end = $script:Uninstall.IndexOf('if (-not (Test-Path')
        $start | Should -BeGreaterThan -1
        $end | Should -BeGreaterThan $start
        $argumentLine = [regex]::Match($script:Uninstall, '(?m)^\$arguments = .+$').Value
        $argumentLine | Should -Not -BeNullOrEmpty
        $cmd = $Command
        . ([scriptblock]::Create($script:Uninstall.Substring($start, $end - $start) + "`n" + $argumentLine))
        $arguments | Should -Be $Expected
        $exe | Should -Match 'icarus\.exe$|uninst\.exe$'
    }

    It 'prefers the entry with the version number in its key name' {
        $script:Uninstall | Should -Match 'Sort-Object -Property PSChildName -Descending'
        @('CCleaner', 'CCleaner 7' | Sort-Object -Descending)[0] | Should -Be 'CCleaner 7'
    }

    It 'fails when the registered uninstaller is missing' {
        $script:Uninstall | Should -Match 'uninstaller not found[^\r\n]*exit 1'
    }
}

Describe 'CCleaner install wrapper' {
    It 'reports success for the installer code 111111 and passes other codes through' {
        $source = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\Packagers\package-ccleaner.ps1'))
        $source | Should -Match ([regex]::Escape('if ($proc.ExitCode -eq 111111) { exit 0 }'))
        $source | Should -Match '(?m)^exit \$proc\.ExitCode\r?$'
        $source | Should -Match '-InstallPs1Content \$installContent'
    }
}
