#Requires -Modules Pester

<#
.SYNOPSIS
    Staging reads an installer; it does not run one on the packaging machine.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

Describe 'Packager staging' {
    It 'defines no function that installs or uninstalls the product for discovery' {
        $offenders = foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1') {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $found = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -match 'ForDiscovery|DiscoveryInstall' }, $true)
            foreach ($f in $found) { '{0}: {1}' -f $file.Name, $f.Name }
        }
        @($offenders).Count | Should -Be 0 -Because ($offenders -join '; ')
    }

    It 'stages WinSCP from the installer name and version alone' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-winscp.ps1'))
        $text | Should -Not -Match 'Start-Process\s+-FilePath\s+\$InstallerPath'
        $text | Should -Match ([regex]::Escape('$appName   = "WinSCP $version"'))
    }
}
