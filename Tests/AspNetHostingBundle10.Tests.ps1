BeforeAll {
    $script:Path = Join-Path $PSScriptRoot '..\Packagers\package-aspnethostingbundle10.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Path, [ref]$t, [ref]$e)
    $script:ParseErrors = $e
    $script:CompoundTables = @($script:Ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.HashtableAst] -and
        @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Type' -and $_.Item2.Extent.Text -match 'Compound' }).Count -gt 0
    }, $true))
}

Describe 'ASP.NET Core 10 Hosting Bundle detection' {
    It 'parses without errors' {
        $script:ParseErrors | Should -BeNullOrEmpty
    }

    It 'joins the version and successor-patch clauses with a plain OR and no clause groups' {
        $script:CompoundTables.Count | Should -Be 1
        $keys = @($script:CompoundTables[0].KeyValuePairs | ForEach-Object { $_.Item1.Extent.Text })
        $keys | Should -Contain 'Connector'
        $keys | Should -Not -Contain 'GroupSizes'
        ($script:CompoundTables[0].KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Connector' }).Item2.Extent.Text | Should -Match 'Or'
    }

    It 'checks one file per version' {
        $clauses = ($script:CompoundTables[0].KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Clauses' }).Item2
        @($clauses.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)).Count | Should -Be 2
    }
}
