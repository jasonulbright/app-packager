BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:Path = Join-Path $PSScriptRoot '..\Packagers\package-aspnethostingbundle8.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Path, [ref]$t, [ref]$e)
    $script:ParseErrors = $e
    $script:ManifestTable = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.HashtableAst] -and
        @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'WsusDetection' }).Count -gt 0
    }, $true)
    $script:EntryAssignment = $script:Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$bundleEntry'
    }, $true)
    # A value that names a variable reads as the literal string the packager
    # assigns to it.
    $script:Literals = @{}
    foreach ($a in $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
        if ($a.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and $a.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
            $script:Literals[$a.Left.VariablePath.UserPath] = $a.Right.Expression.Value
        }
    }
    function Get-ValueAst {
        param($Table, [string]$Key)
        $pair = @($Table.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq $Key })
        if ($pair.Count -eq 0) { return $null }
        $v = $pair[0].Item2
        if ($v -is [System.Management.Automation.Language.PipelineAst]) { $v = $v.PipelineElements[0] }
        if ($v -is [System.Management.Automation.Language.CommandExpressionAst]) { $v = $v.Expression }
        return $v
    }
    function Get-ClauseValue {
        param($Table, [string]$Key)
        $v = Get-ValueAst $Table $Key
        if ($null -eq $v) { return $null }
        if ($v -is [System.Management.Automation.Language.StringConstantExpressionAst]) { return $v.Value }
        if ($v -is [System.Management.Automation.Language.VariableExpressionAst]) {
            $name = $v.VariablePath.UserPath
            if ($name -in @('true', 'false')) { return ($name -eq 'true') }
            if ($script:Literals.ContainsKey($name)) { return $script:Literals[$name] }
            return '$' + $name
        }
        return $v.Extent.Text
    }
    $right = $script:EntryAssignment.Right
    if ($right -is [System.Management.Automation.Language.PipelineAst]) { $right = $right.PipelineElements[0] }
    if ($right -is [System.Management.Automation.Language.CommandExpressionAst]) { $right = $right.Expression }
    $script:EntryTable = $right
}

Describe 'ASP.NET Core 8 Hosting Bundle detection' {
    It 'parses without errors' {
        $script:ParseErrors | Should -BeNullOrEmpty
    }

    It 'detects with a script and no clause groups' {
        $detection = Get-ValueAst $script:ManifestTable 'Detection'
        $detection | Should -BeOfType [System.Management.Automation.Language.HashtableAst]
        Get-ClauseValue $detection 'Type' | Should -Be 'Script'
        Get-ClauseValue $detection 'ScriptLanguage' | Should -Be 'PowerShell'
        Get-ClauseValue $detection 'ScriptText' | Should -Be '(New-ArpEntryDetectionScript -Entry $bundleEntry)'
        @($script:Ast.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.HashtableAst] -and
            @($n.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -in @('GroupSizes', 'Clauses', 'Connector') }).Count -gt 0
        }, $true)).Count | Should -Be 0
    }

    It 'gives WSUS the same Add/Remove Programs entry as the detection script' {
        Get-ClauseValue $script:ManifestTable 'WsusDetection' | Should -Be '$bundleEntry'
    }

    It 'names the hosting bundle''s own 32-bit Add/Remove Programs entry' {
        $script:EntryTable | Should -BeOfType [System.Management.Automation.Language.HashtableAst]
        Get-ClauseValue $script:EntryTable 'Type' | Should -Be 'ArpEntry'
        Get-ClauseValue $script:EntryTable 'View' | Should -Be '32'
        Get-ClauseValue $script:EntryTable 'DisplayNamePrefix' | Should -Be 'Microsoft .NET 8.0.'
        Get-ClauseValue $script:EntryTable 'DisplayNameSuffix' | Should -Be ' - Windows Server Hosting'
        Get-ClauseValue $script:EntryTable 'Publisher' | Should -Be 'Microsoft Corporation'
        Get-ClauseValue $script:EntryTable 'Version' | Should -Be '$version'
    }

    It 'builds a script that reads the 32-bit view for that entry' {
        $entry = @{}
        foreach ($key in 'Type', 'View', 'DisplayNamePrefix', 'DisplayNameSuffix', 'Publisher') { $entry[$key] = Get-ClauseValue $script:EntryTable $key }
        $entry.Version = '8.0.31'
        $text = New-ArpEntryDetectionScript -Entry $entry
        $text | Should -Match ([regex]::Escape('@([Microsoft.Win32.RegistryView]::Registry32)'))
        $text | Should -Not -Match 'Registry64'
        $text | Should -Match ([regex]::Escape("StartsWith('Microsoft .NET 8.0.', [StringComparison]::OrdinalIgnoreCase)"))
        $text | Should -Match ([regex]::Escape("EndsWith(' - Windows Server Hosting', [StringComparison]::OrdinalIgnoreCase)"))
        $text | Should -Match ([regex]::Escape("[string]::Equals(`$publisher, 'Microsoft Corporation', [StringComparison]::OrdinalIgnoreCase)"))
        $text | Should -Match ([regex]::Escape("[version]'8.0.31.0'"))
    }
}
