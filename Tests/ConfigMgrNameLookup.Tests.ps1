#Requires -Modules Pester

<#
.SYNOPSIS
    Verifies that the GUI ConfigMgr version lookup finds the application
    that each packager's Package phase creates.

.DESCRIPTION
    The Package phase names the ConfigMgr application from the stage
    manifest AppName (Get-PackagedApplicationName, Default title mode).
    The GUI lookup (Get-MecmCurrentVersionByCMName) searches for the header
    CMName (default: App), first exactly and then as the prefix "<CMName>*".
    A packager is found only when its created name equals CMName or starts
    with CMName.

    The created name is resolved statically from the manifest AppName
    expression. Dynamic segments (version, channel, architecture variables)
    resolve to a wildcard of any length; some expansion of the created name
    must start with CMName.
#>

BeforeAll {
    $script:PackagersRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Packagers'
    $script:Wild = [char]0x2217

    # The created name comes from the installer ProductName or the ARP
    # DisplayName at Stage time, so no static check applies. A vendor rename
    # of that value makes the GUI lookup miss the application.
    $script:VendorDerivedNames = @(
        'package-7zip.ps1'
        'package-azurepowershell.ps1'
        'package-chromeremotedesktophost.ps1'
        'package-gcpw.ps1'
        'package-intunedebugtoolkit.ps1'
        'package-keepass.ps1'
        'package-msodbcsql18.ps1'
        'package-msoledb.ps1'
        'package-nodejs.ps1'
        'package-paintdotnet.ps1'
        'package-powershell7.ps1'
        'package-putty.ps1'
        'package-tortoisegit.ps1'
        'package-tortoisesvn.ps1'
        'package-vlc.ps1'
        'package-webex.ps1'
        'package-windirstat.ps1'
        'package-winscp.ps1'
        'package-wireshark.ps1'
    )

    function Get-HeaderCMName {
        param([string[]]$Lines)
        $app = $null; $cm = $null
        foreach ($l in ($Lines | Select-Object -First 40)) {
            if (-not $app -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')    { $app = $Matches[1].Trim(); continue }
            if (-not $cm  -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$') { $cm  = $Matches[1].Trim(); continue }
        }
        if ($cm) { return $cm }
        return $app
    }

    function Resolve-NameExpression {
        param($Ast, $Root, [int]$Depth = 0)
        $w = [string]$script:Wild
        if ($Depth -gt 6 -or $null -eq $Ast) { return ,@($w) }
        switch ($Ast.GetType().Name) {
            'StringConstantExpressionAst' { return ,@([string]$Ast.Value) }
            'ConstantExpressionAst' { return ,@([string]$Ast.Value) }
            'ExpandableStringExpressionAst' {
                $parts = @('')
                $pos = 0
                $nested = @($Ast.NestedExpressions | Sort-Object { $_.Extent.StartOffset })
                $base = $Ast.Extent.StartOffset + 1
                foreach ($n in $nested) {
                    $rel = $n.Extent.StartOffset - $base
                    $lit = $Ast.Extent.Text.Substring(1 + $pos, $rel - $pos)
                    $sub = if ($n -is [System.Management.Automation.Language.VariableExpressionAst]) { Resolve-NameExpression -Ast $n -Root $Root -Depth ($Depth + 1) } else { ,@($w) }
                    $next = @()
                    foreach ($p in $parts) { foreach ($s in $sub) { $next += ($p + $lit + $s) } }
                    $parts = $next
                    $pos = $rel + $n.Extent.Text.Length
                }
                $tail = $Ast.Extent.Text.Substring(1 + $pos, $Ast.Extent.Text.Length - 2 - $pos)
                return ,@($parts | ForEach-Object { $_ + $tail } | ForEach-Object { $_ -replace '`', '' })
            }
            'VariableExpressionAst' {
                $name = $Ast.VariablePath.UserPath
                $values = @()
                $assigns = @($Root.FindAll({
                    param($a)
                    $a -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                    $a.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $a.Left.VariablePath.UserPath -eq $name
                }, $true))
                foreach ($a in $assigns) { $values += Resolve-NameExpression -Ast $a.Right -Root $Root -Depth ($Depth + 1) }
                $params = @($Root.FindAll({
                    param($p)
                    $p -is [System.Management.Automation.Language.ParameterAst] -and $p.Name.VariablePath.UserPath -eq $name
                }, $true))
                foreach ($p in $params) {
                    if ($p.DefaultValue) { $values += Resolve-NameExpression -Ast $p.DefaultValue -Root $Root -Depth ($Depth + 1) }
                    else { $values += $w }
                }
                if ($values.Count -eq 0) { return ,@($w) }
                return ,@($values | Select-Object -Unique)
            }
            'CommandExpressionAst' { return Resolve-NameExpression -Ast $Ast.Expression -Root $Root -Depth $Depth }
            'PipelineAst' {
                if ($Ast.PipelineElements.Count -eq 1) { return Resolve-NameExpression -Ast $Ast.PipelineElements[0] -Root $Root -Depth $Depth }
                return ,@($w)
            }
            'ParenExpressionAst' { return Resolve-NameExpression -Ast $Ast.Pipeline -Root $Root -Depth $Depth }
            'IfStatementAst' {
                $values = @()
                foreach ($c in $Ast.Clauses) { foreach ($s in $c.Item2.Statements) { $values += Resolve-NameExpression -Ast $s -Root $Root -Depth ($Depth + 1) } }
                if ($Ast.ElseClause) { foreach ($s in $Ast.ElseClause.Statements) { $values += Resolve-NameExpression -Ast $s -Root $Root -Depth ($Depth + 1) } }
                return ,@($values)
            }
            default { return ,@($w) }
        }
    }

    function Get-CreatedNameCandidates {
        param([string]$Path)
        $tokens = $null; $errors = $null
        $root = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
        $tables = @($root.FindAll({
            param($h)
            $h -is [System.Management.Automation.Language.HashtableAst] -and
            @($h.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'AppName' }).Count -gt 0 -and
            @($h.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'SoftwareVersion' }).Count -gt 0
        }, $true))
        $out = @()
        foreach ($t in $tables) {
            $pair = $t.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'AppName' } | Select-Object -First 1
            $out += Resolve-NameExpression -Ast $pair.Item2 -Root $root
        }
        return ,@($out | Select-Object -Unique)
    }

    function Test-LookupFindsName {
        param([string]$CMName, [string]$Created)
        $w = [string]$script:Wild
        if ($Created.StartsWith($w)) { return $null }
        # A wildcard segment stands for a runtime value of any length; the
        # check asks whether some expansion of Created starts with CMName.
        return (Test-PatternCanStartWith -Pieces $Created.Split($script:Wild) -Index 0 -Target $CMName)
    }

    function Test-PatternCanStartWith {
        param([string[]]$Pieces, [int]$Index, [string]$Target)
        if ($Target.Length -eq 0) { return $true }
        if ($Index -ge $Pieces.Count) { return $false }
        $lit = $Pieces[$Index]
        if ($Target.Length -le $lit.Length) {
            return $lit.StartsWith($Target, [StringComparison]::OrdinalIgnoreCase)
        }
        if (-not $Target.StartsWith($lit, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        if ($Index -eq $Pieces.Count - 1) { return $false }
        $rest = $Target.Substring($lit.Length)
        for ($i = 0; $i -le $rest.Length; $i++) {
            if (Test-PatternCanStartWith -Pieces $Pieces -Index ($Index + 1) -Target $rest.Substring($i)) { return $true }
        }
        return $false
    }

    $script:Audit = foreach ($file in Get-ChildItem -LiteralPath $script:PackagersRoot -Filter 'package-*.ps1' | Sort-Object Name) {
        $cmName = Get-HeaderCMName -Lines (Get-Content -LiteralPath $file.FullName -TotalCount 40)
        $candidates = Get-CreatedNameCandidates -Path $file.FullName
        foreach ($c in $candidates) {
            [pscustomobject]@{
                Script  = $file.Name
                CMName  = $cmName
                Created = $c
                Found   = Test-LookupFindsName -CMName $cmName -Created $c
            }
        }
        if ($candidates.Count -eq 0) {
            [pscustomobject]@{ Script = $file.Name; CMName = $cmName; Created = $null; Found = $null }
        }
    }
}

Describe 'ConfigMgr name lookup matches the Package phase name' {
    It 'finds a manifest AppName in every packager' {
        $missing = @($script:Audit | Where-Object { $null -eq $_.Created } | ForEach-Object { $_.Script })
        ($missing -join [Environment]::NewLine) | Should -BeNullOrEmpty -Because 'the created ConfigMgr name comes from the manifest AppName'
    }

    It 'resolves a literal leading segment for every created name outside the vendor-derived list' {
        $open = @($script:Audit | Where-Object { $_.Created -and $null -eq $_.Found -and $script:VendorDerivedNames -notcontains $_.Script } | ForEach-Object { '{0}: CMName=''{1}'' created=''{2}''' -f $_.Script, $_.CMName, $_.Created })
        ($open -join [Environment]::NewLine) | Should -BeNullOrEmpty -Because 'an unresolved prefix cannot be proven to match the lookup'
    }

    It 'lists only packagers whose created name is still vendor-derived' {
        $unresolved = @($script:Audit | Where-Object { $_.Created -and $null -eq $_.Found } | ForEach-Object { $_.Script })
        $stale = @($script:VendorDerivedNames | Where-Object { $unresolved -notcontains $_ })
        ($stale -join [Environment]::NewLine) | Should -BeNullOrEmpty -Because 'a resolvable name must be checked by the CMName rule'
    }

    It 'starts every created name with the header CMName' {
        $bad = @($script:Audit | Where-Object { $_.Found -eq $false } | ForEach-Object { '{0}: CMName=''{1}'' created=''{2}''' -f $_.Script, $_.CMName, $_.Created })
        ($bad -join [Environment]::NewLine) | Should -BeNullOrEmpty -Because 'the GUI lookup searches "<CMName>" and then "<CMName> - *"'
    }

    It 'names RStudio so that the lookup finds it' {
        $rows = @($script:Audit | Where-Object { $_.Script -eq 'package-rstudio.ps1' })
        $rows | Should -Not -BeNullOrEmpty
        @($rows | Where-Object { $_.Found -ne $true }) | Should -BeNullOrEmpty
    }
}

Describe 'ConfigMgr name lookup prefix safety' {
    It 'never lets one CMName match another CMName through the titled-name wildcard' {
        $names = @(Get-ChildItem "$PSScriptRoot\..\Packagers" -Filter 'package-*.ps1' | ForEach-Object {
            $line = Select-String -Path $_.FullName -Pattern '^CMName:\s*(.+?)\s*$' | Select-Object -First 1
            if ($line) { $line.Matches[0].Groups[1].Value }
        } | Sort-Object -Unique)
        $collisions = foreach ($a in $names) { foreach ($b in $names) { if ($b -ne $a -and $b.StartsWith($a + ' - ', [StringComparison]::OrdinalIgnoreCase)) { "'$a' matches '$b'" } } }
        ($collisions -join [Environment]::NewLine) | Should -BeNullOrEmpty
    }
}
