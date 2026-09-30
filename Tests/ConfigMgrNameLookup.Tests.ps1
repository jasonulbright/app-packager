#Requires -Modules Pester

<#
.SYNOPSIS
    Verifies that the GUI ConfigMgr version lookup finds the application
    that each packager's Package phase creates.

.DESCRIPTION
    The Package phase names the ConfigMgr application from the stage
    manifest AppName (Get-PackagedApplicationName, Default title mode).
    The GUI lookup (Get-MecmCurrentVersionByCMName) searches for the header
    CMName (default: App) as a Get-CMApplication name, where * is a
    wildcard, and then for "<CMName>*" filtered by Test-MecmApplicationTitle
    (release details only after CMName). A packager is found only when its
    created name passes one of the two.

    The created name is resolved statically from the manifest AppName
    expression and checked with sample values: a dynamic segment from a
    variable named for a version becomes a dotted version, and any other
    dynamic segment becomes an architecture, language, or channel.
#>

BeforeAll {
    $script:PackagersRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Packagers'
    $script:Wild = [char]0x2217
    $script:VersionWild = [char]0x2218

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
                    $sub = Resolve-NameExpression -Ast $n -Root $Root -Depth ($Depth + 1)
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
                if ($values.Count -eq 0) { $values = @($w) }
                if ($name -match '(?i)version') { $values = @($values | ForEach-Object { if ($_ -eq $w) { [string]$script:VersionWild } else { $_ } }) }
                return ,@($values | Select-Object -Unique)
            }
            'CommandExpressionAst' { return Resolve-NameExpression -Ast $Ast.Expression -Root $Root -Depth $Depth }
            'PipelineAst' {
                if ($Ast.PipelineElements.Count -eq 1) { return Resolve-NameExpression -Ast $Ast.PipelineElements[0] -Root $Root -Depth $Depth }
                return ,@($w)
            }
            'ParenExpressionAst' { return Resolve-NameExpression -Ast $Ast.Pipeline -Root $Root -Depth $Depth }
            'SubExpressionAst' {
                if ($Ast.SubExpression.Statements.Count -eq 1) { return Resolve-NameExpression -Ast $Ast.SubExpression.Statements[0] -Root $Root -Depth $Depth }
                return ,@($w)
            }
            'InvokeMemberExpressionAst' {
                # "$($ImageType.ToUpper())" and similar case changes of a known value.
                $member = [string]$Ast.Member.Extent.Text
                if (($null -eq $Ast.Arguments -or $Ast.Arguments.Count -eq 0) -and $member -in @('ToUpper', 'ToLower', 'ToUpperInvariant', 'ToLowerInvariant')) {
                    $inner = Resolve-NameExpression -Ast $Ast.Expression -Root $Root -Depth ($Depth + 1)
                    return ,@($inner | ForEach-Object { if ($member -like 'ToUpper*') { $_.ToUpperInvariant() } else { $_.ToLowerInvariant() } })
                }
                return ,@($w)
            }
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

    # The filter under test is the GUI's own function.
    $guiErrors = $null
    $guiAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1'), [ref]$null, [ref]$guiErrors)
    $titleFilter = $guiAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-MecmApplicationTitle' }, $false)
    if (-not $titleFilter) { throw 'Test-MecmApplicationTitle is not defined in start-apppackager.ps1.' }
    . ([scriptblock]::Create($titleFilter.Extent.Text))

    function Get-NameSamples {
        # Concrete titles for a created name. A version segment is a dotted
        # version; when CMName ends with a major, one sample continues it.
        param([string]$Created, [string]$CMName)
        $versions = @('1.2.3', '24.08')
        if ($CMName -match '(\d+)$') { $versions += ($Matches[1] + '.0.1') }
        $others = @('x64', 'en-US', 'Current Channel')
        $samples = @('')
        foreach ($ch in $Created.ToCharArray()) {
            $options = if ($ch -eq $script:VersionWild) { $versions } elseif ($ch -eq $script:Wild) { $others } else { @([string]$ch) }
            $samples = @(foreach ($s in $samples) { foreach ($o in $options) { $s + $o } })
        }
        return $samples
    }

    function Test-LookupFindsName {
        param([string]$CMName, [string]$Created)
        if ($Created.StartsWith([string]$script:Wild) -or $Created.StartsWith([string]$script:VersionWild)) { return $null }
        $pattern = New-Object System.Management.Automation.WildcardPattern($CMName, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        foreach ($title in (Get-NameSamples -Created $Created -CMName $CMName)) {
            if ($pattern.IsMatch($title) -or (Test-MecmApplicationTitle -CMName $CMName -Title $title)) { return $true }
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

    It 'finds every created name through the GUI lookup' {
        $bad = @($script:Audit | Where-Object { $_.Found -eq $false } | ForEach-Object { '{0}: CMName=''{1}'' created=''{2}''' -f $_.Script, $_.CMName, $_.Created })
        ($bad -join [Environment]::NewLine) | Should -BeNullOrEmpty -Because 'the GUI lookup searches "<CMName>" and then "<CMName> - *"'
    }

    It 'keeps release details after CMName and rejects a longer product name' {
        Test-MecmApplicationTitle -CMName 'Microsoft .NET 8' -Title 'Microsoft .NET 8.0.31 - Windows Server Hosting' | Should -BeTrue
        Test-MecmApplicationTitle -CMName 'Microsoft .NET 1' -Title 'Microsoft .NET 10.0.12 - Windows Server Hosting' | Should -BeFalse
        Test-MecmApplicationTitle -CMName 'Audacity' -Title 'Audacity 3.7.5' | Should -BeTrue
        Test-MecmApplicationTitle -CMName 'AIMP' -Title 'AIMP (x64)' | Should -BeTrue
        Test-MecmApplicationTitle -CMName 'WinMerge' -Title 'WinMerge x64' | Should -BeTrue
        Test-MecmApplicationTitle -CMName 'Contoso Tool' -Title 'Contoso Tool - 5.4.2' | Should -BeTrue
        Test-MecmApplicationTitle -CMName 'Git' -Title 'Git Extensions 4.3' | Should -BeFalse
        Test-MecmApplicationTitle -CMName 'Git' -Title 'GitHub Desktop' | Should -BeFalse
        Test-MecmApplicationTitle -CMName 'Mozilla Firefox' -Title 'Mozilla Firefox ESR (x64 en-US)' | Should -BeFalse
        Test-MecmApplicationTitle -CMName 'Microsoft Edge' -Title 'Microsoft Edge WebView2 Runtime' | Should -BeFalse
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

    It 'never lets one packager''s lookup find another packager''s application' {
        # Packagers that share a CMName are variants of one product by design.
        $rows = @($script:Audit | Where-Object { $_.Created -and $null -ne $_.Found })
        $collisions = foreach ($owner in @($rows | Select-Object Script, CMName -Unique)) {
            $prefix = ($owner.CMName -split '\*')[0]
            foreach ($other in @($rows | Where-Object { $_.Script -ne $owner.Script -and $_.CMName -ne $owner.CMName -and $_.Created.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) })) {
                if (Test-LookupFindsName -CMName $owner.CMName -Created $other.Created) { "'{0}' ({1}) finds '{2}' ({3})" -f $owner.CMName, $owner.Script, $other.Created, $other.Script }
            }
        }
        ($collisions -join [Environment]::NewLine) | Should -BeNullOrEmpty
    }}
