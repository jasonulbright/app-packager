BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-defraggler.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $functions = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
    $fn = $functions | Where-Object { $_.Name -eq 'Find-NewestDefragglerBuild' } | Select-Object -First 1
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Defraggler build walk' {
    It 'returns nothing when the floor build is not served' {
        Find-NewestDefragglerBuild -Floor 222 -MaxHops 12 -IsPublished { param($b) $false } | Should -BeNullOrEmpty
    }

    It 'returns the floor build when no later build is served' {
        Find-NewestDefragglerBuild -Floor 222 -MaxHops 12 -IsPublished { param($b) $b -eq 222 } | Should -Be 222
    }

    It 'walks to <Expected> when the host serves <Served>' -TestCases @(
        @{ Served = @(222, 223, 224); Expected = 224 }
        @{ Served = @(222, 224); Expected = 224 }
        @{ Served = @(222, 225); Expected = 225 }
        @{ Served = @(222, 223, 226, 227); Expected = 227 }
        @{ Served = @(222, 223, 227); Expected = 223 }
        @{ Served = @(222, 226); Expected = 222 }
        @{ Served = @(221, 222); Expected = 222 }
    ) {
        param($Served, $Expected)
        Find-NewestDefragglerBuild -Floor 222 -MaxHops 12 -IsPublished { param($b) $Served -contains $b } | Should -Be $Expected
    }

    It 'stops after the hop limit' {
        $served = 222..260
        Find-NewestDefragglerBuild -Floor 222 -MaxHops 3 -IsPublished { param($b) $served -contains $b } | Should -Be 225
    }

    It 'probes at most three builds above the current one per hop' {
        $script:Probed = New-Object System.Collections.Generic.List[int]
        Find-NewestDefragglerBuild -Floor 222 -MaxHops 12 -IsPublished { param($b) $script:Probed.Add($b); $b -eq 222 } | Out-Null
        $script:Probed | Should -Be @(222, 223, 224, 225)
    }
}
