BeforeDiscovery {
    $script:PackagerFiles = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Packagers') -Filter 'package-*.ps1' |
        ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-PackagerMetadata', 'Get-Packagers') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    function New-HeaderFixture {
        param([string[]]$Lines)
        $path = Join-Path $TestDrive ('package-' + [guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $path -Encoding ASCII -Value (@('<#', 'Vendor: Contoso', 'App: Widget') + $Lines + @('#>'))
        return $path
    }
}

Describe 'ArpName header tag' {
    It 'reads <Line> as <Expected>' -ForEach @(
        @{ Line = 'ArpName: Contoso Widget*';                 Expected = 'Contoso Widget*' }
        @{ Line = 'ArpName: Contoso Widget* | Widget [0-9]*'; Expected = 'Contoso Widget* | Widget [0-9]*' }
        @{ Line = 'ArpName: none';                            Expected = 'none' }
        @{ Line = 'ArpName:    Widget   ';                    Expected = 'Widget' }
    ) {
        $meta = Get-PackagerMetadata -Path (New-HeaderFixture -Lines @($Line))
        $meta.ArpName | Should -Be $Expected
    }

    It 'is null when the header has no ArpName line' {
        $meta = Get-PackagerMetadata -Path (New-HeaderFixture -Lines @('CMName: Widget'))
        $meta.ArpName | Should -BeNullOrEmpty
    }

    It 'is carried onto the packager row' {
        $root = Join-Path $TestDrive ('root-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $root | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'package-widget.ps1') -Encoding ASCII -Value @('<#', 'Vendor: Contoso', 'App: Widget', 'ArpName: Widget*', '#>')
        $row = @(Get-Packagers -Root $root)[0]
        $row.ArpName | Should -Be 'Widget*'
    }

    It 'is present and valid in <Name>' -ForEach $script:PackagerFiles {
        $meta = Get-PackagerMetadata -Path $Path
        $meta.ArpName | Should -Not -BeNullOrEmpty
        if ($meta.ArpName -ne 'none') {
            foreach ($pattern in ($meta.ArpName -split '\|')) {
                $p = $pattern.Trim()
                $p | Should -Not -BeNullOrEmpty
                $p | Should -Not -Match '^[\*\?\s]+$'
                $p | Should -Not -BeExactly 'none'
                # A pattern must name at least three literal characters before its first wildcard.
                ($p -replace '[\*\?\[].*$', '').Length | Should -BeGreaterOrEqual 3
            }
        }
    }
}
