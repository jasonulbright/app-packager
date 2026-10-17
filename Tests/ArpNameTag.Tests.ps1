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
                { [void]([System.Management.Automation.WildcardPattern]::new($p, 'IgnoreCase,CultureInvariant')).IsMatch('') } | Should -Not -Throw
            }
        }
        $lines = @(Get-Content -LiteralPath $Path -TotalCount 200)
        if ($lines -match '^\s*(?:#\s*)?ArpArch\s*:') { $meta.ArpArch | Should -BeIn @('x64', 'x86') }
        if ($lines -match '^\s*(?:#\s*)?ArpVersion\s*:') {
            $meta.ArpVersion | Should -Not -BeNullOrEmpty
            { [void]([System.Management.Automation.WildcardPattern]::new($meta.ArpVersion, 'IgnoreCase,CultureInvariant')).IsMatch('') } | Should -Not -Throw
        }
        if ($meta.ArpName -eq 'none') {
            $meta.ArpVersion | Should -BeNullOrEmpty
            $meta.ArpArch | Should -BeNullOrEmpty
        }
    }
}

Describe 'ArpVersion and ArpArch header tags' {
    It 'reads <Line> as <Property> = <Expected>' -ForEach @(
        @{ Line = 'ArpVersion: 17.*';  Property = 'ArpVersion'; Expected = '17.*' }
        @{ Line = 'ArpArch: x64';      Property = 'ArpArch';    Expected = 'x64' }
        @{ Line = 'ArpArch: X86';      Property = 'ArpArch';    Expected = 'x86' }
        @{ Line = 'ArpArch: arm64';    Property = 'ArpArch';    Expected = 'arm64' }
    ) {
        $meta = Get-PackagerMetadata -Path (New-HeaderFixture -Lines @('ArpName: Widget*', $Line))
        $meta.$Property | Should -BeExactly $Expected
    }

    It 'ignores an ArpArch value that is not x64, x86 or arm64' {
        $meta = Get-PackagerMetadata -Path (New-HeaderFixture -Lines @('ArpName: Widget*', 'ArpArch: ia64'))
        $meta.ArpArch | Should -BeNullOrEmpty
    }

    It 'is null when the header has neither tag' {
        $meta = Get-PackagerMetadata -Path (New-HeaderFixture -Lines @('ArpName: Widget*'))
        $meta.ArpVersion | Should -BeNullOrEmpty
        $meta.ArpArch | Should -BeNullOrEmpty
    }

    It 'is carried onto the packager row' {
        $root = Join-Path $TestDrive ('root-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $root | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'package-widget.ps1') -Encoding ASCII -Value @('<#', 'Vendor: Contoso', 'App: Widget', 'ArpName: Widget*', 'ArpVersion: 2.*', 'ArpArch: x86', '#>')
        $row = @(Get-Packagers -Root $root)[0]
        $row.ArpVersion | Should -Be '2.*'
        $row.ArpArch | Should -Be 'x86'
    }
}

Describe 'Effective Add/Remove Programs match keys' {
    It 'gives no two packagers the same pattern, version and architecture' {
        $owners = @{}
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Packagers') -Filter 'package-*.ps1') {
            $meta = Get-PackagerMetadata -Path $file.FullName
            if (-not $meta.ArpName -or $meta.ArpName -eq 'none') { continue }
            $patterns = @($meta.ArpName -split '\|' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Sort-Object -Unique) -join '|'
            $key = '{0} / {1} / {2}' -f $patterns, ([string]$meta.ArpVersion).ToLowerInvariant(), ([string]$meta.ArpArch)
            $owners[$key] = @($owners[$key]) + $file.Name | Where-Object { $_ }
        }
        $shared = @($owners.GetEnumerator() | Where-Object { @($_.Value).Count -gt 1 } | ForEach-Object { '{0} => {1}' -f $_.Key, (@($_.Value) -join ', ') })
        $shared | Should -BeNullOrEmpty
    }
}
