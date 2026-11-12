BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerVulnerability.psd1') -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PackagerMetadata' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
    $script:Headers = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Packagers') -Filter 'package-*.ps1' | ForEach-Object {
            [pscustomobject]@{ Name = $_.Name; Cpe = [string](Get-PackagerMetadata -Path $_.FullName).CPE }
        })
}

Describe 'CPE header tag' {
    It 'reads a header with several CPEs as one value' {
        $path = Join-Path $TestDrive ('package-' + [guid]::NewGuid().ToString('N') + '.ps1')
        $value = 'cpe:2.3:a:contoso:widget_esr:*:*:*:*:*:*:*:* cpe:2.3:a:contoso:widget:*:*:*:*:esr:*:*:*'
        Set-Content -LiteralPath $path -Encoding ASCII -Value @('<#', 'Vendor: Contoso', 'App: Widget', ('CPE: ' + $value), '#>')
        (Get-PackagerMetadata -Path $path).CPE | Should -Be $value
        @(ConvertTo-VulnerabilityCpeList -Cpe $value).Count | Should -Be 2
        Test-VulnerabilityCpeList -Cpe $value | Should -BeTrue
    }

    It 'is a valid CPE 2.3 value in every packager that has one' {
        $script:Headers.Count | Should -BeGreaterThan 0
        $invalid = @($script:Headers | Where-Object { $_.Cpe -and -not (Test-VulnerabilityCpeList -Cpe $_.Cpe) } | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Cpe })
        $invalid | Should -BeNullOrEmpty
    }

    It 'names an application product with thirteen fields in every packager that has one' {
        # A header with fewer fields, or with part h or o, reads as another product in NVD.
        $odd = @(foreach ($h in $script:Headers) {
                foreach ($one in @(ConvertTo-VulnerabilityCpeList -Cpe $h.Cpe)) {
                    if (@($one -split '(?<!\\):').Count -ne 13 -or $one -notmatch '^cpe:2\.3:a:') { '{0}: {1}' -f $h.Name, $one }
                }
            })
        $odd | Should -BeNullOrEmpty
    }
}
