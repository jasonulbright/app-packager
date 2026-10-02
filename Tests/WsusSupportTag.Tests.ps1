BeforeDiscovery {
    $script:PackagerFiles = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Packagers') -Filter 'package-*.ps1' |
        ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PackagerMetadata' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    # The blocking codes Get-WsusCompatibilityFindings can report.
    $wsusSource = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\Packagers\AppPackagerWsus.psm1'))
    $script:BlockingCodes = @([regex]::Matches($wsusSource, "& \`$add 'Blocking' '(\w+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
}

Describe 'WsusSupport header tag' {
    It 'reads <Line> as <Support> <Reason>' -ForEach @(
        @{ Line = 'WsusSupport: Yes';                              Support = 'Yes'; Reason = $null }
        @{ Line = 'WsusSupport: No (PerUserInstall)';              Support = 'No';  Reason = 'PerUserInstall' }
        @{ Line = 'WsusSupport: No (CustomInstall|PayloadTooLarge)'; Support = 'No';  Reason = 'CustomInstall|PayloadTooLarge' }
        @{ Line = 'WsusSupport: No';                               Support = 'No';  Reason = $null }
        @{ Line = 'WsusSupport: Maybe';                            Support = $null; Reason = $null }
    ) {
        $path = Join-Path $TestDrive ('package-' + [guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $path -Encoding ASCII -Value @('<#', 'Vendor: Contoso', 'App: Widget', $Line, '#>')
        $meta = Get-PackagerMetadata -Path $path
        $meta.WsusSupport | Should -Be $Support
        $meta.WsusSupportReason | Should -Be $Reason
    }

    It 'is present and valid in <Name>' -ForEach $script:PackagerFiles {
        $meta = Get-PackagerMetadata -Path $Path
        $meta.WsusSupport | Should -BeIn @('Yes', 'No')
        if ($meta.WsusSupport -eq 'No') {
            $meta.WsusSupportReason | Should -Not -BeNullOrEmpty
            foreach ($code in ($meta.WsusSupportReason -split '\|')) { $code | Should -BeIn $script:BlockingCodes }
        }
        else { $meta.WsusSupportReason | Should -BeNullOrEmpty }
    }
}
