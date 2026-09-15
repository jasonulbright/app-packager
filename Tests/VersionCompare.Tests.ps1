BeforeAll {
    $t = $null; $e = $null
    $gui = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $gui.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Compare-SemVer' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    $vm = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\VersionMonitor\Module\VersionMonitorCommon.psm1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $vm.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Compare-Versions' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Single-number vendor versions' {
    It 'orders <A> against <B> as <Expected>' -TestCases @(
        @{ A = '30'; B = '31'; Expected = -1 }
        @{ A = '31'; B = '31'; Expected = 0 }
        @{ A = '31'; B = '30'; Expected = 1 }
        @{ A = '31'; B = '31.0.1'; Expected = 0 }
        @{ A = '26.2.2.2'; B = '26.2.2'; Expected = 0 }
        @{ A = '11.0.30+7'; B = '11.0.31'; Expected = -1 }
    ) {
        param($A, $B, $Expected)
        Compare-SemVer -A $A -B $B | Should -Be $Expected
    }

    It 'reports a deployed <Mecm> against vendor <Vendor> as <Status>' -TestCases @(
        @{ Mecm = '30'; Vendor = '31'; Status = 'Stale' }
        @{ Mecm = '31'; Vendor = '31'; Status = 'Current' }
        @{ Mecm = '31.0'; Vendor = '31'; Status = 'Current' }
    ) {
        param($Mecm, $Vendor, $Status)
        (Compare-Versions -MecmVersion $Mecm -VendorVersion $Vendor).Status | Should -Be $Status
    }

    It 'accepts a single-number latest version in both version checks' {
        $gui = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\start-apppackager.ps1') -Raw
        $vmText = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\VersionMonitor\Module\VersionMonitorCommon.psm1') -Raw
        $guiPattern = [regex]::Match($gui, "\`$version -notmatch '([^']+)'").Groups[1].Value
        $vmPattern = [regex]::Match($vmText, "\`$version -notmatch '([^']+)'").Groups[1].Value
        '31' | Should -Match $guiPattern
        '31' | Should -Match $vmPattern
        'Error 31 found' | Should -Not -Match $guiPattern
    }
}
