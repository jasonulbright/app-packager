BeforeAll {
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-firefoxesr.ps1'
    $t = $null; $e = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-EsrFileVersion' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    $gui = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $runner = $gui.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-PackagerGetLatestVersion' }, $false)
    $script:GuiVersionPattern = [regex]::Match($runner.Extent.Text, "\`$pattern = '([^']+)'").Groups[1].Value
}

Describe 'Firefox ESR version strings' {
    It 'strips the esr suffix from the release train name' {
        Get-EsrFileVersion -Version '140.17.0esr' | Should -Be '140.17.0'
        Get-EsrFileVersion -Version '140.17.0' | Should -Be '140.17.0'
    }

    It 'prints a version the GUI accepts for -GetLatestVersionOnly' {
        $block = $script:Ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.IfStatementAst] -and
            $n.Clauses[0].Item1.Extent.Text -eq '$GetLatestVersionOnly' }, $false)
        $block | Should -Not -BeNullOrEmpty
        $block.Extent.Text | Should -Match 'Write-Output \(Get-EsrFileVersion -Version \$v\)'
        (Get-EsrFileVersion -Version '140.17.0esr') | Should -Match $script:GuiVersionPattern
        '140.17.0esr' | Should -Not -Match $script:GuiVersionPattern
    }

    It 'records the numeric version and downloads the esr-named installer' {
        $stage = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-StageFirefoxEsr' }, $false).Extent.Text
        $stage | Should -Match '\$version = Get-EsrFileVersion -Version \$releaseTrain'
        $stage | Should -Match '"Firefox Setup \$releaseTrain\.msi"'
        $stage | Should -Match '\$DownloadBase/\$releaseTrain/win64'
        $stage | Should -Match 'SoftwareVersion = \$version'
    }
}
