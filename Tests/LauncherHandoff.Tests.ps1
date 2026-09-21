BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-LauncherConnectionDefault' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Set-LauncherConnectionDefault' {
    BeforeEach {
        $script:savedSite = $env:SUITE_CM_SITECODE
        $script:savedProvider = $env:SUITE_CM_PROVIDER
    }
    AfterEach {
        $env:SUITE_CM_SITECODE = $script:savedSite
        $env:SUITE_CM_PROVIDER = $script:savedProvider
    }

    It 'fills an empty site code and provider from the launcher values' {
        $env:SUITE_CM_SITECODE = ' ABC '
        $env:SUITE_CM_PROVIDER = 'cm.example.test'
        $prefs = [pscustomobject]@{ SiteCode = ''; ProviderMachineName = '' }
        Set-LauncherConnectionDefault -Prefs $prefs
        $prefs.SiteCode | Should -Be 'ABC'
        $prefs.ProviderMachineName | Should -Be 'cm.example.test'
    }

    It 'keeps a saved site code and provider' {
        $env:SUITE_CM_SITECODE = 'ABC'
        $env:SUITE_CM_PROVIDER = 'cm.example.test'
        $prefs = [pscustomobject]@{ SiteCode = 'XYZ'; ProviderMachineName = 'own.example.test' }
        Set-LauncherConnectionDefault -Prefs $prefs
        $prefs.SiteCode | Should -Be 'XYZ'
        $prefs.ProviderMachineName | Should -Be 'own.example.test'
    }

    It 'changes nothing when the launcher gives no values' {
        $env:SUITE_CM_SITECODE = $null
        $env:SUITE_CM_PROVIDER = ''
        $prefs = [pscustomobject]@{ SiteCode = ''; ProviderMachineName = '' }
        Set-LauncherConnectionDefault -Prefs $prefs
        $prefs.SiteCode | Should -Be ''
        $prefs.ProviderMachineName | Should -Be ''
    }
}
