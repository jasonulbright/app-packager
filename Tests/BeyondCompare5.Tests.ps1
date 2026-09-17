BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-beyondcompare5.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in @('Get-BeyondCompare5ReleaseFromPage', 'Resolve-BeyondCompare5KeyFile')) {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:DownloadHost = 'https://www.scootersoftware.com'
    function Get-PackagerPreferences { $script:StoredPreferences }
}

Describe 'Beyond Compare 5 release lookup' {
    It 'reads the version and zip URL from the alternate download page' {
        $html = '<p>Current version 5.2.5</p><a href="/files/BCompareSetup-5.2.5.32528.zip">BCompareSetup-5.2.5.32528.zip</a> (27299kb)'
        $release = Get-BeyondCompare5ReleaseFromPage -Html $html
        $release.Version | Should -Be '5.2.5.32528'
        $release.FileName | Should -Be 'BCompareSetup-5.2.5.32528.zip'
        $release.DownloadUrl | Should -Be 'https://www.scootersoftware.com/files/BCompareSetup-5.2.5.32528.zip'
    }

    It 'returns nothing when the page carries no setup zip link' {
        Get-BeyondCompare5ReleaseFromPage -Html '<html>maintenance</html>' | Should -BeNullOrEmpty
    }
}

Describe 'Beyond Compare 5 license key file' {
    BeforeEach { $script:StoredPreferences = $null }

    It 'refuses to continue when no key file is configured' {
        { Resolve-BeyondCompare5KeyFile -Override '' } | Should -Throw '*requires a license key file*'
    }

    It 'refuses a configured key file that does not exist' {
        $script:StoredPreferences = [pscustomobject]@{ BeyondCompareKeyFile = (Join-Path $TestDrive 'missing\BC5Key.txt') }
        { Resolve-BeyondCompare5KeyFile -Override '' } | Should -Throw '*key file not found*'
    }

    It 'uses the key file chosen in packager preferences' {
        $key = Join-Path $TestDrive 'bc5key.txt'
        Set-Content -LiteralPath $key -Value 'placeholder' -Encoding ASCII
        $script:StoredPreferences = [pscustomobject]@{ BeyondCompareKeyFile = $key }
        Resolve-BeyondCompare5KeyFile -Override '' | Should -Be $key
    }

    It 'lets -KeyFile override the stored preference' {
        $stored = Join-Path $TestDrive 'stored.txt'
        $override = Join-Path $TestDrive 'override.txt'
        Set-Content -LiteralPath $stored, $override -Value 'placeholder' -Encoding ASCII
        $script:StoredPreferences = [pscustomobject]@{ BeyondCompareKeyFile = $stored }
        Resolve-BeyondCompare5KeyFile -Override $override | Should -Be $override
    }
}
