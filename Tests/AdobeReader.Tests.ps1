BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-adobereader.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $script:AdobeDownloadBase = 'https://ardownload3.adobe.com/pub/adobe/reader/win/AcrobatDC'
    $script:AdobeMuiLocales = @('ca_ES','cs_CZ','da_DK','de_DE','en_US','es_ES','eu_ES','fi_FI','fr_FR','hr_HR','hu_HU','it_IT','ja_JP','ko_KR','nb_NO','nl_NL','pl_PL','pt_BR','ro_RO','ru_RU','sk_SK','sl_SI','sv_SE','tr_TR','uk_UA','zh_CN','zh_TW')
    foreach ($name in 'ConvertTo-AdobeUrlVersion', 'Get-AdobeInstallerSuffix', 'Get-AdobeInstallerInfo', 'Get-AdobePatchInfo', 'ConvertTo-AdobeLanguageList', 'Get-AdobeReaderInstallOptions', 'Set-AdobeSetupIniCommandLine') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function Get-PackagerPreferences { $script:TestPackagerPrefs }
    $script:TestPackagerPrefs = $null
}

Describe 'Adobe Reader installer names' {
    It 'names the English full installer and patch' {
        (Get-AdobeInstallerInfo -Version '26.002.21901').FileName | Should -Be 'AcroRdrDC2600221901_en_US.exe'
        (Get-AdobePatchInfo -Version '26.002.21901').FileName | Should -Be 'AcroRdrDCUpd2600221901.msp'
    }
    It 'names the MUI full installer and pairs it with the MUI patch' {
        $full = Get-AdobeInstallerInfo -Version '26.002.21901' -Edition 'MUI'
        $full.FileName | Should -Be 'AcroRdrDC2600221901_MUI.exe'
        $full.DownloadUrl | Should -Be 'https://ardownload3.adobe.com/pub/adobe/reader/win/AcrobatDC/2600221901/AcroRdrDC2600221901_MUI.exe'
        (Get-AdobePatchInfo -Version '26.002.21901' -Edition 'MUI').FileName | Should -Be 'AcroRdrDCUpd2600221901_MUI.msp'
    }
}

Describe 'ConvertTo-AdobeLanguageList' {
    It 'adds English, removes duplicates and sorts' {
        ConvertTo-AdobeLanguageList -Languages @('fr_FR', 'de_DE', 'fr_FR') | Should -Be @('de_DE', 'en_US', 'fr_FR')
    }
    It 'keeps All alone' {
        ConvertTo-AdobeLanguageList -Languages @('de_DE', 'All') | Should -Be @('All')
    }
    It 'gives English only for an empty selection' {
        ConvertTo-AdobeLanguageList -Languages @() | Should -Be @('en_US')
    }
    It 'refuses an unknown locale code' {
        { ConvertTo-AdobeLanguageList -Languages @('xx_XX') } | Should -Throw '*Unknown Adobe locale*'
    }
}

Describe 'Get-AdobeReaderInstallOptions' {
    It 'defaults to English with no stored preference' {
        $script:TestPackagerPrefs = $null
        $o = Get-AdobeReaderInstallOptions
        $o.Edition | Should -Be 'English'
        @($o.Languages).Count | Should -Be 0
    }
    It 'reads the stored MUI preference and normalizes its languages' {
        $script:TestPackagerPrefs = [pscustomobject]@{ AdobeReaderInstallOptions = [pscustomobject]@{ Edition = 'MUI'; Languages = @('ja_JP', 'de_DE') } }
        $o = Get-AdobeReaderInstallOptions
        $o.Edition | Should -Be 'MUI'
        $o.Languages | Should -Be @('de_DE', 'en_US', 'ja_JP')
    }
    It 'lets the parameters win over the stored preference' {
        $script:TestPackagerPrefs = [pscustomobject]@{ AdobeReaderInstallOptions = [pscustomobject]@{ Edition = 'MUI'; Languages = @('ja_JP') } }
        (Get-AdobeReaderInstallOptions -Edition 'English').Edition | Should -Be 'English'
        (Get-AdobeReaderInstallOptions -Languages 'fr_FR, es_ES').Languages | Should -Be @('en_US', 'es_ES', 'fr_FR')
    }
    It 'drops the languages for the English edition' {
        $script:TestPackagerPrefs = [pscustomobject]@{ AdobeReaderInstallOptions = [pscustomobject]@{ Edition = 'English'; Languages = @('ja_JP') } }
        @((Get-AdobeReaderInstallOptions).Languages).Count | Should -Be 0
    }
}

Describe 'Set-AdobeSetupIniCommandLine' {
    It 'adds CmdLine to the Product section and keeps the other keys' {
        $ini = Join-Path $TestDrive 'setup.ini'
        Set-Content $ini -Value "[Startup]`r`nRequireMSI=3.0`r`n`r`n[Product]`r`nPATCH=AcroRdrDCUpd2600221901_MUI.msp`r`nmsi=AcroRead.msi`r`n`r`n[Windows 10]`r`nPlatformID=2"
        Set-AdobeSetupIniCommandLine -Path $ini -CommandLine 'LANG_LIST="de_DE,en_US" SUPPRESSLANGSELECTION=YES'
        $lines = Get-Content $ini
        $lines | Should -Contain 'CmdLine=LANG_LIST="de_DE,en_US" SUPPRESSLANGSELECTION=YES'
        $lines | Should -Contain 'PATCH=AcroRdrDCUpd2600221901_MUI.msp'
        $lines | Should -Contain 'msi=AcroRead.msi'
        $lines | Should -Contain 'PlatformID=2'
        $productIndex = [array]::IndexOf($lines, '[Product]')
        $cmdIndex = [array]::IndexOf($lines, 'CmdLine=LANG_LIST="de_DE,en_US" SUPPRESSLANGSELECTION=YES')
        $nextSection = [array]::IndexOf($lines, '[Windows 10]')
        ($cmdIndex -gt $productIndex -and $cmdIndex -lt $nextSection) | Should -BeTrue
    }
    It 'replaces an existing CmdLine once' {
        $ini = Join-Path $TestDrive 'setup2.ini'
        Set-Content $ini -Value "[Product]`r`nCmdLine=TRANSFORMS=`"old.mst`"`r`nmsi=AcroRead.msi"
        Set-AdobeSetupIniCommandLine -Path $ini -CommandLine 'LANG_LIST="All" SUPPRESSLANGSELECTION=YES'
        $lines = Get-Content $ini
        @($lines | Where-Object { $_ -like 'CmdLine=*' }).Count | Should -Be 1
        $lines | Should -Contain 'CmdLine=LANG_LIST="All" SUPPRESSLANGSELECTION=YES'
        $lines | Should -Not -Contain 'CmdLine=TRANSFORMS="old.mst"'
    }
    It 'removes the key for an empty command line' {
        $ini = Join-Path $TestDrive 'setup3.ini'
        Set-Content $ini -Value "[Product]`r`nCmdLine=LANG_LIST=`"All`"`r`nmsi=AcroRead.msi"
        Set-AdobeSetupIniCommandLine -Path $ini -CommandLine ''
        $lines = Get-Content $ini
        @($lines | Where-Object { $_ -like 'CmdLine=*' }).Count | Should -Be 0
        $lines | Should -Contain 'msi=AcroRead.msi'
    }
    It 'creates the section when the file has none' {
        $ini = Join-Path $TestDrive 'setup4.ini'
        Set-Content $ini -Value "[Startup]`r`nRequireMSI=3.0"
        Set-AdobeSetupIniCommandLine -Path $ini -CommandLine 'LANG_LIST="en_US,fr_FR" SUPPRESSLANGSELECTION=YES'
        $lines = Get-Content $ini
        $lines | Should -Contain '[Product]'
        $lines | Should -Contain 'CmdLine=LANG_LIST="en_US,fr_FR" SUPPRESSLANGSELECTION=YES'
    }
}

Describe 'Adobe Reader preferences round trip' {
    BeforeAll {
        $gt = $null; $ge = $null
        $guiAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$gt, [ref]$ge)
        foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted') {
            $fn = $guiAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            . ([scriptblock]::Create($fn.Extent.Text))
        }
        $script:prefPath = Join-Path $TestDrive 'preferences.json'
        function Get-PreferencesPath { $script:prefPath }
    }
    It 'defaults to English with no languages' {
        Set-Content $script:prefPath -Value '{}'
        $p = Read-Preferences
        $p.AdobeReaderInstallOptions.Edition | Should -Be 'English'
        @($p.AdobeReaderInstallOptions.Languages).Count | Should -Be 0
    }
    It 'loads a stored MUI selection and drops malformed codes' {
        Set-Content $script:prefPath -Value '{ "AdobeReaderInstallOptions": { "Edition": "MUI", "Languages": ["de_DE", "bad", "All"] } }'
        $p = Read-Preferences
        $p.AdobeReaderInstallOptions.Edition | Should -Be 'MUI'
        $p.AdobeReaderInstallOptions.Languages | Should -Be @('de_DE', 'All')
    }
}
