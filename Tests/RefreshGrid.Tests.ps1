BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-RefreshGrid' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    function Update-GridFilter { }
    function Read-PackagerHistory { @{ 'package-chrome' = @{ LastKnownVersion = '130.0'; LastChecked = 'yesterday' } } }
    function Get-Packagers {
        param($Root)
        @(
            [pscustomobject]@{ Script = 'package-chrome.ps1'; Vendor = 'Google'; Application = 'Chrome'; Status = 'Ready'; CMName = 'Google Chrome'; FullPath = 'x'; VendorUrl = ''; Description = '' }
            [pscustomobject]@{ Script = 'package-firefox.ps1'; Vendor = 'Mozilla'; Application = 'Firefox'; Status = $script:FirefoxStatus; CMName = 'Mozilla Firefox'; FullPath = 'y'; VendorUrl = ''; Description = '' }
        )
    }
    $script:PackagersRoot = 'unused'
    $script:txtStatus = [pscustomobject]@{ Text = '' }
    $script:Prefs = [pscustomobject]@{ HiddenApplications = @() }
}

Describe 'Grid refresh keeps session results' {
    BeforeEach {
        $script:FirefoxStatus = 'Ready'
        $script:PackagerData = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Invoke-RefreshGrid
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.Selected = $true
        $chrome.CurrentVersion = '129.0'
        $chrome.LatestVersion = '131.0'
        $chrome.Status = 'Update available'
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $firefox.CurrentVersion = '140.0'
        $firefox.Status = 'Up to date'
    }

    It 'keeps ConfigMgr versions, compare results, latest versions and selections' {
        Invoke-RefreshGrid
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CurrentVersion | Should -Be '129.0'
        $chrome.LatestVersion | Should -Be '131.0'
        $chrome.Status | Should -Be 'Update available'
        $chrome.Selected | Should -BeTrue
    }

    It 'clears ConfigMgr versions and compare results when the site changed' {
        Invoke-RefreshGrid -DiscardSiteResults
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CurrentVersion | Should -Be ''
        $chrome.Status | Should -Be 'Ready'
        $chrome.LatestVersion | Should -Be '131.0'
        $chrome.Selected | Should -BeTrue
    }

    It 'shows a new read error instead of the kept result' {
        $script:FirefoxStatus = 'Read error: bad header'
        Invoke-RefreshGrid
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $firefox.Status | Should -Be 'Read error: bad header'
        $firefox.CurrentVersion | Should -Be ''
    }
}
