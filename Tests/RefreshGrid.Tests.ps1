BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Invoke-RefreshGrid', 'Get-SiteIdentityKey', 'Clear-SucceededSelection') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        . ([scriptblock]::Create($fn.Extent.Text))
    }

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
}

Describe 'Grid refresh keeps session results' {
    BeforeEach {
        $script:FirefoxStatus = 'Ready'
        $script:Prefs = [pscustomobject]@{ HiddenApplications = @() }
        $script:GridSessionRows = $null
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

    It 'keeps the same row object so background writes after a rebuild stay visible' {
        $before = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        Invoke-RefreshGrid
        $before.CurrentVersion = '130.5'
        $after = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        [object]::ReferenceEquals($before, $after) | Should -BeTrue
        $after.CurrentVersion | Should -Be '130.5'
    }

    It 'keeps both version columns for a row hidden and shown again' {
        $script:Prefs.HiddenApplications = @('package-chrome.ps1')
        Invoke-RefreshGrid
        @($script:PackagerData | Where-Object Script -eq 'package-chrome.ps1').Count | Should -Be 0
        $script:Prefs.HiddenApplications = @()
        Invoke-RefreshGrid
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CurrentVersion | Should -Be '129.0'
        $chrome.LatestVersion | Should -Be '131.0'
        $chrome.Selected | Should -BeFalse
    }
}

Describe 'Site identity for the Options site-change check' {
    It 'treats whitespace and case differences as the same site' {
        (Get-SiteIdentityKey -SiteCode ' mcm ' -ProviderMachineName 'CM01.contoso.com ') |
            Should -BeExactly (Get-SiteIdentityKey -SiteCode 'MCM' -ProviderMachineName 'cm01.contoso.com')
    }
    It 'treats a null and an empty provider as the same site' {
        (Get-SiteIdentityKey -SiteCode 'MCM' -ProviderMachineName $null) | Should -BeExactly (Get-SiteIdentityKey -SiteCode 'MCM' -ProviderMachineName '')
    }
    It 'reports a different site code as a change' {
        (Get-SiteIdentityKey -SiteCode 'MCM' -ProviderMachineName 'cm01') | Should -Not -Be (Get-SiteIdentityKey -SiteCode 'PS1' -ProviderMachineName 'cm01')
    }
}

Describe 'Selection after a Stage or Package run' {
    It 'unchecks succeeded rows and keeps failed rows checked' {
        $rows = @(
            [pscustomobject]@{ Script = 'package-a.ps1'; Selected = $true }
            [pscustomobject]@{ Script = 'package-b.ps1'; Selected = $true }
            [pscustomobject]@{ Script = 'package-c.ps1'; Selected = $false }
        )
        Clear-SucceededSelection -Rows $rows -Scripts @('package-a.ps1')
        $rows[0].Selected | Should -BeFalse
        $rows[1].Selected | Should -BeTrue
        $rows[2].Selected | Should -BeFalse
    }
    It 'accepts an observable collection and an empty success list' {
        $data = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        $data.Add([pscustomobject]@{ Script = 'package-a.ps1'; Selected = $true })
        Clear-SucceededSelection -Rows $data -Scripts @()
        $data[0].Selected | Should -BeTrue
        Clear-SucceededSelection -Rows $data -Scripts @('PACKAGE-A.ps1')
        $data[0].Selected | Should -BeFalse
    }
}
