BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Invoke-RefreshGrid', 'Get-SiteIdentityKey', 'Clear-SucceededSelection', 'Get-CveGridCells', 'Select-OnlyUpdateAvailable', 'Compare-SemVer') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    function Update-GridFilter { }
    function Read-PackagerHistory {
        @{
            'package-chrome' = @{
                LastKnownVersion = $script:ChromeLatest; LastChecked = 'yesterday'
                LastCmVersion = '129.0'; LastCmSite = 'MCM'; LastCmChecked = 'yesterday'
                VulnSummary = [pscustomobject]@{ QueriedAt = '2026-10-11T12:00:00Z'; From = '129.0'; To = '130.0'; Source = 'NVD'; Count = 4; MaxCvss = 8.8; MaxSeverity = 'HIGH'; KevCount = 1; MaxEpss = 0.3; Error = $null; FixedInLatestCount = 1; FixedEarlierCount = 3; FixedUnknownCount = 0; NewestPublished = '2026-09-30'; LastNvdUpdate = '2026-10-08' }
            }
            'package-firefox' = @{ LastCmVersion = '140.0'; LastCmSite = 'OLD' }
        }
    }
    function Get-Packagers {
        param($Root)
        @(
            [pscustomobject]@{ Script = 'package-chrome.ps1'; Vendor = 'Google'; Application = 'Chrome'; Status = 'Ready'; CMName = 'Google Chrome'; FullPath = 'x'; VendorUrl = ''; Description = ''; CPE = 'cpe:2.3:a:google:chrome:*:*:*:*:*:*:*:*'; SecurityAdvisoryUrl = 'https://chromereleases.googleblog.com/' }
            # Get-Packagers always sets CPE, empty when the header has none.
            [pscustomobject]@{ Script = 'package-firefox.ps1'; Vendor = 'Mozilla'; Application = 'Firefox'; Status = $script:FirefoxStatus; CMName = 'Mozilla Firefox'; FullPath = 'y'; VendorUrl = ''; Description = ''; CPE = ''; SecurityAdvisoryUrl = '' }
        )
    }
    $script:PackagersRoot = 'unused'
    $script:txtStatus = [pscustomobject]@{ Text = '' }
}

Describe 'Grid refresh keeps session results' {
    BeforeEach {
        $script:FirefoxStatus = 'Ready'
        $script:ChromeLatest = '130.0'
        $script:Prefs = [pscustomobject]@{ HiddenApplications = @(); SiteCode = 'MCM' }
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

    It 'preloads the ConfigMgr version and the CVE cells from the history' {
        $script:GridSessionRows = $null
        $script:PackagerData = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Invoke-RefreshGrid
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CurrentVersion | Should -Be '129.0'
        $chrome.Cpe | Should -Be 'cpe:2.3:a:google:chrome:*:*:*:*:*:*:*:*'
        $chrome.CveCount | Should -Be '4'
        $chrome.CveCvss | Should -Be '8.8 HIGH'
        $chrome.CveKev | Should -Be 'Yes'
        $chrome.CveEpss | Should -Match '^30'
        $chrome.CveChecked | Should -Be '2026-10-11T12:00:00Z'
        $chrome.CveFixed | Should -Be '1 of 4'
        $chrome.CveNewest | Should -Be '2026-09-30'
        $chrome.CveUpdated | Should -Be '2026-10-08'
        $chrome.Status | Should -Be 'Update available (4 CVEs, 1 in KEV)'
        $chrome.AdvisoryUrl | Should -Be 'https://chromereleases.googleblog.com/'
        ($script:PackagerData | Where-Object Script -eq 'package-firefox.ps1').AdvisoryUrl | Should -Be ''
    }

    It 'leaves the ConfigMgr version empty when the history came from another site' {
        $script:GridSessionRows = $null
        $script:PackagerData = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Invoke-RefreshGrid
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $firefox.CurrentVersion | Should -Be ''
    }

    It 'marks a packager without a CPE header' {
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $firefox.Cpe | Should -Be ''
        $firefox.CveCount | Should -Be 'CPE missing'
        $firefox.CveCvss | Should -Be ''
    }

    It 'clears the CVE cells with the site results' {
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CveCount = '7'; $chrome.CveCvss = '9.8 CRITICAL'; $chrome.CveKev = 'Yes'; $chrome.CveChecked = 'now'
        Invoke-RefreshGrid -DiscardSiteResults
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.CveCount | Should -Be ''
        $chrome.CveCvss | Should -Be ''
        $chrome.CveKev | Should -Be ''
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $firefox.CveCount | Should -Be 'CPE missing'
    }

    It 'clears the CVE cells when the stored summary is for other versions' {
        # The history summary covers 129.0 to 130.0; the vendor now reports 131.0.
        $script:ChromeLatest = '131.0'
        $script:GridSessionRows = $null
        $script:PackagerData = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        Invoke-RefreshGrid
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $chrome.LatestVersion | Should -Be '131.0'
        $chrome.CveCount | Should -Be ''
        $chrome.CveCvss | Should -Be ''
        $chrome.CveKev | Should -Be ''
        $chrome.CveChecked | Should -Be ''
    }

    It 'selects rows whose status carries the CVE count' {
        $chrome = $script:PackagerData | Where-Object Script -eq 'package-chrome.ps1'
        $firefox = $script:PackagerData | Where-Object Script -eq 'package-firefox.ps1'
        $chrome.Status = 'Update available (3 CVEs, 1 in KEV)'
        $firefox.Status = 'Up to date'
        $script:dataGrid = [pscustomobject]@{ ItemsSource = @($chrome, $firefox) }
        Select-OnlyUpdateAvailable
        $chrome.Selected | Should -BeTrue
        $firefox.Selected | Should -BeFalse
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
    It 'is not applied after a Stage run, so Stage then Publish keeps the selection' {
        $source = Get-Content (Join-Path $PSScriptRoot '..\start-apppackager.ps1') -Raw
        $source | Should -Match 'if \(\$doneState\.Succeeded -and \$script:BgOperation -ne ''Stage''\) \{\s+Clear-SucceededSelection'
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
