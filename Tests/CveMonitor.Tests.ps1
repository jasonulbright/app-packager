#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Headless CVE monitor and the CVE helpers of the grid, from
    start-apppackager.ps1.

.DESCRIPTION
    The functions are lifted out of the application script by AST and run with
    stubbed vendor, ConfigMgr and CVE lookups. The stub lookup builds its
    result with the vulnerability module and saves it into the result cache.
    The report writers, the result cache and the history file are the real
    ones; LOCALAPPDATA is redirected into the test drive.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerVulnerability.psd1') -Force
    $script:VulnModule = Get-Module -Name AppPackagerVulnerability

    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $script:Ast = $ast
    foreach ($name in 'Invoke-CveMonitor', 'Get-CveGridCells', 'Get-CveCachedResult', 'Find-CveSessionResult', 'Get-CveReportRows', 'Compare-SemVer', 'Remove-ExpiredCveReports', 'ConvertTo-CveVersionKey', 'ConvertTo-CvePairKey', 'Clear-StaleCveCells', 'Set-CveCellsPair', 'Get-CveSkipReason') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # The Check CVEs step of the background run is one clause of the switch
    # over the operation. The clause runs here inside the same loop shape:
    # its "continue" ends the switch for the row, as in the application.
    $stepClauses = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] }, $true) |
            ForEach-Object { $_.Clauses } |
            Where-Object { $_.Item1.Extent.Text -eq "'CheckCve'" -and $_.Item2.Extent.Text -match 'Get-ProductVulnerabilities' })
    if ($stepClauses.Count -ne 1) { throw ('expected one CheckCve step, found {0}' -f $stepClauses.Count) }
    $script:CheckCveStepAst = $stepClauses[0].Item2
    $script:CheckCveStep = [scriptblock]::Create(
        'param($Rows, $Ctx, $State, $counts) $i = 0; $n = @($Rows).Count; foreach ($row in @($Rows)) { $i++; $app = [string]$row.Application; $scrName = [string]$row.Script; $baseName = [System.IO.Path]::GetFileNameWithoutExtension($scrName); switch (''CheckCve'') { ''CheckCve'' ' +
        $script:CheckCveStepAst.Extent.Text + ' } }')

    $script:OrigLocalAppData = $env:LOCALAPPDATA
    $env:LOCALAPPDATA = Join-Path $TestDrive 'LocalAppData'
    $script:DataFolder = Join-Path $env:LOCALAPPDATA 'AppPackager'
    New-Item -ItemType Directory -Path $script:DataFolder -Force | Out-Null

    $script:AlphaCpe = 'cpe:2.3:a:v:alpha:*:*:*:*:*:*:*:*'
    $script:BetaCpe = 'cpe:2.3:a:v:beta:*:*:*:*:*:*:*:*'
    $script:EpsilonCpe = 'cpe:2.3:a:v:epsilon:*:*:*:*:*:*:*:*'
    $script:GridCpe = 'cpe:2.3:a:v:p:*:*:*:*:*:*:*:*'
    $script:StampPattern = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'

    $script:ResetState = {
        $script:Log = @(); $script:LatestCalls = @(); $script:CmCalls = @(); $script:CveCalls = @(); $script:KevCalls = 0; $script:Notified = @()
        $script:CveFail = $false; $script:CveNoCache = $false; $script:CveEmpty = $false; $script:CveUnknown = $false; $script:CveThrow = ''
        $script:CveSessionResults = $null
        $script:Reports = Join-Path $TestDrive ('reports-' + [guid]::NewGuid().ToString('N'))
        foreach ($name in 'app-history.json', 'vuln-results') {
            $path = Join-Path $script:DataFolder $name
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
        }
    }

    $script:NewRecord = {
        param([string]$Id, [double]$Score, [string]$Severity, [string]$Status, [switch]$Kev)
        $rec = New-VulnerabilityRecord -Id $Id -Score $Score -Severity $Severity -Published '2025-01-01T00:00:00.000' -LastModified '2025-02-01T00:00:00.000' -Summary ('Summary of ' + $Id) -Source 'NVD'
        $rec.Status = $Status
        $rec.Kev = [bool]$Kev
        $rec.Links = (Get-VulnerabilityLinks -CveId $Id -Kev:$Kev).Cve
        $rec
    }

    # A result as a lookup returns it. The module makes the result, so it has
    # every field of the current layout; the counts follow the records.
    $script:NewResult = {
        param([string]$Cpe, [string]$From = '', [string]$To = '', [string]$Advisory = '', [object[]]$Cves = @(), [object[]]$Unversioned = @(), [string]$QueriedAt = '')
        $r = & $script:VulnModule { param($c, $f, $t, $a) New-VulnerabilityResult -Cpe $c -From $f -To $t -VendorAdvisoryUrl $a } $Cpe $From $To $Advisory
        $r.Source = 'NVD'
        if ($QueriedAt) { $r.QueriedAt = $QueriedAt }
        $r.Cves = @($Cves)
        $r.UnversionedCves = @($Unversioned)
        $r.Count = @($Cves).Count
        $r.FixedCount = @($Cves | Where-Object { $_.Status -eq 'Fixed' }).Count
        $r.OpenCount = @($Cves | Where-Object { $_.Status -eq 'Open' }).Count
        $r.UnversionedCount = @($Unversioned).Count
        $r.KevCount = @($Cves | Where-Object { $_.Kev }).Count
        $r.FixedKevCount = @($Cves | Where-Object { $_.Kev -and $_.Status -eq 'Fixed' }).Count
        $top = @($Cves | Sort-Object -Property Score -Descending | Select-Object -First 1)
        if ($top.Count -gt 0) {
            $r.MaxCvss = [double]$top[0].Score
            $r.MaxSeverity = [string]$top[0].Severity
            $r.NewestPublished = '2025-01-01'
            $r.LastNvdUpdate = '2025-02-01'
        }
        $r
    }

    # A history summary of the current layout for one CPE and one version pair.
    $script:NewSummary = {
        param([hashtable]$Set = @{})
        $s = @{
            Schema = 2; QueriedAt = '2026-11-01T00:00:00Z'; Cpe = $script:GridCpe; From = '1.0'; To = '2.0'; Mode = 'Update'; Source = 'NVD'
            Count = 3; FixedCount = 1; OpenCount = 2; UnversionedCount = 0; MaxCvss = 7.5; MaxSeverity = 'HIGH'; KevCount = 0; FixedKevCount = 0
            MaxEpss = 0.25; NewestPublished = '2026-09-30'; LastNvdUpdate = '2026-10-08'; ProductRecords = 12; Error = $null
        }
        foreach ($k in $Set.Keys) { $s[$k] = $Set[$k] }
        $s
    }

    function Write-Log { param([string]$Message, [string]$Level = 'INFO') $script:Log += $Message }
    function Get-Packagers {
        param($Root)
        @(
            [pscustomobject]@{ Script = 'package-alpha.ps1'; FullPath = 'C:\x\package-alpha.ps1'; Application = 'Alpha'; Vendor = 'V'; CMName = 'Alpha'; Status = 'Ready'; CPE = $script:AlphaCpe; SecurityAdvisoryUrl = 'https://v.example.test/security' }
            [pscustomobject]@{ Script = 'package-beta.ps1'; FullPath = 'C:\x\package-beta.ps1'; Application = 'Beta'; Vendor = 'V'; CMName = 'Beta'; Status = 'Ready'; CPE = $script:BetaCpe; SecurityAdvisoryUrl = '' }
            [pscustomobject]@{ Script = 'package-gamma.ps1'; FullPath = 'C:\x\package-gamma.ps1'; Application = 'Gamma'; Vendor = 'V'; CMName = 'Gamma'; Status = 'Ready'; CPE = ''; SecurityAdvisoryUrl = '' }
            [pscustomobject]@{ Script = 'package-delta.ps1'; FullPath = 'C:\x\package-delta.ps1'; Application = 'Delta'; Vendor = 'V'; CMName = 'Delta'; Status = 'Ready'; CPE = 'cpe:2.3:a:v:delta:*:*:*:*:*:*:*:*'; SecurityAdvisoryUrl = '' }
            [pscustomobject]@{ Script = 'package-epsilon.ps1'; FullPath = 'C:\x\package-epsilon.ps1'; Application = 'Epsilon'; Vendor = 'V'; CMName = 'Epsilon'; Status = 'Ready'; CPE = $script:EpsilonCpe; SecurityAdvisoryUrl = '' }
        )
    }
    # Alpha has an update. Beta is up to date. Gamma has an update and no CPE.
    # The vendor lookup of Delta fails. ConfigMgr does not hold Epsilon.
    function Invoke-PackagerGetLatestVersion {
        param($PackagerPath, $SiteCode, $FileServerPath, $DownloadRoot, $M365Channel, $M365DeployMode)
        $script:LatestCalls += $PackagerPath
        switch -Wildcard ($PackagerPath) {
            '*alpha*' { '2.0' }
            '*beta*'  { '5.0' }
            '*gamma*' { '1.1' }
            '*delta*' { throw 'vendor page unreachable' }
            '*epsilon*' { '3.0' }
        }
    }
    function Get-MecmCurrentVersionByCMName {
        param($SiteCode, $ProviderMachineName, $CMName, [switch]$ExactMatch)
        $script:CmCalls += $CMName
        switch ($CMName) {
            'Alpha' { [pscustomobject]@{ Found = $true; SoftwareVersion = '1.0'; MatchCount = 1 } }
            'Beta'  { [pscustomobject]@{ Found = $true; SoftwareVersion = '5.0'; MatchCount = 1 } }
            'Gamma' { [pscustomobject]@{ Found = $true; SoftwareVersion = '1.0'; MatchCount = 1 } }
            default { [pscustomobject]@{ Found = $false; SoftwareVersion = ''; MatchCount = 0 } }
        }
    }
    function Get-KevCatalog { $script:KevCalls++; @{} }
    # The records follow the mode of the version pair:
    #   Update    3 CVEs: 2 fixed (1 in KEV), 1 open; 1 record with no version
    #   Deployed  2 CVEs, both open (1 in KEV)
    #   Latest    1 CVE, open
    function Get-ProductVulnerabilities {
        param($Cpe, $CurrentVersion, $LatestVersion, $ApiKey, $CacheFolder, $CacheTtlMinutes, [switch]$Force, $VendorAdvisoryUrl, $KevCatalog, [switch]$SkipEnrichment)
        $script:CveCalls += [pscustomobject]@{ Cpe = $Cpe; From = $CurrentVersion; To = $LatestVersion; Advisory = $VendorAdvisoryUrl; Key = $ApiKey; Force = [bool]$Force }
        if ($script:CveThrow -and $Cpe -like $script:CveThrow) { throw 'lookup stopped' }
        $from = ConvertTo-VulnerabilityVersion -Version $CurrentVersion
        $to = ConvertTo-VulnerabilityVersion -Version $LatestVersion
        $cves = @(); $unversioned = @()
        if (-not $script:CveFail -and -not $script:CveEmpty -and -not $script:CveUnknown) {
            switch (Get-VulnerabilityLookupMode -From $from -To $to) {
                'Update' {
                    $cves = @(
                        (& $script:NewRecord -Id 'CVE-2025-0001' -Score 7.5 -Severity 'HIGH' -Status 'Fixed' -Kev)
                        (& $script:NewRecord -Id 'CVE-2025-0003' -Score 9.1 -Severity 'CRITICAL' -Status 'Open')
                        (& $script:NewRecord -Id 'CVE-2025-0002' -Score 5.0 -Severity 'MEDIUM' -Status 'Fixed')
                    )
                    $unversioned = @(& $script:NewRecord -Id 'CVE-2025-0009' -Score 4.0 -Severity 'MEDIUM' -Status 'Unversioned')
                }
                'Deployed' {
                    $cves = @(
                        (& $script:NewRecord -Id 'CVE-2025-0004' -Score 8.1 -Severity 'HIGH' -Status 'Open' -Kev)
                        (& $script:NewRecord -Id 'CVE-2025-0005' -Score 4.3 -Severity 'MEDIUM' -Status 'Open')
                    )
                }
                default { $cves = @(& $script:NewRecord -Id 'CVE-2025-0006' -Score 6.5 -Severity 'MEDIUM' -Status 'Open') }
            }
        }
        $r = & $script:NewResult -Cpe $Cpe -From $from -To $to -Advisory $VendorAdvisoryUrl -Cves $cves -Unversioned $unversioned
        # NVD answered and holds no record under the product name.
        if ($script:CveUnknown) { $r.ProductRecords = 0 }
        if ($script:CveFail) {
            $r.Source = ''
            $r.Error = 'NVD and Shodan CVEDB both failed: no answer'
            return $r
        }
        # A lookup saves its result. A failed write leaves the earlier file.
        if (-not $script:CveNoCache) { Save-VulnerabilityCachedResult -Result $r }
        return $r
    }
    function Send-VulnerabilityNotification {
        param($Paths, $Summary, $WebhookUrl, $DropFolder)
        $script:Notified += [pscustomobject]@{ Summary = $Summary; WebhookUrl = $WebhookUrl; DropFolder = $DropFolder }
        [pscustomobject]@{ DropCopied = 0; WebhookStatus = 200 }
    }
}

AfterAll {
    $env:LOCALAPPDATA = $script:OrigLocalAppData
}

Describe 'Invoke-CveMonitor' {
    BeforeEach { & $script:ResetState }

    It 'runs latest, ConfigMgr and CVE lookups and writes every report format' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -NvdApiKey 'k' -WebhookUrl 'https://hooks.example.test/x'
        $s.Total | Should -Be 5
        # Alpha and Gamma have an update. Alpha and Beta have a CPE and a ConfigMgr version.
        $s.Updates | Should -Be 2
        $s.LookedUp | Should -Be 2
        $s.Cached | Should -Be 0
        $s.Failed | Should -Be 1
        $script:LatestCalls.Count | Should -Be 5
        # ConfigMgr is not queried for Delta: its vendor lookup failed.
        $script:CmCalls.Count | Should -Be 4
        $script:CveCalls.Count | Should -Be 2
        $script:KevCalls | Should -Be 1
        $script:CveCalls[0].Cpe | Should -Be $script:AlphaCpe
        $script:CveCalls[0].From | Should -Be '1.0'
        $script:CveCalls[0].To | Should -Be '2.0'
        $script:CveCalls[0].Advisory | Should -Be 'https://v.example.test/security'
        $script:CveCalls[0].Key | Should -Be 'k'
        $script:CveCalls[1].Cpe | Should -Be $script:BetaCpe
        foreach ($p in 'Html', 'ApplicationsCsv', 'CvesCsv', 'Json', 'Xlsx', 'Markdown') { Test-Path -LiteralPath $s.Paths.$p | Should -BeTrue }
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        $alpha = $apps | Where-Object Application -eq 'Alpha'
        $alpha.CVEs | Should -Be '3'
        $alpha.FixedByUpdate | Should -Be '2'
        $alpha.Open | Should -Be '1'
        $alpha.Unversioned | Should -Be '1'
        $alpha.Status | Should -Be 'Update available (fixes 2 of 3 CVEs, 1 in KEV)'
        $alpha.VendorAdvisory | Should -Be 'https://v.example.test/security'
        ($apps | Where-Object Application -eq 'Beta').Status | Should -Be 'Up to date'
        ($apps | Where-Object Application -eq 'Gamma').Note | Should -Be 'CPE missing'
        ($apps | Where-Object Application -eq 'Delta').Status | Should -Be 'Error'
        ($apps | Where-Object Application -eq 'Epsilon').Status | Should -Be 'Not found in ConfigMgr'
        $script:Notified.Count | Should -Be 1
        $script:Notified[0].Summary.Updates | Should -Be 2
    }

    It 'returns the totals of the affecting, the fixed and the open CVEs' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports
        @($s.PSObject.Properties.Name) | Should -Be @('Total', 'Updates', 'LookedUp', 'Cached', 'Cves', 'Fixed', 'Open', 'Kev', 'Failed', 'Started', 'Ended', 'Computer', 'Paths')
        # Alpha: 3 CVEs, 2 fixed, 1 open, 1 in KEV. Beta: 2 CVEs, both open, 1 in KEV.
        $s.Cves | Should -Be 5
        $s.Fixed | Should -Be 2
        $s.Open | Should -Be 3
        $s.Kev | Should -Be 2
        $s.Started | Should -Match $script:StampPattern
        $s.Ended | Should -Match $script:StampPattern
        $s.Computer | Should -Be $env:COMPUTERNAME
        ($script:Log | Where-Object { $_ -like '*Monitor finished*' }) | Should -Match '5 application\(s\), 2 with an update, 2 looked up, 0 from the cache, 5 CVE\(s\) in the ConfigMgr versions, 2 fixed by updating, 3 open, 2 in KEV, 1 failed'
    }

    It 'shows what the update fixes in the status of an application with an update' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.Updates | Should -Be 1
        $s.LookedUp | Should -Be 1
        $s.Cves | Should -Be 3
        $s.Fixed | Should -Be 2
        $s.Open | Should -Be 1
        $s.Kev | Should -Be 1
        $alpha = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha'
        $alpha.Status | Should -Be 'Update available (fixes 2 of 3 CVEs, 1 in KEV)'
        $alpha.ExaminedVersion | Should -Be '1.0'
        $alpha.Note | Should -Be ''
    }

    It 'says no CVE in the status when no CVE affects the ConfigMgr version' {
        $script:CveEmpty = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.LookedUp | Should -Be 1
        $s.Cves | Should -Be 0
        (@(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha').Status | Should -Be 'Update available (no CVE)'
    }

    It 'says that NVD has no record of the product, and not that the version has no CVE' {
        $script:CveUnknown = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.LookedUp | Should -Be 1
        $s.Cves | Should -Be 0
        $alpha = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha'
        $alpha.Status | Should -Be 'Update available (product not in NVD)'
        $alpha.CVEs | Should -Be '0'
        $alpha.NVDRecords | Should -Be '0'
        $alpha.Note | Should -Be 'NVD holds no CVE record under the product name of the CPE header. The product has no known CVE, or NVD files it under another name.'
        @($script:Log | Where-Object { $_ -like '*Alpha: NVD holds no CVE record under the product name of the CPE header.*(NVD)' }).Count | Should -Be 1
        # The history keeps the number, so the grid of the next session says the same.
        $stored = (Read-PackagerHistory)['package-alpha'].VulnSummary
        (Get-CveGridCells -Summary $stored -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be 'Not in NVD'
    }

    It 'looks up an up-to-date application and counts its open CVEs' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta'
        $s.Updates | Should -Be 0
        $s.LookedUp | Should -Be 1
        $script:CveCalls.Count | Should -Be 1
        $script:CveCalls[0].From | Should -Be '5.0'
        $script:CveCalls[0].To | Should -Be '5.0'
        $s.Cves | Should -Be 2
        $s.Fixed | Should -Be 0
        $s.Open | Should -Be 2
        $s.Kev | Should -Be 1
        $beta = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Beta'
        $beta.Status | Should -Be 'Up to date'
        $beta.CVEs | Should -Be '2'
        $beta.FixedByUpdate | Should -Be ''
        $beta.Open | Should -Be '2'
        (Read-PackagerHistory)['package-beta'].VulnSummary.Mode | Should -Be 'Deployed'
    }

    It 'looks up an application whose ConfigMgr version is newer than the latest version' {
        Update-PackagerHistory -PackagerName 'package-alpha' -Event Checked -Version '1.0' -Result Updated
        Update-PackagerHistory -PackagerName 'package-alpha' -Event CmChecked -Version '2.0' -Data @{ SiteCode = 'MCM' }
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipLatest -SkipMecm
        $s.Updates | Should -Be 0
        $s.LookedUp | Should -Be 1
        $s.Open | Should -Be 2
        (@(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha').Status | Should -Be 'Current newer'
    }

    It 'does not look up an application that ConfigMgr does not hold' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-epsilon'
        $script:CmCalls | Should -Be @('Epsilon')
        $script:CveCalls.Count | Should -Be 0
        $script:KevCalls | Should -Be 0
        $s.LookedUp | Should -Be 0
        $s.Failed | Should -Be 0
        $epsilon = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Epsilon'
        $epsilon.Status | Should -Be 'Not found in ConfigMgr'
        $epsilon.CVEs | Should -Be ''
        $epsilon.Note | Should -Be 'Not checked'
        (Read-PackagerHistory)['package-epsilon'].VulnSummary | Should -BeNullOrEmpty
    }

    It 'names the missing CPE and makes no lookup' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-gamma'
        $s.Updates | Should -Be 1
        $s.LookedUp | Should -Be 0
        $s.Failed | Should -Be 0
        $script:CveCalls.Count | Should -Be 0
        $gamma = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Gamma'
        $gamma.Status | Should -Be 'Update available'
        $gamma.Note | Should -Be 'CPE missing'
        @($script:Log | Where-Object { $_ -like '*Gamma: CPE missing*' }).Count | Should -Be 1
    }

    It 'records the versions and the summary in the history file' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports
        $h = Read-PackagerHistory
        ([string]$h['package-alpha'].LastKnownVersion) | Should -Be '2.0'
        ([string]$h['package-alpha'].LastCmVersion) | Should -Be '1.0'
        ([string]$h['package-alpha'].LastCmSite) | Should -Be 'MCM'
        $summary = $h['package-alpha'].VulnSummary
        $summary.Schema | Should -Be 2
        $summary.Cpe | Should -Be $script:AlphaCpe
        $summary.From | Should -Be '1.0'
        $summary.To | Should -Be '2.0'
        $summary.Mode | Should -Be 'Update'
        $summary.Count | Should -Be 3
        $summary.FixedCount | Should -Be 2
        $summary.OpenCount | Should -Be 1
        $summary.UnversionedCount | Should -Be 1
        $summary.KevCount | Should -Be 1
        $summary.FixedKevCount | Should -Be 1
        ([string]$h['package-beta'].LastCmVersion) | Should -Be '5.0'
        $h['package-beta'].VulnSummary.Schema | Should -Be 2
        $h['package-beta'].VulnSummary.Count | Should -Be 2
        $h['package-gamma'].VulnSummary | Should -BeNullOrEmpty
        $h.ContainsKey('package-delta') | Should -BeTrue
        ([string]$h['package-delta'].LastResult) | Should -Be 'Failed'
        $h['package-delta'].LastKnownVersion | Should -BeNullOrEmpty
        ([string]$h['package-epsilon'].LastCmVersion) | Should -Be ''
    }

    It 'stores a summary that the grid shows for the same CPE and versions' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $stored = (Read-PackagerHistory)['package-alpha'].VulnSummary
        $cells = Get-CveGridCells -Summary $stored -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be '3'
        $cells.Fixed | Should -Be '2 of 3'
        $cells.Text | Should -Be 'fixes 2 of 3 CVEs, 1 in KEV'
        $cells.Kev | Should -Be 'Yes'
        # The history file comes back from JSON. The time stays the text the lookup wrote.
        $cells.Checked | Should -Match $script:StampPattern
        (Get-CveGridCells -Summary $stored -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '3.0').Count | Should -Be ''
    }

    It 'uses the history instead of the vendors and ConfigMgr with the skip switches' {
        Update-PackagerHistory -PackagerName 'package-alpha' -Event Checked -Version '3.0' -Result Updated
        Update-PackagerHistory -PackagerName 'package-alpha' -Event CmChecked -Version '1.5' -Data @{ SiteCode = 'MCM' }
        Update-PackagerHistory -PackagerName 'package-beta' -Event Checked -Version '5.0' -Result Updated
        Update-PackagerHistory -PackagerName 'package-beta' -Event CmChecked -Version '4.0' -Data @{ SiteCode = 'OLD' }
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -SkipLatest -SkipMecm
        $script:LatestCalls.Count | Should -Be 0
        $script:CmCalls.Count | Should -Be 0
        $s.Updates | Should -Be 1
        # Only Alpha has a ConfigMgr version from this site.
        $script:CveCalls.Count | Should -Be 1
        $script:CveCalls[0].From | Should -Be '1.5'
        $script:CveCalls[0].To | Should -Be '3.0'
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        # Beta's ConfigMgr version came from another site and is not used.
        ($apps | Where-Object Application -eq 'Beta').ConfigMgrVersion | Should -Be ''
        ($apps | Where-Object Application -eq 'Beta').Status | Should -Be 'Latest retrieved'
        ($apps | Where-Object Application -eq 'Beta').CVEs | Should -Be ''
    }

    It 'limits the run to -Apps and skips the lookup with -SkipCve' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-beta.ps1' -SkipCve
        $s.Total | Should -Be 2
        $s.Updates | Should -Be 1
        $s.LookedUp | Should -Be 0
        $s.Cves | Should -Be 0
        $script:CveCalls.Count | Should -Be 0
        $script:KevCalls | Should -Be 0
        $script:LatestCalls.Count | Should -Be 2
        (@(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha').Status | Should -Be 'Update available'
        (Read-PackagerHistory)['package-alpha'].VulnSummary | Should -BeNullOrEmpty
    }

    It 'notifies only with an update unless told otherwise' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta' -WebhookUrl 'https://hooks.example.test/x'
        # Beta has open CVEs and no update.
        $script:CveCalls.Count | Should -Be 1
        $script:Notified.Count | Should -Be 0
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta' -WebhookUrl 'https://hooks.example.test/x' -NotifyAlways
        $script:Notified.Count | Should -Be 1
        $script:Notified[0].Summary.Open | Should -Be 2
    }

    It 'sends no notification without a webhook and a drop folder' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -NotifyAlways
        $script:Notified.Count | Should -Be 0
    }

    It 'writes an empty report when -Apps matches nothing' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-nothing'
        $s.Total | Should -Be 0
        $s.Failed | Should -Be 0
        Test-Path -LiteralPath $s.Paths.Html | Should -BeTrue
        @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv).Count | Should -Be 0
    }

    It 'does not record a failed lookup over a good summary' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        (Read-PackagerHistory)['package-alpha'].VulnSummary.Count | Should -Be 3
        $script:CveFail = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.Failed | Should -Be 1
        $s.LookedUp | Should -Be 0
        $s.Cves | Should -Be 0
        $after = (Read-PackagerHistory)['package-alpha'].VulnSummary
        $after.Count | Should -Be 3
        $after.Error | Should -BeNullOrEmpty
        # The report shows the cached good result and says that the latest lookup failed.
        $alpha = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha'
        $alpha.CVEs | Should -Be '3'
        $alpha.FixedByUpdate | Should -Be '2'
        $alpha.Note | Should -Match '^Last lookup failed; this is the result from \d{4}-\d{2}-\d{2}T'
        $alpha.Status | Should -Be 'Update available'
    }

    It 'reports a failed lookup with no earlier result as an error' {
        $script:CveFail = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.Failed | Should -Be 1
        $s.LookedUp | Should -Be 0
        $s.Cached | Should -Be 0
        (Read-PackagerHistory)['package-alpha'].VulnSummary | Should -BeNullOrEmpty
        Read-VulnerabilityCachedResult -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '2.0' | Should -BeNullOrEmpty
        $alpha = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha'
        $alpha.CVEs | Should -Be ''
        $alpha.Note | Should -Be 'Error'
        $alpha.Status | Should -Be 'Update available'
    }

    It 'counts a lookup that throws as failed and continues with the next application' {
        $script:CveThrow = '*:alpha:*'
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-beta'
        $s.Failed | Should -Be 1
        $s.LookedUp | Should -Be 1
        $s.Cves | Should -Be 2
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').Note | Should -Be 'Error'
        ($apps | Where-Object Application -eq 'Beta').CVEs | Should -Be '2'
        (Read-PackagerHistory)['package-alpha'].VulnSummary | Should -BeNullOrEmpty
    }

    It 'prunes reports older than KeepReportDays' {
        New-Item -ItemType Directory -Path $script:Reports -Force | Out-Null
        $old = Join-Path $script:Reports 'cve-report-20200101-000000.html'
        [System.IO.File]::WriteAllText($old, 'x')
        (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-100)
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta' -SkipCve -KeepReportDays 30
        Test-Path -LiteralPath $old | Should -BeFalse
    }
}

Describe 'Get-CveReportRows marks a result that is not the current lookup' {
    BeforeAll {
        & $script:ResetState
        $script:NoteCpe = 'cpe:2.3:a:vendor:product:*:*:*:*:*:*:*:*'
        $script:OldStamp = [DateTime]::UtcNow.AddHours(-3).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $script:NewStamp = [DateTime]::UtcNow.AddHours(-2).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $script:NoteResult = & $script:NewResult -Cpe $script:NoteCpe -From '1.0' -To '2.0' -QueriedAt $script:OldStamp -Cves @(
            (& $script:NewRecord -Id 'CVE-2026-0001' -Score 7.5 -Severity 'HIGH' -Status 'Fixed')
            (& $script:NewRecord -Id 'CVE-2026-0002' -Score 5.0 -Severity 'MEDIUM' -Status 'Open')
        )
        Save-VulnerabilityCachedResult -Result $script:NoteResult
        $script:NoteRow = { param($count, $checked) [pscustomobject]@{ Application = 'App'; Vendor = 'V'; Script = 'package-app.ps1'; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available'; Cpe = $script:NoteCpe; CveCount = $count; CveChecked = $checked } }
        $script:NoteRows = Get-CveReportRows -Rows @((& $script:NoteRow 'Error' $script:OldStamp), (& $script:NoteRow '2' $script:NewStamp), (& $script:NoteRow '2' $script:OldStamp))
    }
    It 'keeps the last good result and says the latest lookup failed' {
        $script:NoteRows[0].Result.Count | Should -Be 2
        $script:NoteRows[0].Note | Should -Be "Last lookup failed; this is the result from $($script:OldStamp)"
    }
    It 'says when the grid holds a newer lookup than the cache' {
        $script:NoteRows[1].Note | Should -Match "^Result from $($script:OldStamp); the lookup of $($script:NewStamp) was not cached$"
    }
    It 'adds no note when the cache is the current lookup' {
        $script:NoteRows[2].Note | Should -Be ''
        $script:NoteRows[2].FromRun | Should -BeFalse
    }
    It 'marks a cached result when the caller asks for it' {
        $rows = Get-CveReportRows -Rows @((& $script:NoteRow '2' $script:OldStamp)) -MarkCached
        $rows[0].Note | Should -Be 'From the cache; not looked up in this run'
        $rows[0].FromRun | Should -BeFalse
    }
    It 'adds no note to a result of the run' {
        $rows = Get-CveReportRows -Rows @((& $script:NoteRow '2' $script:OldStamp)) -Results @{ 'package-app.ps1' = $script:NoteResult } -MarkCached
        $rows[0].Note | Should -Be ''
        $rows[0].FromRun | Should -BeTrue
        $rows[0].Result.Count | Should -Be 2
    }
    It 'says Not checked for a row with no result' {
        $other = & $script:NoteRow '' ''
        $other.LatestVersion = '3.0'
        $digits = & $script:NoteRow '7' ''
        $digits.LatestVersion = '3.0'
        $rows = Get-CveReportRows -Rows @($other, $digits) -MarkCached
        $rows[0].Result | Should -BeNullOrEmpty
        $rows[0].Note | Should -Be 'Not checked'
        $rows[1].Note | Should -Be 'Not checked'
    }
    It 'passes the reason of a row that was not looked up on as the note' {
        $noVersion = & $script:NoteRow 'No version' ''
        $noVersion.CurrentVersion = ''; $noVersion.LatestVersion = ''
        $noCpe = & $script:NoteRow 'CPE missing' ''
        $noCpe.Cpe = ''
        $failed = & $script:NoteRow 'Error' ''
        $failed.LatestVersion = '3.0'
        $rows = Get-CveReportRows -Rows @($noVersion, $noCpe, $failed)
        $rows[0].Note | Should -Be 'No version'
        $rows[1].Note | Should -Be 'CPE missing'
        $rows[2].Note | Should -Be 'Error'
        @($rows | Where-Object { $_.Result }).Count | Should -Be 0
    }
    It 'returns an array for no rows' {
        $rows = Get-CveReportRows -Rows @()
        ($rows -is [array]) | Should -BeTrue
        $rows.Count | Should -Be 0
    }
    It 'prints the note beside the count in the HTML, Markdown and CSV reports' {
        $out = Export-VulnerabilityReport -Rows $script:NoteRows -Folder (Join-Path $TestDrive 'note-report') -Name 'n1' -Formats @('Html', 'Markdown', 'Csv')
        $html = [System.IO.File]::ReadAllText($out.Html)
        $html | Should -Match "<td>2 <span class='note'>Last lookup failed; this is the result from $($script:OldStamp)</span></td>"
        $md = [System.IO.File]::ReadAllText($out.Markdown)
        $md | Should -Match "\| 2 \(Last lookup failed; this is the result from $($script:OldStamp)\) \|"
        $apps = @(Import-Csv -LiteralPath $out.ApplicationsCsv)
        @($apps | Where-Object Note -like 'Last lookup failed*').Count | Should -Be 1
        @($apps | Where-Object Note -eq '').Count | Should -Be 1
    }
}

Describe 'Report pruning deletes only what the writer produces' {
    It 'removes old report files by their exact names and leaves every other file' {
        $folder = Join-Path $TestDrive 'prune'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $mine = @('cve-report-20200101-000000.html', 'cve-report-20200101-000000.json', 'cve-report-20200101-000000.md', 'cve-report-20200101-000000.xlsx', 'cve-report-20200101-000000-applications.csv', 'cve-report-20200101-000000-cves.csv')
        $others = @('cve-report-notes.docx', 'cve-report-20200101-000000.txt', 'cve-report-20200101-000000.html.bak', 'cve-report-budget.xlsx', 'other.html')
        foreach ($n in $mine + $others) {
            $path = Join-Path $folder $n
            [System.IO.File]::WriteAllText($path, 'x')
            (Get-Item -LiteralPath $path).LastWriteTime = (Get-Date).AddDays(-100)
        }
        $recent = Join-Path $folder 'cve-report-20990101-000000.html'
        [System.IO.File]::WriteAllText($recent, 'x')
        Remove-ExpiredCveReports -Folder $folder -KeepDays 30
        foreach ($n in $mine) { Test-Path -LiteralPath (Join-Path $folder $n) | Should -BeFalse }
        foreach ($n in $others) { Test-Path -LiteralPath (Join-Path $folder $n) | Should -BeTrue }
        Test-Path -LiteralPath $recent | Should -BeTrue
    }
}

Describe 'The monitor exports its own results' {
    BeforeEach { & $script:ResetState }

    It 'reports the fresh result when the cache write failed, not an older entry' {
        $five = @(1..5 | ForEach-Object { & $script:NewRecord -Id ('CVE-2020-000' + $_) -Score 9.1 -Severity 'CRITICAL' -Status 'Fixed' })
        $old = & $script:NewResult -Cpe $script:AlphaCpe -From '1.0' -To '2.0' -QueriedAt ([DateTime]::UtcNow.AddHours(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')) -Cves $five
        Save-VulnerabilityCachedResult -Result $old
        $script:CveNoCache = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.LookedUp | Should -Be 1
        $s.Cached | Should -Be 0
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').CVEs | Should -Be '3'
        ($apps | Where-Object Application -eq 'Alpha').Note | Should -Be ''
        (Read-VulnerabilityCachedResult -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be 5
    }

    It 'marks a cached result for an application the run did not look up' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipCve -SkipLatest -SkipMecm
        $script:CveCalls.Count | Should -Be 1
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').CVEs | Should -Be '3'
        ($apps | Where-Object Application -eq 'Alpha').Note | Should -Be 'From the cache; not looked up in this run'
    }

    It 'does not report the cached result of other versions' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        # The history now holds a newer latest version than the cached lookup.
        Update-PackagerHistory -PackagerName 'package-alpha' -Event Checked -Version '2.5' -Result Updated
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipCve -SkipLatest -SkipMecm
        $s.Cached | Should -Be 0
        $alpha = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Alpha'
        $alpha.CVEs | Should -Be ''
        $alpha.Note | Should -Be 'Not checked'
    }

    It 'reports a cached result of the latest version for an application that ConfigMgr does not hold' {
        $latestOnly = & $script:NewResult -Cpe $script:EpsilonCpe -From '' -To '3.0' -Cves @(& $script:NewRecord -Id 'CVE-2025-0006' -Score 6.5 -Severity 'MEDIUM' -Status 'Open')
        Save-VulnerabilityCachedResult -Result $latestOnly
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-epsilon'
        $script:CveCalls.Count | Should -Be 0
        $s.LookedUp | Should -Be 0
        $s.Cached | Should -Be 1
        $epsilon = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv) | Where-Object Application -eq 'Epsilon'
        $epsilon.ExaminedVersion | Should -Be '3.0'
        $epsilon.CVEs | Should -Be '1'
        $epsilon.Note | Should -Be 'From the cache; not looked up in this run'
    }
}

Describe 'A stored summary belongs to one product' {
    It 'shows the count under the CPE it was looked up for' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Cpe = 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*' }) -Cpe 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*' -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be '3'
        $cells.Cpe | Should -Be 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*'
    }
    It 'shows nothing after the CPE of the packager changed' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Cpe = 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*' }) -Cpe 'cpe:2.3:a:v:new:*:*:*:*:*:*:*:*' -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be ''
        $cells.Checked | Should -Be ''
        $cells.Cpe | Should -Be 'cpe:2.3:a:v:new:*:*:*:*:*:*:*:*'
    }
    It 'shows nothing for a summary that records no CPE' {
        $noCpe = & $script:NewSummary
        $noCpe.Remove('Cpe')
        $cells = Get-CveGridCells -Summary $noCpe -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be ''
        $cells.Checked | Should -Be ''
    }
    It 'shows the placeholder for a row with no CPE header' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary) -Cpe '' -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be 'CPE missing'
        foreach ($p in 'Cvss', 'Kev', 'Epss', 'Checked', 'Text', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be '' -Because $p }
    }
    It 'shows nothing without a summary' {
        $cells = Get-CveGridCells -Summary $null -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        foreach ($p in 'Count', 'Cvss', 'Kev', 'Epss', 'Checked', 'Text', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be '' -Because $p }
    }
}

Describe 'A stored summary belongs to one layout' {
    It 'shows nothing for a summary of the former layout' {
        $former = @{ QueriedAt = '2026-11-01T00:00:00Z'; From = '1.0'; To = '2.0'; Cpe = $script:GridCpe; Source = 'NVD'; Count = 3; KevCount = 1; MaxCvss = 7.5; MaxSeverity = 'HIGH'; MaxEpss = $null; Error = $null; FixedInLatestCount = 1; FixedEarlierCount = 2; FixedUnknownCount = 0; NewestPublished = '2026-09-30'; LastNvdUpdate = '2026-10-08' }
        foreach ($cells in @(
                (Get-CveGridCells -Summary $former -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0')
                (Get-CveGridCells -Summary $former -Cpe $script:GridCpe)
                (Get-CveGridCells -Summary ([pscustomobject]$former) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0')
            )) {
            foreach ($p in 'Count', 'Cvss', 'Kev', 'Epss', 'Checked', 'Text', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be '' -Because $p }
        }
    }
    It 'shows nothing for the layout number <Schema>' -ForEach @(
        @{ Schema = 1 }
        @{ Schema = 3 }
        @{ Schema = '' }
        @{ Schema = $null }
    ) {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Schema = $Schema }) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be ''
        $cells.Checked | Should -Be ''
    }
    It 'reads the layout number as a number or as text' {
        (Get-CveGridCells -Summary (& $script:NewSummary @{ Schema = '2' }) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be '3'
        (Get-CveGridCells -Summary (& $script:NewSummary @{ Schema = [long]2 }) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be '3'
    }
}

Describe 'A stored summary belongs to one version pair' {
    It 'shows the summary when the row holds the versions <Current> and <Latest>' -ForEach @(
        @{ Current = '1.0'; Latest = '2.0' }
        @{ Current = ' 1.0 '; Latest = '2.0+build5' }
        @{ Current = '1.0 (x64)'; Latest = '2.0 beta' }
    ) {
        (Get-CveGridCells -Summary (& $script:NewSummary) -Cpe $script:GridCpe -CurrentVersion $Current -LatestVersion $Latest).Count | Should -Be '3'
    }
    It 'shows nothing when the row holds the versions <Current> and <Latest>' -ForEach @(
        @{ Current = '1.0'; Latest = '3.0' }
        @{ Current = '1.5'; Latest = '2.0' }
        @{ Current = ''; Latest = '2.0' }
        @{ Current = '1.0'; Latest = '' }
        @{ Current = ''; Latest = '' }
        @{ Current = 'build'; Latest = 'nightly' }
    ) {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary) -Cpe $script:GridCpe -CurrentVersion $Current -LatestVersion $Latest
        $cells.Count | Should -Be ''
        $cells.Checked | Should -Be ''
    }
    It 'does not compare the versions when the caller gives none' {
        (Get-CveGridCells -Summary (& $script:NewSummary) -Cpe $script:GridCpe).Count | Should -Be '3'
    }
    It 'shows the CVEs of the latest version when ConfigMgr holds none' {
        $latestOnly = & $script:NewSummary @{ From = ''; To = '2.0'; Mode = 'Latest'; Count = 1; FixedCount = 0; OpenCount = 1 }
        $cells = Get-CveGridCells -Summary $latestOnly -Cpe $script:GridCpe -CurrentVersion '' -LatestVersion '2.0'
        $cells.Count | Should -Be '1'
        $cells.Fixed | Should -Be ''
        $cells.Text | Should -Be '1 CVE'
        # A ConfigMgr version text with no number counts as no ConfigMgr version.
        (Get-CveGridCells -Summary $latestOnly -Cpe $script:GridCpe -CurrentVersion 'build' -LatestVersion '2.0').Count | Should -Be '1'
    }
    It 'does not show a latest-only summary on a row with a ConfigMgr version' {
        $latestOnly = & $script:NewSummary @{ From = ''; To = '2.0'; Mode = 'Latest'; Count = 1; FixedCount = 0; OpenCount = 1 }
        (Get-CveGridCells -Summary $latestOnly -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be ''
    }
    It 'shows the summary of a ConfigMgr version with no latest version' {
        $deployedOnly = & $script:NewSummary @{ From = '1.0'; To = ''; Mode = 'Deployed'; Count = 2; FixedCount = 0; OpenCount = 2 }
        $cells = Get-CveGridCells -Summary $deployedOnly -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion ''
        $cells.Count | Should -Be '2'
        $cells.Fixed | Should -Be ''
    }
}

Describe 'The CVE cells say what the update fixes' {
    It 'shows the count, the fixed count and the text of an update' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be '3'
        $cells.Fixed | Should -Be '1 of 3'
        $cells.Text | Should -Be 'fixes 1 of 3 CVEs'
        $cells.Cvss | Should -Be '7.5 HIGH'
        $cells.Kev | Should -Be ''
        $cells.Epss | Should -Match '^25'
        $cells.Checked | Should -Be '2026-11-01T00:00:00Z'
        $cells.Newest | Should -Be '2026-09-30'
        $cells.Updated | Should -Be '2026-10-08'
    }
    It 'adds the KEV count to the text and marks the KEV cell' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ KevCount = 2 }) -Cpe $script:GridCpe
        $cells.Text | Should -Be 'fixes 1 of 3 CVEs, 2 in KEV'
        $cells.Kev | Should -Be 'Yes'
    }
    It 'writes one CVE in the singular' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Count = 1; FixedCount = 1; OpenCount = 0 }) -Cpe $script:GridCpe
        $cells.Fixed | Should -Be '1 of 1'
        $cells.Text | Should -Be 'fixes 1 of 1 CVE'
    }
    It 'says when the update fixes none of the CVEs' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ FixedCount = 0; OpenCount = 3; KevCount = 1 }) -Cpe $script:GridCpe
        $cells.Fixed | Should -Be '0 of 3'
        $cells.Text | Should -Be 'fixes 0 of 3 CVEs, 1 in KEV'
    }
    It 'says no CVE and leaves the fixed cell empty for a count of zero' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Count = 0; FixedCount = 0; OpenCount = 0; MaxCvss = $null; MaxSeverity = ''; MaxEpss = $null; NewestPublished = ''; LastNvdUpdate = '' }) -Cpe $script:GridCpe
        $cells.Count | Should -Be '0'
        $cells.Text | Should -Be 'no CVE'
        $cells.Fixed | Should -Be ''
        $cells.Cvss | Should -Be ''
        $cells.Epss | Should -Be ''
        $cells.Checked | Should -Be '2026-11-01T00:00:00Z'
    }
    It 'says Not in NVD, and no count, when NVD holds no record of the product' {
        $unknown = & $script:NewSummary @{ Count = 0; FixedCount = 0; OpenCount = 0; MaxCvss = $null; MaxSeverity = ''; MaxEpss = $null; NewestPublished = ''; LastNvdUpdate = ''; ProductRecords = 0 }
        $fromJson = ($unknown | ConvertTo-Json -Depth 4) | ConvertFrom-Json
        foreach ($source in @($unknown, $fromJson)) {
            $cells = Get-CveGridCells -Summary $source -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
            $cells.Count | Should -Be 'Not in NVD'
            $cells.Text | Should -Be 'product not in NVD'
            $cells.Checked | Should -Be '2026-11-01T00:00:00Z'
            foreach ($p in 'Cvss', 'Kev', 'Epss', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be '' -Because $p }
        }
    }
    It 'keeps the count of zero when NVD holds records of the product or the summary does not say' {
        $zero = @{ Count = 0; FixedCount = 0; OpenCount = 0; MaxCvss = $null; MaxSeverity = ''; MaxEpss = $null; NewestPublished = ''; LastNvdUpdate = '' }
        foreach ($records in @(12, 1, $null, '')) {
            $summary = & $script:NewSummary $zero
            $summary['ProductRecords'] = $records
            $cells = Get-CveGridCells -Summary $summary -Cpe $script:GridCpe
            $cells.Count | Should -Be '0' -Because "ProductRecords '$records'"
            $cells.Text | Should -Be 'no CVE' -Because "ProductRecords '$records'"
        }
        $without = & $script:NewSummary $zero
        $without.Remove('ProductRecords')
        (Get-CveGridCells -Summary $without -Cpe $script:GridCpe).Count | Should -Be '0'
    }
    It 'decides Not in NVD as the module does' {
        foreach ($set in @(
                @{ Count = 0; ProductRecords = 0 }
                @{ Count = 0; ProductRecords = 7 }
                @{ Count = 0; ProductRecords = $null }
                @{ Count = 2; ProductRecords = 9 }
            )) {
            $summary = & $script:NewSummary $set
            $cells = Get-CveGridCells -Summary $summary -Cpe $script:GridCpe
            ($cells.Count -eq 'Not in NVD') | Should -Be (Test-VulnerabilityProductUnknown -Result $summary) -Because "ProductRecords '$($set.ProductRecords)'"
        }
    }
    It 'leaves the fixed cell empty for a ConfigMgr version with no update' {
        $deployed = & $script:NewSummary @{ From = '2.0'; To = '2.0'; Mode = 'Deployed'; Count = 3; FixedCount = 0; OpenCount = 3; KevCount = 1 }
        $cells = Get-CveGridCells -Summary $deployed -Cpe $script:GridCpe -CurrentVersion '2.0' -LatestVersion '2.0'
        $cells.Count | Should -Be '3'
        $cells.Fixed | Should -Be ''
        $cells.Text | Should -Be '3 CVEs, 1 in KEV'
        $cells.Kev | Should -Be 'Yes'
    }
    It 'says no CVE for a ConfigMgr version with no update and no CVE' {
        $deployed = & $script:NewSummary @{ From = '2.0'; To = '2.0'; Mode = 'Deployed'; Count = 0; FixedCount = 0; OpenCount = 0 }
        $cells = Get-CveGridCells -Summary $deployed -Cpe $script:GridCpe
        $cells.Count | Should -Be '0'
        $cells.Text | Should -Be 'no CVE'
        $cells.Fixed | Should -Be ''
    }
    It 'shows Error and the time of a failed lookup' {
        $cells = Get-CveGridCells -Summary (& $script:NewSummary @{ Error = 'NVD and Shodan CVEDB both failed' }) -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be 'Error'
        $cells.Checked | Should -Be '2026-11-01T00:00:00Z'
        foreach ($p in 'Cvss', 'Kev', 'Epss', 'Text', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be '' -Because $p }
    }
    It 'gives the same cells for a hashtable, for the summary of the module and for a summary read back from JSON' {
        $result = & $script:NewResult -Cpe $script:GridCpe -From '1.0' -To '2.0' -QueriedAt '2026-11-01T00:00:00Z' -Cves @(
            (& $script:NewRecord -Id 'CVE-2025-0001' -Score 7.5 -Severity 'HIGH' -Status 'Fixed' -Kev)
            (& $script:NewRecord -Id 'CVE-2025-0003' -Score 9.1 -Severity 'CRITICAL' -Status 'Open')
        )
        $result.MaxEpss = 0.25
        $summary = Get-VulnerabilitySummary -Result $result
        ($summary -is [System.Collections.IDictionary]) | Should -BeTrue
        $table = @{}
        foreach ($k in $summary.Keys) { $table[$k] = $summary[$k] }
        $fromJson = ($summary | ConvertTo-Json -Depth 4) | ConvertFrom-Json
        ($fromJson -is [System.Collections.IDictionary]) | Should -BeFalse
        $expected = Get-CveGridCells -Summary $table -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $expected.Count | Should -Be '2'
        $expected.Fixed | Should -Be '1 of 2'
        $expected.Text | Should -Be 'fixes 1 of 2 CVEs, 1 in KEV'
        $expected.Cvss | Should -Be '9.1 CRITICAL'
        $expected.Checked | Should -Be '2026-11-01T00:00:00Z'
        foreach ($source in @($summary, $fromJson)) {
            $cells = Get-CveGridCells -Summary $source -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
            foreach ($p in 'Cpe', 'Count', 'Cvss', 'Kev', 'Epss', 'Checked', 'Text', 'Fixed', 'Newest', 'Updated') { $cells.$p | Should -Be $expected.$p -Because $p }
        }
    }

    It 'shows the lookup time as the text of the lookup when the summary holds a <Kind> DateTime' -ForEach @(
        @{ Kind = 'Utc' }
        @{ Kind = 'Unspecified' }
    ) {
        # PowerShell 7 reads the time text of a JSON file back as a DateTime.
        $time = [datetime]::SpecifyKind([datetime]::new(2026, 11, 1, 13, 5, 9), [System.DateTimeKind]$Kind)
        $summary = & $script:NewSummary @{ QueriedAt = $time }
        $cells = Get-CveGridCells -Summary $summary -Cpe $script:GridCpe -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Checked | Should -BeExactly '2026-11-01T13:05:09Z'
    }
}

Describe 'One version key for the grid and the module' {
    It 'reduces the version text "<Text>" to "<Key>", as the module does' -ForEach @(
        @{ Text = '1.0'; Key = '1.0' }
        @{ Text = ' 1.0 '; Key = '1.0' }
        @{ Text = '11.0.30+7'; Key = '11.0.30' }
        @{ Text = '140.17.0esr'; Key = '140.17.0esr' }
        @{ Text = '2.0.0-rc1'; Key = '2.0.0-rc1' }
        @{ Text = 'v3.2'; Key = 'v3.2' }
        @{ Text = '1.2.3.4.5'; Key = '1.2.3.4.5' }
        @{ Text = '7.60c'; Key = '7.60c' }
        @{ Text = '1.0 (x64)'; Key = '1.0' }
        @{ Text = "3.1`tbeta"; Key = '3.1' }
        @{ Text = '20250101000000'; Key = '20250101000000' }
        @{ Text = '1.0+'; Key = '1.0' }
        @{ Text = '+7'; Key = '' }
        @{ Text = 'build 0143'; Key = '' }
        @{ Text = 'release-5'; Key = 'release-5' }
        @{ Text = 'nightly'; Key = '' }
        @{ Text = '   '; Key = '' }
        @{ Text = ''; Key = '' }
        @{ Text = $null; Key = '' }
    ) {
        ConvertTo-CveVersionKey -Version $Text | Should -BeExactly $Key
        ConvertTo-CveVersionKey -Version $Text | Should -BeExactly (ConvertTo-VulnerabilityVersion -Version $Text)
    }
    It 'joins the keys of a row as <Key> for the versions "<Current>" and "<Latest>"' -ForEach @(
        @{ Current = '1.0'; Latest = '2.0'; Key = '1.0|2.0' }
        @{ Current = ''; Latest = '2.0'; Key = '|2.0' }
        @{ Current = '1.0'; Latest = ''; Key = '1.0|' }
        @{ Current = ''; Latest = ''; Key = '' }
        @{ Current = $null; Latest = $null; Key = '' }
        @{ Current = 'build'; Latest = 'nightly'; Key = '' }
        @{ Current = 'build'; Latest = '2.0'; Key = '|2.0' }
        @{ Current = '1.0 (x64)'; Latest = '2.0+7'; Key = '1.0|2.0' }
    ) {
        ConvertTo-CvePairKey -Current $Current -Latest $Latest | Should -BeExactly $Key
    }
}

Describe 'One comparer decides whether an update exists' {
    It 'decides <Current> to <Latest> as update: <Update>, in the grid and in the lookup of the module' -ForEach @(
        @{ Current = '1.0'; Latest = '2.0'; Update = $true }
        @{ Current = '2.0'; Latest = '1.0'; Update = $false }
        @{ Current = '2.0'; Latest = '2.0'; Update = $false }
        @{ Current = '8.9.8'; Latest = '8.9.8.1'; Update = $true }
        @{ Current = '26.2.2.2'; Latest = '26.2.2'; Update = $false }
        @{ Current = '31'; Latest = '31.0.1'; Update = $true }
        @{ Current = '140.17.0esr'; Latest = '140.17.1'; Update = $true }
        @{ Current = '11.0.30+7'; Latest = '11.0.31'; Update = $true }
        @{ Current = '20240101000000'; Latest = '20250101000000'; Update = $true }
        @{ Current = '20250101000000'; Latest = '20240101000000'; Update = $false }
        @{ Current = 'build 0143'; Latest = 'build 0144'; Update = $false }
    ) {
        ((Compare-SemVer -A $Current -B $Latest) -lt 0) | Should -Be $Update
        # The grid writes "fixes F of N" for a row with an update; the module counts the fixed CVEs in mode Update only.
        $mode = Get-VulnerabilityLookupMode -From (ConvertTo-CveVersionKey -Version $Current) -To (ConvertTo-CveVersionKey -Version $Latest)
        ($mode -eq 'Update') | Should -Be $Update
    }
    It 'compares the strict form over all four parts' {
        Compare-SemVer -A '1.2.3.4' -B '1.2.3.5' -Strict | Should -Be (-1)
        Compare-SemVer -A '1.2.3.5' -B '1.2.3' -Strict | Should -Be 1
        Compare-SemVer -A '20250101000000.1' -B '20250101000000.1' -Strict | Should -Be 0
    }
    It 'leaves no use of the vulnerability module comparer in the application script' {
        $script:Ast.Extent.Text | Should -Not -Match 'Compare-VulnerabilityVersionValue'
    }
}

Describe 'Check CVEs skips a row only for a missing CPE or a missing version' {
    It 'does not skip the versions "<Current>" and "<Latest>"' -ForEach @(
        @{ Current = '1.0'; Latest = '2.0' }
        @{ Current = '2.0'; Latest = '2.0' }
        @{ Current = '2.0'; Latest = '1.0' }
        @{ Current = '26.2.2.2'; Latest = '26.2.2' }
        @{ Current = '1.0'; Latest = '' }
        @{ Current = ''; Latest = '2.0' }
        @{ Current = '1.0'; Latest = $null }
        @{ Current = 'build 0143'; Latest = '2.0' }
    ) {
        Get-CveSkipReason -Cpe $script:GridCpe -Current $Current -Latest $Latest | Should -BeExactly ''
    }
    It 'names the missing version for the versions "<Current>" and "<Latest>"' -ForEach @(
        @{ Current = ''; Latest = '' }
        @{ Current = $null; Latest = $null }
        @{ Current = '  '; Latest = '' }
        @{ Current = 'build 0143'; Latest = 'build 0144' }
        @{ Current = 'nightly'; Latest = '' }
    ) {
        Get-CveSkipReason -Cpe $script:GridCpe -Current $Current -Latest $Latest | Should -BeExactly 'No version'
    }
    It 'names the missing CPE first' {
        Get-CveSkipReason -Cpe '' -Current '1' -Latest '2' | Should -BeExactly 'CPE missing'
        Get-CveSkipReason -Cpe '   ' -Current '1' -Latest '2' | Should -BeExactly 'CPE missing'
        Get-CveSkipReason -Cpe $null -Current '1' -Latest '' | Should -BeExactly 'CPE missing'
        Get-CveSkipReason -Cpe '' -Current '' -Latest '' | Should -BeExactly 'CPE missing'
    }
}

Describe 'CVE cells follow the version pair they describe' {
    BeforeAll {
        $script:NewRow = {
            param($cur, $lat, $status = 'Update available (fixes 1 of 3 CVEs, 1 in KEV)')
            $r = [pscustomobject]@{ Cpe = $script:GridCpe; CurrentVersion = $cur; LatestVersion = $lat; Status = $status
                CveCount = '3'; CveCvss = '7.5 HIGH'; CveKev = 'Yes'; CveEpss = '1.0%'; CveChecked = '2026-11-01T00:00:00Z'; CveFixed = '1 of 3'; CveNewest = '2026-01-01'; CveUpdated = '2026-01-02' }
            Set-CveCellsPair -Row $r
            $r
        }
        $script:CellNames = @('CveCount', 'CveCvss', 'CveKev', 'CveEpss', 'CveChecked', 'CveFixed', 'CveNewest', 'CveUpdated')
    }
    It 'records the pair that the row holds' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.CvePair | Should -BeExactly '1.0|2.0'
        (& $script:NewRow '' '2.0').CvePair | Should -BeExactly '|2.0'
        (& $script:NewRow '1.0' '').CvePair | Should -BeExactly '1.0|'
        $r.CurrentVersion = '1.5'
        Set-CveCellsPair -Row $r
        $r.CvePair | Should -BeExactly '1.5|2.0'
        Set-CveCellsPair -Row $r -Pair ''
        $r.CvePair | Should -BeExactly ''
    }
    It 'keeps the cells while the pair is the same, also when the text of a version differs' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.LatestVersion = '2.0+build5'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be '3'
        $r.CveFixed | Should -Be '1 of 3'
        $r.Status | Should -Be 'Update available (fixes 1 of 3 CVEs, 1 in KEV)'
        $r.CvePair | Should -BeExactly '1.0|2.0'
    }
    It 'clears the cells and the status suffix when the latest version changes' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.LatestVersion = '3.0'
        Clear-StaleCveCells -Row $r
        foreach ($p in $script:CellNames) { $r.$p | Should -Be '' -Because $p }
        $r.Status | Should -Be 'Update available'
        $r.CvePair | Should -BeExactly ''
    }
    It 'clears the cells when the ConfigMgr version changes' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.CurrentVersion = '1.5'
        Clear-StaleCveCells -Row $r
        foreach ($p in $script:CellNames) { $r.$p | Should -Be '' -Because $p }
        $r.Status | Should -Be 'Update available'
    }
    It 'clears the cells of the latest version when ConfigMgr reports a version' {
        $r = & $script:NewRow '' '2.0' 'Latest retrieved'
        $r.CurrentVersion = '1.0'
        Clear-StaleCveCells -Row $r
        foreach ($p in $script:CellNames) { $r.$p | Should -Be '' -Because $p }
        $r.Status | Should -Be 'Latest retrieved'
    }
    It 'clears the cells when ConfigMgr holds the application no more' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.CurrentVersion = ''
        Clear-StaleCveCells -Row $r
        foreach ($p in $script:CellNames) { $r.$p | Should -Be '' -Because $p }
        $r.Status | Should -Be 'Update available'
    }
    It 'keeps the cells of the latest version while ConfigMgr holds no version' {
        $r = & $script:NewRow '' '2.0' 'Latest retrieved'
        $r.LatestVersion = '2.0 '
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be '3'
        $r.CvePair | Should -BeExactly '|2.0'
    }
    It 'clears the cells of a ConfigMgr version when the first latest version arrives' {
        $r = & $script:NewRow '1.0' '' 'ConfigMgr version retrieved'
        $r.LatestVersion = '2.0'
        Clear-StaleCveCells -Row $r
        foreach ($p in $script:CellNames) { $r.$p | Should -Be '' -Because $p }
    }
    It 'leaves a status that carries no CVE text as it is' {
        $r = & $script:NewRow '2.0' '2.0' 'Up to date'
        $r.LatestVersion = '2.1'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be ''
        $r.Status | Should -Be 'Up to date'
    }
    It 'clears filled cells of a row that recorded no pair' {
        $r = [pscustomobject]@{ Cpe = $script:GridCpe; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available (3 CVEs)'; CveCount = '3'; CveCvss = '7.5 HIGH' }
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be ''
        $r.CveCvss | Should -Be ''
        $r.Status | Should -Be 'Update available'
        $r.CvePair | Should -BeExactly ''
    }
    It 'does nothing for a row with no filled cell' {
        $r = [pscustomobject]@{ Cpe = $script:GridCpe; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available'; CveCount = ''; CveCvss = '' }
        Clear-StaleCveCells -Row $r
        $r.Status | Should -Be 'Update available'
        $r.PSObject.Properties['CvePair'] | Should -BeNullOrEmpty
    }
    It 'keeps the missing-CPE text of a row with no CPE' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.Cpe = ''
        $r.LatestVersion = '3.0'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be 'CPE missing'
        $r.CveCvss | Should -Be ''
    }
    It 'runs after every place that changes the version of a row' {
        $assignments = $script:Ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and
                $n.Left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Left.Expression.VariablePath.UserPath -eq 'row' -and
                "$($n.Left.Member)" -in 'CurrentVersion', 'LatestVersion'
            }, $true)
        $assignments.Count | Should -BeGreaterThan 5
        $missing = @()
        foreach ($a in $assignments) {
            $block = $a.Parent
            if ($block -isnot [System.Management.Automation.Language.StatementBlockAst] -and $block -isnot [System.Management.Automation.Language.NamedBlockAst]) { $missing += $a.Extent.StartLineNumber; continue }
            $statements = @($block.Statements)
            $at = [array]::IndexOf($statements, $a)
            $next = $(if ($at -ge 0 -and $at + 1 -lt $statements.Count) { $statements[$at + 1] } else { $null })
            if (-not $next -or $next.Extent.Text -notmatch '^Clear-StaleCveCells\b') { $missing += $a.Extent.StartLineNumber }
        }
        $missing | Should -BeNullOrEmpty -Because ('lines without a Clear-StaleCveCells call right after the assignment: ' + ($missing -join ', '))
    }
}

Describe 'Detail and report read the result of this session first' {
    BeforeAll {
        $script:SCpe = 'cpe:2.3:a:vendor:sessionprod:*:*:*:*:*:*:*:*'
        $script:SRow = { param($count) [pscustomobject]@{ Application = 'S'; Vendor = 'V'; Script = 'package-s.ps1'; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available'; Cpe = $script:SCpe; CveCount = $count; CveChecked = '' } }
        $script:SResult = & $script:NewResult -Cpe $script:SCpe -From '1.0' -To '2.0' -Cves @(1..4 | ForEach-Object { & $script:NewRecord -Id ('CVE-2026-100' + $_) -Score 5.0 -Severity 'MEDIUM' -Status 'Fixed' })
        $script:SCached = { param($from, $to, $cpe = $script:SCpe) & $script:NewResult -Cpe $cpe -From $from -To $to -Cves @(& $script:NewRecord -Id 'CVE-2026-2001' -Score 6.5 -Severity 'MEDIUM' -Status 'Open') }
    }
    BeforeEach { & $script:ResetState }
    AfterEach { $script:CveSessionResults = $null }

    It 'returns the session result when the cache holds none' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $found = Get-CveCachedResult -Row (& $script:SRow '4')
        $found.Count | Should -Be 4
        $found.Schema | Should -Be 2
        $found.Mode | Should -Be 'Update'
    }
    It 'returns the session result before the cached result' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        (Get-CveCachedResult -Row (& $script:SRow '4')).Count | Should -Be 4
    }
    It 'returns the cached result when the session holds none' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        $found = Get-CveCachedResult -Row (& $script:SRow '1')
        $found.Count | Should -Be 1
        $found.Cves[0].Id | Should -Be 'CVE-2026-2001'
        $found.Cves[0].Status | Should -Be 'Open'
        $script:CveSessionResults = @{ 'package-other.ps1' = $script:SResult }
        (Get-CveCachedResult -Row (& $script:SRow '1')).Count | Should -Be 1
    }
    It 'ignores a session result for another CPE or version pair' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $other = & $script:SRow '4'
        $other.LatestVersion = '3.0'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
        $other = & $script:SRow '4'
        $other.CurrentVersion = ''
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
        $other = & $script:SRow '4'
        $other.Cpe = 'cpe:2.3:a:vendor:otherprod:*:*:*:*:*:*:*:*'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
    }
    It 'ignores a cached result for another CPE or version pair' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        $other = & $script:SRow '1'
        $other.LatestVersion = '3.0'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
        $other = & $script:SRow '1'
        $other.CurrentVersion = '1.5'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
        $other = & $script:SRow '1'
        $other.CurrentVersion = ''
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
        $other = & $script:SRow '1'
        $other.Cpe = 'cpe:2.3:a:vendor:otherprod:*:*:*:*:*:*:*:*'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
    }
    It 'reads the cached result by the version keys of the row' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        $row = & $script:SRow '1'
        $row.CurrentVersion = '1.0 (x64)'
        $row.LatestVersion = '2.0+build5'
        (Get-CveCachedResult -Row $row).Count | Should -Be 1
    }
    It 'reads the result of the latest version for a row with no ConfigMgr version' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '' '2.0')
        $row = & $script:SRow '1'
        $row.CurrentVersion = ''
        $found = Get-CveCachedResult -Row $row
        $found.Count | Should -Be 1
        $found.Mode | Should -Be 'Latest'
        $found.Examined | Should -Be '2.0'
        Get-CveCachedResult -Row (& $script:SRow '1') | Should -BeNullOrEmpty
    }
    It 'ignores a session result that holds an error' {
        $failed = & $script:NewResult -Cpe $script:SCpe -From '1.0' -To '2.0'
        $failed.Error = 'NVD and Shodan CVEDB both failed'
        $script:CveSessionResults = @{ 'package-s.ps1' = $failed }
        Find-CveSessionResult -Row (& $script:SRow 'Error') -Results $script:CveSessionResults | Should -BeNullOrEmpty
        Get-CveCachedResult -Row (& $script:SRow 'Error') | Should -BeNullOrEmpty
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        (Get-CveCachedResult -Row (& $script:SRow 'Error')).Count | Should -Be 1
    }
    It 'returns nothing for a row with no CPE header or no version' {
        Save-VulnerabilityCachedResult -Result (& $script:SCached '1.0' '2.0')
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $noCpe = & $script:SRow 'CPE missing'
        $noCpe.Cpe = ''
        Get-CveCachedResult -Row $noCpe | Should -BeNullOrEmpty
        $noVersion = & $script:SRow 'No version'
        $noVersion.CurrentVersion = ''; $noVersion.LatestVersion = ''
        Get-CveCachedResult -Row $noVersion | Should -BeNullOrEmpty
    }
    It 'finds a session result by the script name of the row' {
        $results = @{ 'package-s.ps1' = $script:SResult }
        (Find-CveSessionResult -Row (& $script:SRow '4') -Results $results).Count | Should -Be 4
        Find-CveSessionResult -Row (& $script:SRow '4') -Results $null | Should -BeNullOrEmpty
        Find-CveSessionResult -Row (& $script:SRow '4') -Results @{ 'package-other.ps1' = $script:SResult } | Should -BeNullOrEmpty
    }
    It 'keeps the failure note for a row whose latest lookup failed' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $rows = Get-CveReportRows -Rows @((& $script:SRow 'Error')) -Results $script:CveSessionResults
        $rows[0].Result.Count | Should -Be 4
        $rows[0].Note | Should -Match '^Last lookup failed'
    }
    It 'reports the session result when the caller passes no results' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $rows = Get-CveReportRows -Rows @((& $script:SRow '4'))
        $rows[0].Result.Count | Should -Be 4
        $rows[0].Note | Should -Be ''
    }
}

Describe 'The monitor counts cached results apart and logs the ConfigMgr query' {
    BeforeEach { & $script:ResetState }

    It 'reports the results taken from the cache when the CVE lookup is skipped' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $script:Log = @()
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipCve
        $s.LookedUp | Should -Be 0
        $s.Cached | Should -Be 1
        # The totals count the lookups of the run only.
        $s.Cves | Should -Be 0
        ($script:Log | Where-Object { $_ -like '*Monitor finished*' }) | Should -Match '1 from the cache'
    }

    It 'counts an older cached result that stands in for a failed lookup' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $script:CveFail = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.LookedUp | Should -Be 0
        $s.Cached | Should -Be 1
        $s.Failed | Should -Be 1
    }

    It 'reports no cached result when the run looked every application up itself' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-beta'
        $s.LookedUp | Should -Be 2
        $s.Cached | Should -Be 0
    }

    It 'logs one line for each lookup with the result text and the source' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-beta'
        $alpha = @($script:Log | Where-Object { $_ -like '`[cve`] Alpha: *' })
        $alpha.Count | Should -Be 1
        $alpha[0] | Should -Be ('[cve] Alpha: {0} (NVD)' -f (ConvertTo-VulnerabilityResultText -Result (Read-VulnerabilityCachedResult -Cpe $script:AlphaCpe -CurrentVersion '1.0' -LatestVersion '2.0')))
        @($script:Log | Where-Object { $_ -like '`[cve`] Beta: *(NVD)' }).Count | Should -Be 1
    }

    It 'logs one line for the ConfigMgr query with the site, the provider and the counts' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ProviderMachineName 'cm01.contoso.test' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-delta' -SkipLatest
        $line = @($script:Log | Where-Object { $_ -like '*ConfigMgr query*' })
        $line.Count | Should -Be 1
        $line[0] | Should -Match 'MCM'
        $line[0] | Should -Match 'cm01\.contoso\.test'
        $line[0] | Should -Match '1 found'
        $line[0] | Should -Match '1 not found'
    }

    It 'logs no ConfigMgr query line when the query is skipped' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipMecm
        @($script:Log | Where-Object { $_ -like '*ConfigMgr query*' }).Count | Should -Be 0
    }
}

Describe 'The Check CVEs step of the background run' {
    BeforeAll {
        $script:NewStepRow = {
            param([string]$Name, [string]$Cpe, [string]$Current, [string]$Latest, [string]$Status)
            [pscustomobject]@{ Application = $Name; Script = ('package-' + $Name.ToLowerInvariant() + '.ps1'); Cpe = $Cpe; CurrentVersion = $Current; LatestVersion = $Latest; Status = $Status; AdvisoryUrl = ''
                CveCount = ''; CveCvss = ''; CveKev = ''; CveEpss = ''; CveChecked = ''; CveFixed = ''; CveNewest = ''; CveUpdated = '' }
        }
        $script:InvokeStep = {
            param([object[]]$Rows, [hashtable]$Ctx = $null)
            if (-not $Ctx) { $Ctx = @{ CveSession = [hashtable]::Synchronized(@{}); NvdApiKey = 'k'; CacheTtlMinutes = 360; KevCatalog = $null; ForceRefresh = $false } }
            $state = [hashtable]::Synchronized(@{ Step = ''; LogQueue = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'); CveResults = [hashtable]::Synchronized(@{}) })
            $counts = [ordered]@{ Checked = 0; Skipped = 0; CheckFailed = 0 }
            $null = & $script:CheckCveStep -Rows $Rows -Ctx $Ctx -State $state -counts $counts
            [pscustomobject]@{ Ctx = $Ctx; State = $state; Counts = $counts; Log = @($state.LogQueue.ToArray()) }
        }
        $script:StepCells = @('CveCvss', 'CveKev', 'CveEpss', 'CveChecked', 'CveFixed', 'CveNewest', 'CveUpdated')
    }
    BeforeEach { & $script:ResetState }

    It 'decides a skip with Get-CveSkipReason alone' {
        $assignments = @($script:CheckCveStepAst.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$reason'
                }, $true))
        $assignments.Count | Should -Be 1
        $assignments[0].Right.Extent.Text | Should -Be 'Get-CveSkipReason -Cpe $cpe -Current $current -Latest $latest'
        $script:CheckCveStepAst.Extent.Text | Should -Not -Match 'Compare-SemVer'
    }

    It 'skips a row with no CPE header and clears its other cells' {
        $row = & $script:NewStepRow 'Gamma' '' '1.0' '2.0' 'Update available'
        foreach ($p in $script:StepCells) { $row.$p = 'old' }
        $run = & $script:InvokeStep @($row)
        $script:CveCalls.Count | Should -Be 0
        $row.CveCount | Should -Be 'CPE missing'
        foreach ($p in $script:StepCells) { $row.$p | Should -Be '' -Because $p }
        $row.Status | Should -Be 'Update available'
        $run.Counts['Skipped'] | Should -Be 1
        $run.Counts['Checked'] | Should -Be 0
        $run.Log | Should -Contain 'Skipped: CPE missing.'
    }

    It 'skips a row with no version' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '' '' 'Ready'
        $textOnly = & $script:NewStepRow 'Beta' $script:BetaCpe 'build' 'nightly' 'Ready'
        $run = & $script:InvokeStep @($row, $textOnly)
        $script:CveCalls.Count | Should -Be 0
        $script:KevCalls | Should -Be 0
        $row.CveCount | Should -Be 'No version'
        $textOnly.CveCount | Should -Be 'No version'
        $run.Counts['Skipped'] | Should -Be 2
        $run.Log | Should -Contain 'Skipped: No version.'
    }

    It 'looks up a row with an update and shows what the update fixes' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'
        $run = & $script:InvokeStep @($row)
        $script:CveCalls.Count | Should -Be 1
        $script:CveCalls[0].From | Should -Be '1.0'
        $script:CveCalls[0].To | Should -Be '2.0'
        $script:CveCalls[0].Key | Should -Be 'k'
        $script:CveCalls[0].Force | Should -BeFalse
        $row.CveCount | Should -Be '3'
        $row.CveFixed | Should -Be '2 of 3'
        $row.CveKev | Should -Be 'Yes'
        $row.CveCvss | Should -Be '9.1 CRITICAL'
        $row.CveChecked | Should -Match $script:StampPattern
        $row.CveNewest | Should -Be '2025-01-01'
        $row.CveUpdated | Should -Be '2025-02-01'
        $row.Status | Should -Be 'Update available (fixes 2 of 3 CVEs, 1 in KEV)'
        $row.CvePair | Should -BeExactly '1.0|2.0'
        $run.Counts['Checked'] | Should -Be 1
        $run.Counts['Skipped'] | Should -Be 0
        $run.State.CveResults['package-alpha.ps1'].Count | Should -Be 3
        $run.Ctx.CveSession['package-alpha.ps1'].Count | Should -Be 3
        $stored = (Read-PackagerHistory)['package-alpha'].VulnSummary
        $stored.Schema | Should -Be 2
        $stored.Mode | Should -Be 'Update'
        $stored.FixedCount | Should -Be 2
        @($run.Log | Where-Object { $_ -like '3 CVEs affect 1.0*Source NVD; max CVSS 9.1 CRITICAL' }).Count | Should -Be 1
    }

    It 'replaces the CVE text of the status on the next lookup' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'
        $null = & $script:InvokeStep @($row)
        $script:CveEmpty = $true
        $null = & $script:InvokeStep @($row)
        $row.Status | Should -Be 'Update available (no CVE)'
        $row.CveCount | Should -Be '0'
        $row.CveFixed | Should -Be ''
    }

    It 'shows Not in NVD for a product NVD holds no record of' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'
        $script:CveUnknown = $true
        $run = & $script:InvokeStep @($row)
        $row.CveCount | Should -Be 'Not in NVD'
        $row.Status | Should -Be 'Update available (product not in NVD)'
        foreach ($p in 'CveFixed', 'CveKev', 'CveCvss', 'CveNewest', 'CveUpdated') { $row.$p | Should -Be '' -Because $p }
        $row.CveChecked | Should -Match $script:StampPattern
        $run.Counts['Checked'] | Should -Be 1
        @($run.Log | Where-Object { $_ -like 'NVD holds no CVE record under the product name of the CPE header.*' }).Count | Should -Be 1
        # The report reads the result of the row and gives no note of a skipped lookup.
        $rows = Get-CveReportRows -Rows @($row) -Results $run.State.CveResults
        $rows[0].Note | Should -Be ''
        $rows[0].Result.ProductRecords | Should -Be 0
    }

    It 'looks up a row with no update and lists its CVEs as open' {
        $row = & $script:NewStepRow 'Beta' $script:BetaCpe '5.0' '5.0' 'Up to date'
        $run = & $script:InvokeStep @($row)
        $script:CveCalls.Count | Should -Be 1
        $row.CveCount | Should -Be '2'
        $row.CveFixed | Should -Be ''
        $row.CveKev | Should -Be 'Yes'
        $row.Status | Should -Be 'Up to date'
        $row.CvePair | Should -BeExactly '5.0|5.0'
        $run.Counts['Checked'] | Should -Be 1
        (Read-PackagerHistory)['package-beta'].VulnSummary.Mode | Should -Be 'Deployed'
    }

    It 'looks up the latest version of a row with no ConfigMgr version' {
        $row = & $script:NewStepRow 'Epsilon' $script:EpsilonCpe '' '3.0' 'Latest retrieved'
        $run = & $script:InvokeStep @($row)
        $script:CveCalls.Count | Should -Be 1
        $script:CveCalls[0].From | Should -Be ''
        $script:CveCalls[0].To | Should -Be '3.0'
        $row.CveCount | Should -Be '1'
        $row.CveFixed | Should -Be ''
        $row.Status | Should -Be 'Latest retrieved'
        $row.CvePair | Should -BeExactly '|3.0'
        $run.Counts['Checked'] | Should -Be 1
        (Read-PackagerHistory)['package-epsilon'].VulnSummary.Mode | Should -Be 'Latest'
    }

    It 'looks up the ConfigMgr version of a row with no latest version' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '' 'ConfigMgr version retrieved'
        $null = & $script:InvokeStep @($row)
        $script:CveCalls[0].From | Should -Be '1.0'
        $script:CveCalls[0].To | Should -Be ''
        $row.CveCount | Should -Be '2'
        $row.CveFixed | Should -Be ''
        $row.CvePair | Should -BeExactly '1.0|'
    }

    It 'reads the KEV catalog one time for the rows of a run' {
        $rows = @((& $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'), (& $script:NewStepRow 'Beta' $script:BetaCpe '5.0' '5.0' 'Up to date'))
        $run = & $script:InvokeStep $rows
        $script:CveCalls.Count | Should -Be 2
        $script:KevCalls | Should -Be 1
        $run.Counts['Checked'] | Should -Be 2
    }

    It 'passes the forced refresh of the CVE window on to the lookup' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'
        $null = & $script:InvokeStep @($row) @{ CveSession = [hashtable]::Synchronized(@{}); NvdApiKey = ''; CacheTtlMinutes = 360; KevCatalog = @{}; ForceRefresh = $true }
        $script:CveCalls[0].Force | Should -BeTrue
        $script:KevCalls | Should -Be 0
    }

    It 'shows a failed lookup and keeps the last good summary and session result' {
        $row = & $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'
        $first = & $script:InvokeStep @($row)
        $script:CveFail = $true
        $run = & $script:InvokeStep @($row) $first.Ctx
        $row.CveCount | Should -Be 'Error'
        $row.CveChecked | Should -Match $script:StampPattern
        foreach ($p in 'CveCvss', 'CveKev', 'CveEpss', 'CveFixed', 'CveNewest', 'CveUpdated') { $row.$p | Should -Be '' -Because $p }
        $run.Counts['CheckFailed'] | Should -Be 1
        $run.Counts['Checked'] | Should -Be 0
        @($run.Log | Where-Object { $_ -like 'Error: NVD and Shodan CVEDB both failed*' }).Count | Should -Be 1
        $run.State.CveResults['package-alpha.ps1'].Error | Should -Not -BeNullOrEmpty
        $run.Ctx.CveSession['package-alpha.ps1'].Count | Should -Be 3
        $stored = (Read-PackagerHistory)['package-alpha'].VulnSummary
        $stored.Count | Should -Be 3
        $stored.Error | Should -BeNullOrEmpty
        # The detail window and the report still find the good result.
        $script:CveSessionResults = $run.Ctx.CveSession
        (Get-CveCachedResult -Row $row).Count | Should -Be 3
    }

    It 'shows Error for a lookup that throws and continues with the next row' {
        $script:CveThrow = '*:alpha:*'
        $rows = @((& $script:NewStepRow 'Alpha' $script:AlphaCpe '1.0' '2.0' 'Update available'), (& $script:NewStepRow 'Beta' $script:BetaCpe '5.0' '5.0' 'Up to date'))
        $run = & $script:InvokeStep $rows
        $rows[0].CveCount | Should -Be 'Error'
        $rows[0].Status | Should -Be 'Update available'
        $rows[1].CveCount | Should -Be '2'
        $run.Counts['CheckFailed'] | Should -Be 1
        $run.Counts['Checked'] | Should -Be 1
        $run.Log | Should -Contain 'Error: lookup stopped'
        (Read-PackagerHistory).ContainsKey('package-alpha') | Should -BeFalse
    }
}
