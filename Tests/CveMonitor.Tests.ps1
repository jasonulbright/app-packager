#Requires -Modules Pester

<#
.SYNOPSIS
    Headless CVE monitor: Invoke-CveMonitor from start-apppackager.ps1.

.DESCRIPTION
    The function is lifted out of the application script by AST and run with
    stubbed vendor, ConfigMgr and CVE lookups. The report writers and the
    history file are the real ones; the history is redirected into the test
    drive.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerVulnerability.psd1') -Force

    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $script:Ast = $ast
    foreach ($name in 'Invoke-CveMonitor', 'Get-CveGridCells', 'Get-CveCachedResult', 'Find-CveSessionResult', 'Get-CveReportRows', 'Compare-SemVer', 'Remove-ExpiredCveReports', 'ConvertTo-CveVersionKey', 'ConvertTo-CvePairKey', 'Clear-StaleCveCells', 'Set-CveCellsPair', 'Get-CveSkipReason') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    $script:OrigLocalAppData = $env:LOCALAPPDATA
    $env:LOCALAPPDATA = Join-Path $TestDrive 'LocalAppData'
    New-Item -ItemType Directory -Path (Join-Path $env:LOCALAPPDATA 'AppPackager') -Force | Out-Null

    function Write-Log { param([string]$Message, [string]$Level = 'INFO') $script:Log += $Message }
    function Get-Packagers {
        param($Root)
        @(
            [pscustomobject]@{ Script = 'package-alpha.ps1'; FullPath = 'C:\x\package-alpha.ps1'; Application = 'Alpha'; Vendor = 'V'; CMName = 'Alpha'; Status = 'Ready'; CPE = 'cpe:2.3:a:v:alpha:*:*:*:*:*:*:*:*'; SecurityAdvisoryUrl = 'https://v.example.test/security' }
            [pscustomobject]@{ Script = 'package-beta.ps1'; FullPath = 'C:\x\package-beta.ps1'; Application = 'Beta'; Vendor = 'V'; CMName = 'Beta'; Status = 'Ready'; CPE = 'cpe:2.3:a:v:beta:*:*:*:*:*:*:*:*'; SecurityAdvisoryUrl = '' }
            [pscustomobject]@{ Script = 'package-gamma.ps1'; FullPath = 'C:\x\package-gamma.ps1'; Application = 'Gamma'; Vendor = 'V'; CMName = 'Gamma'; Status = 'Ready'; CPE = ''; SecurityAdvisoryUrl = '' }
            [pscustomobject]@{ Script = 'package-delta.ps1'; FullPath = 'C:\x\package-delta.ps1'; Application = 'Delta'; Vendor = 'V'; CMName = 'Delta'; Status = 'Ready'; CPE = 'cpe:2.3:a:v:delta:*:*:*:*:*:*:*:*'; SecurityAdvisoryUrl = '' }
        )
    }
    function Invoke-PackagerGetLatestVersion {
        param($PackagerPath, $SiteCode, $FileServerPath, $DownloadRoot, $M365Channel, $M365DeployMode)
        $script:LatestCalls += $PackagerPath
        switch -Wildcard ($PackagerPath) {
            '*alpha*' { '2.0' }
            '*beta*'  { '5.0' }
            '*gamma*' { '1.1' }
            '*delta*' { throw 'vendor page unreachable' }
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
    function Get-KevCatalog { @{} }
    function Get-ProductVulnerabilities {
        param($Cpe, $CurrentVersion, $LatestVersion, $ApiKey, $CachePath, $CacheTtlMinutes, [switch]$Force, $VendorAdvisoryUrl, $KevCatalog, [switch]$SkipEnrichment)
        $script:CveCalls += [pscustomobject]@{ Cpe = $Cpe; From = $CurrentVersion; To = $LatestVersion; Advisory = $VendorAdvisoryUrl; Key = $ApiKey }
        if ($script:CveFail) {
            return [pscustomobject]@{ Cpe = $Cpe; From = $CurrentVersion; To = $LatestVersion; Source = ''; QueriedAt = 'q'; Count = 0; MaxCvss = $null; MaxSeverity = ''; KevCount = 0; MaxEpss = $null; UnfixedCount = 0; Error = 'NVD and Shodan CVEDB both failed'; Links = [ordered]@{}; Cves = @(); UnfixedCves = @() }
        }
        $cve = [pscustomobject]@{ Id = 'CVE-2025-0001'; Score = 7.5; Severity = 'HIGH'; Vector = ''; Published = '2025-01-01'; LastModified = ''; Summary = 'x'; References = @(); Epss = 0.1; EpssPercentile = 0.5; Kev = $true; KevDateAdded = '2025-02-01'; EuvdId = ''; Source = 'NVD'; Matches = @(); Links = (Get-VulnerabilityLinks -CveId 'CVE-2025-0001' -Kev).Cve }
        $r = [pscustomobject]@{
            Cpe = $Cpe; From = $CurrentVersion; To = $LatestVersion; Source = 'NVD'; QueriedAt = [DateTime]::UtcNow.AddHours(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')
            Count = 1; MaxCvss = 7.5; MaxSeverity = 'HIGH'; KevCount = 1; MaxEpss = 0.1; UnfixedCount = 2; Error = $null
            Links = (Get-VulnerabilityLinks -Cpe $Cpe -VersionStart $CurrentVersion -VendorAdvisoryUrl $VendorAdvisoryUrl).Product
            Cves = @($cve); UnfixedCves = @()
        }
        # The GUI reads a result back from the cache; a failed cache write leaves the old entry.
        if (-not $script:CveNoCache) {
            $cache = Read-VulnerabilityCache
            $cache[('{0}|{1}|{2}' -f $Cpe, $CurrentVersion, $LatestVersion)] = $r
            Save-VulnerabilityCache -Cache $cache
        }
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
    BeforeEach {
        $script:Log = @(); $script:LatestCalls = @(); $script:CmCalls = @(); $script:CveCalls = @(); $script:Notified = @(); $script:CveFail = $false; $script:CveNoCache = $false
        $script:Reports = Join-Path $TestDrive ('reports-' + [guid]::NewGuid().ToString('N'))
        $hist = Join-Path $env:LOCALAPPDATA 'AppPackager\app-history.json'
        if (Test-Path -LiteralPath $hist) { Remove-Item -LiteralPath $hist -Force }
        $cache = Join-Path $env:LOCALAPPDATA 'AppPackager\vuln-cache.json'
        if (Test-Path -LiteralPath $cache) { Remove-Item -LiteralPath $cache -Force }
    }

    It 'runs latest, ConfigMgr and CVE lookups and writes every report format' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -NvdApiKey 'k' -WebhookUrl 'https://hooks.example.test/x'
        $s.Total | Should -Be 4
        # Alpha and Gamma have an update; Gamma has no CPE, so only Alpha is looked up.
        $s.Updates | Should -Be 2
        $s.LookedUp | Should -Be 1
        $s.Cves | Should -Be 1
        $s.Kev | Should -Be 1
        $s.Failed | Should -Be 1
        $script:LatestCalls.Count | Should -Be 4
        $script:CmCalls.Count | Should -Be 3
        $script:CveCalls.Count | Should -Be 1
        $script:CveCalls[0].Cpe | Should -Be 'cpe:2.3:a:v:alpha:*:*:*:*:*:*:*:*'
        $script:CveCalls[0].From | Should -Be '1.0'
        $script:CveCalls[0].To | Should -Be '2.0'
        $script:CveCalls[0].Advisory | Should -Be 'https://v.example.test/security'
        $script:CveCalls[0].Key | Should -Be 'k'
        foreach ($p in 'Html', 'ApplicationsCsv', 'CvesCsv', 'Json', 'Xlsx', 'Markdown') { Test-Path -LiteralPath $s.Paths.$p | Should -BeTrue }
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').CVEs | Should -Be '1'
        ($apps | Where-Object Application -eq 'Alpha').Status | Should -Be 'Update available (1 CVE, 1 in KEV)'
        ($apps | Where-Object Application -eq 'Alpha').StillInLatest | Should -Be '2'
        ($apps | Where-Object Application -eq 'Alpha').VendorAdvisory | Should -Be 'https://v.example.test/security'
        ($apps | Where-Object Application -eq 'Beta').Status | Should -Be 'Up to date'
        ($apps | Where-Object Application -eq 'Gamma').Note | Should -Be 'CPE missing'
        ($apps | Where-Object Application -eq 'Delta').Status | Should -Be 'Error'
        $script:Notified.Count | Should -Be 1
        $script:Notified[0].Summary.Updates | Should -Be 2
    }

    It 'records the versions and the summary in the history file' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports
        $h = Read-PackagerHistory
        ([string]$h['package-alpha'].LastKnownVersion) | Should -Be '2.0'
        ([string]$h['package-alpha'].LastCmVersion) | Should -Be '1.0'
        ([string]$h['package-alpha'].LastCmSite) | Should -Be 'MCM'
        $h['package-alpha'].VulnSummary.Count | Should -Be 1
        $h['package-alpha'].VulnSummary.KevCount | Should -Be 1
        ([string]$h['package-beta'].LastCmVersion) | Should -Be '5.0'
        $h.ContainsKey('package-delta') | Should -BeTrue
        ([string]$h['package-delta'].LastResult) | Should -Be 'Failed'
        $h['package-delta'].LastKnownVersion | Should -BeNullOrEmpty
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
        $script:CveCalls[0].From | Should -Be '1.5'
        $script:CveCalls[0].To | Should -Be '3.0'
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        # Beta's ConfigMgr version came from another site and is not used.
        ($apps | Where-Object Application -eq 'Beta').ConfigMgrVersion | Should -Be ''
        ($apps | Where-Object Application -eq 'Beta').Status | Should -Be 'Latest retrieved'
    }

    It 'limits the run to -Apps and skips the lookup with -SkipCve' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha', 'package-beta.ps1' -SkipCve
        $s.Total | Should -Be 2
        $s.Updates | Should -Be 1
        $script:CveCalls.Count | Should -Be 0
        $script:LatestCalls.Count | Should -Be 2
    }

    It 'notifies only with an update unless told otherwise' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta' -WebhookUrl 'https://hooks.example.test/x'
        $script:Notified.Count | Should -Be 0
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-beta' -WebhookUrl 'https://hooks.example.test/x' -NotifyAlways
        $script:Notified.Count | Should -Be 1
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
        (Read-PackagerHistory)['package-alpha'].VulnSummary.Count | Should -Be 1
        $script:CveFail = $true
        try {
            $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
            $s.Failed | Should -Be 1
            (Read-PackagerHistory)['package-alpha'].VulnSummary.Count | Should -Be 1
        }
        finally { $script:CveFail = $false }
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
        $script:NoteCpe = 'cpe:2.3:a:vendor:product:*:*:*:*:*:*:*:*'
        $script:OldStamp = [DateTime]::UtcNow.AddHours(-3).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $script:NewStamp = [DateTime]::UtcNow.AddHours(-2).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $result = [pscustomobject]@{
            Cpe = $script:NoteCpe; From = '1.0'; To = '2.0'; Source = 'NVD'; QueriedAt = $script:OldStamp
            Count = 2; MaxCvss = 7.5; MaxSeverity = 'HIGH'; KevCount = 0; MaxEpss = $null; Error = $null
            FixedInLatestCount = 1; FixedEarlierCount = 1; FixedUnknownCount = 0; NewestPublished = '2026-09-01'; LastNvdUpdate = '2026-09-02'
            Links = @{}; Cves = @(); UnfixedCves = @()
        }
        Save-VulnerabilityCache -Cache @{ ($script:NoteCpe + '|1.0|2.0') = $result } -CachePath (Get-VulnerabilityCachePath)
        $row = { param($count, $checked) [pscustomobject]@{ Application = 'App'; Vendor = 'V'; Script = 'package-app.ps1'; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available'; Cpe = $script:NoteCpe; CveCount = $count; CveChecked = $checked } }
        $script:NoteRows = Get-CveReportRows -Rows @((& $row 'Error' $script:OldStamp), (& $row '2' $script:NewStamp), (& $row '2' $script:OldStamp))
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
    BeforeEach {
        $script:Log = @(); $script:LatestCalls = @(); $script:CmCalls = @(); $script:CveCalls = @(); $script:Notified = @(); $script:CveFail = $false; $script:CveNoCache = $false
        $script:Reports = Join-Path $TestDrive ('reports-' + [guid]::NewGuid().ToString('N'))
        $hist = Join-Path $env:LOCALAPPDATA 'AppPackager\app-history.json'
        if (Test-Path -LiteralPath $hist) { Remove-Item -LiteralPath $hist -Force }
        $cache = Join-Path $env:LOCALAPPDATA 'AppPackager\vuln-cache.json'
        if (Test-Path -LiteralPath $cache) { Remove-Item -LiteralPath $cache -Force }
    }

    It 'reports the fresh result when the cache write failed, not an older entry' {
        $alpha = 'cpe:2.3:a:v:alpha:*:*:*:*:*:*:*:*'
        $old = [pscustomobject]@{
            Cpe = $alpha; From = '1.0'; To = '2.0'; Source = 'NVD'; QueriedAt = [DateTime]::UtcNow.AddHours(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
            Count = 5; MaxCvss = 9.1; MaxSeverity = 'CRITICAL'; KevCount = 0; MaxEpss = $null; Error = $null
            FixedInLatestCount = 0; FixedEarlierCount = 5; FixedUnknownCount = 0; NewestPublished = '2020-01-01'; LastNvdUpdate = '2020-01-01'
            Links = @{}; Cves = @(); UnfixedCves = @()
        }
        Save-VulnerabilityCache -Cache @{ ($alpha + '|1.0|2.0') = $old } -CachePath (Get-VulnerabilityCachePath) -KeepDays 0
        $script:CveNoCache = $true
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').CVEs | Should -Be '1'
        ($apps | Where-Object Application -eq 'Alpha').Note | Should -Be ''
    }

    It 'marks a cached result for an application the run did not look up' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipCve -SkipLatest -SkipMecm
        $script:CveCalls.Count | Should -Be 1
        $apps = @(Import-Csv -LiteralPath $s.Paths.ApplicationsCsv)
        ($apps | Where-Object Application -eq 'Alpha').Note | Should -Be 'From the cache; not looked up in this run'
    }
}

Describe 'A stored summary belongs to one product' {
    BeforeAll {
        $script:SummaryFor = { param($cpe) @{ QueriedAt = '2026-11-01T00:00:00Z'; From = '1.0'; To = '2.0'; Cpe = $cpe; Count = 3; KevCount = 0; MaxCvss = 7.5; MaxSeverity = 'HIGH'; MaxEpss = $null; FixedInLatestCount = 1; NewestPublished = ''; LastNvdUpdate = '' } }
    }
    It 'shows the count under the CPE it was looked up for' {
        $cells = Get-CveGridCells -Summary (& $script:SummaryFor 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*') -Cpe 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*' -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be '3'
    }
    It 'shows nothing after the CPE of the packager changed' {
        $cells = Get-CveGridCells -Summary (& $script:SummaryFor 'cpe:2.3:a:v:old:*:*:*:*:*:*:*:*') -Cpe 'cpe:2.3:a:v:new:*:*:*:*:*:*:*:*' -CurrentVersion '1.0' -LatestVersion '2.0'
        $cells.Count | Should -Be ''
        $cells.Checked | Should -Be ''
    }
    It 'still shows a summary stored before the CPE was recorded' {
        $legacy = & $script:SummaryFor 'x'
        $legacy.Remove('Cpe')
        (Get-CveGridCells -Summary $legacy -Cpe 'cpe:2.3:a:v:new:*:*:*:*:*:*:*:*' -CurrentVersion '1.0' -LatestVersion '2.0').Count | Should -Be '3'
    }
}

Describe 'One comparer decides whether an update exists' {
    It 'decides <Current> to <Latest> as update: <Update>, in the grid and in the CVE check' -ForEach @(
        @{ Current = '1.0'; Latest = '2.0'; Update = $true }
        @{ Current = '2.0'; Latest = '1.0'; Update = $false }
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
        $reason = Get-CveSkipReason -Cpe 'cpe:2.3:a:v:p:*:*:*:*:*:*:*:*' -Current $Current -Latest $Latest
        ($reason -eq '') | Should -Be $Update
        if (-not $Update) { $reason | Should -Be 'No update' }
    }
    It 'compares the strict form over all four parts' {
        Compare-SemVer -A '1.2.3.4' -B '1.2.3.5' -Strict | Should -Be (-1)
        Compare-SemVer -A '1.2.3.5' -B '1.2.3' -Strict | Should -Be 1
        Compare-SemVer -A '20250101000000.1' -B '20250101000000.1' -Strict | Should -Be 0
    }
    It 'names the missing input first' {
        Get-CveSkipReason -Cpe '' -Current '1' -Latest '2' | Should -Be 'CPE missing'
        Get-CveSkipReason -Cpe 'cpe:2.3:a:v:p' -Current '' -Latest '2' | Should -Be 'No ConfigMgr version'
        Get-CveSkipReason -Cpe 'cpe:2.3:a:v:p' -Current '1' -Latest '' | Should -Be 'No latest version'
    }
    It 'leaves no use of the vulnerability module comparer in the application script' {
        $script:Ast.Extent.Text | Should -Not -Match 'Compare-VulnerabilityVersionValue'
    }
}

Describe 'CVE cells follow the version pair they describe' {
    BeforeAll {
        $script:NewRow = {
            param($cur, $lat)
            $r = [pscustomobject]@{ Cpe = 'cpe:2.3:a:v:p:*:*:*:*:*:*:*:*'; CurrentVersion = $cur; LatestVersion = $lat; Status = 'Update available (3 CVEs, 1 in KEV)'
                CveCount = '3'; CveCvss = '7.5 HIGH'; CveKev = 'Yes'; CveEpss = '1.0%'; CveChecked = '2026-11-01T00:00:00Z'; CveFixed = '1 of 3'; CveNewest = '2026-01-01'; CveUpdated = '2026-01-02' }
            Set-CveCellsPair -Row $r
            $r
        }
    }
    It 'keeps the cells while the pair is the same, also when the text of a version differs' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.LatestVersion = '2.0+build5'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be '3'
        $r.Status | Should -Be 'Update available (3 CVEs, 1 in KEV)'
    }
    It 'clears the cells and the status suffix when the latest version changes' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.LatestVersion = '3.0'
        Clear-StaleCveCells -Row $r
        foreach ($p in 'CveCount', 'CveCvss', 'CveKev', 'CveEpss', 'CveChecked', 'CveFixed', 'CveNewest', 'CveUpdated') { $r.$p | Should -Be '' -Because $p }
        $r.Status | Should -Be 'Update available'
    }
    It 'clears the cells when the ConfigMgr version changes' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.CurrentVersion = '1.5'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be ''
    }
    It 'keeps the missing-CPE text of a row with no CPE' {
        $r = & $script:NewRow '1.0' '2.0'
        $r.Cpe = ''
        $r.LatestVersion = '3.0'
        Clear-StaleCveCells -Row $r
        $r.CveCount | Should -Be 'CPE missing'
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
        $script:SResult = [pscustomobject]@{
            Cpe = $script:SCpe; From = '1.0'; To = '2.0'; Source = 'NVD'; QueriedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
            Count = 4; MaxCvss = 5.0; MaxSeverity = 'MEDIUM'; KevCount = 0; MaxEpss = $null; Error = $null
            FixedInLatestCount = 4; FixedEarlierCount = 0; FixedUnknownCount = 0; NewestPublished = ''; LastNvdUpdate = ''
            Links = @{}; Cves = @(); UnfixedCves = @()
        }
        $cache = Join-Path $env:LOCALAPPDATA 'AppPackager\vuln-cache.json'
        if (Test-Path -LiteralPath $cache) { Remove-Item -LiteralPath $cache -Force }
    }
    AfterEach { $script:CveSessionResults = $null }

    It 'returns the session result when the cache holds none' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        (Get-CveCachedResult -Row (& $script:SRow '4')).Count | Should -Be 4
    }
    It 'ignores a session result for another CPE or version pair' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $other = & $script:SRow '4'
        $other.LatestVersion = '3.0'
        Get-CveCachedResult -Row $other | Should -BeNullOrEmpty
    }
    It 'keeps the failure note for a row whose latest lookup failed' {
        $script:CveSessionResults = @{ 'package-s.ps1' = $script:SResult }
        $rows = Get-CveReportRows -Rows @((& $script:SRow 'Error')) -Results $script:CveSessionResults
        $rows[0].Result.Count | Should -Be 4
        $rows[0].Note | Should -Match '^Last lookup failed'
    }
}
Describe 'The monitor counts cached results apart and logs the ConfigMgr query' {
    BeforeEach {
        $script:Log = @(); $script:LatestCalls = @(); $script:CmCalls = @(); $script:CveCalls = @(); $script:Notified = @(); $script:CveFail = $false; $script:CveNoCache = $false
        $script:Reports = Join-Path $TestDrive ('reports-' + [guid]::NewGuid().ToString('N'))
        foreach ($f in 'app-history.json', 'vuln-cache.json') {
            $p = Join-Path $env:LOCALAPPDATA ('AppPackager\' + $f)
            if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
        }
    }

    It 'reports the results taken from the cache when the CVE lookup is skipped' {
        $null = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $script:Log = @()
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha' -SkipCve
        $s.LookedUp | Should -Be 0
        $s.Cached | Should -Be 1
        ($script:Log | Where-Object { $_ -like '*Monitor finished*' }) | Should -Match '1 from the cache'
    }

    It 'reports no cached result when the run looked every application up itself' {
        $s = Invoke-CveMonitor -PackagersRoot 'C:\x' -SiteCode 'MCM' -ReportFolder $script:Reports -Apps 'package-alpha'
        $s.LookedUp | Should -Be 1
        $s.Cached | Should -Be 0
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
