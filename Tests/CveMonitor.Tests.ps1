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
    foreach ($name in 'Invoke-CveMonitor', 'Get-CveGridCells', 'Get-CveCachedResult', 'Get-CveReportRows', 'Compare-SemVer') {
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
        # The monitor reads the report rows back from the cache the way the GUI does.
        $cache = Read-VulnerabilityCache
        $cache[('{0}|{1}|{2}' -f $Cpe, $CurrentVersion, $LatestVersion)] = $r
        Save-VulnerabilityCache -Cache $cache
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
        $script:Log = @(); $script:LatestCalls = @(); $script:CmCalls = @(); $script:CveCalls = @(); $script:Notified = @(); $script:CveFail = $false
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
