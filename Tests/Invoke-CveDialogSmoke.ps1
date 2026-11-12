#Requires -Version 5.1
# Headless probe for the CVE detail window. Builds the real window from a
# lookup result saved in a test result cache; checks the head line, the
# product link strip, the CVE rows with their status and link sets, the
# not-checked states, the Refresh conditions, and the export.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$script:SharedStyleRootOverride = $root

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerVulnerability.psd1') -Force

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
# The window handlers are closures; a closure resolves a command in the global scope only.
foreach ($name in @('Expand-SharedStyle', 'Get-SharedStyleXaml', 'Show-CveDetailDialog', 'ConvertTo-CveDialogRows', 'Get-CveGridCells', 'ConvertTo-CveVersionKey', 'ConvertTo-CvePairKey', 'Get-CveSkipReason', 'Find-CveSessionResult', 'Get-CveCachedResult', 'Get-CveReportRows', 'Write-CveReportForRows', 'Open-ExternalUrl')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw "function $name not found" }
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$dataRoot = Join-Path $env:TEMP ('ap-cve-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $dataRoot -Force)
# The result cache is under LOCALAPPDATA. The probe points it at its own folder.
$savedLocalAppData = $env:LOCALAPPDATA
$env:LOCALAPPDATA = Join-Path $dataRoot 'LocalAppData'
[void](New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force)

try {
    if (-not (Get-VulnerabilityResultFolder).StartsWith($dataRoot, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'The result cache is not in the probe folder.' }

    $appCpe = 'cpe:2.3:a:examplevendor:exampleapp:*:*:*:*:*:*:*:*'
    $toolCpe = 'cpe:2.3:a:examplevendor:exampletool:*:*:*:*:*:*:*:*'
    $viewerCpe = 'cpe:2.3:a:examplevendor:exampleviewer:*:*:*:*:*:*:*:*'
    $appAdvisory = 'https://example.test/advisories'

    $newRecord = {
        param([string]$Id, $Score, [string]$Severity, [string]$Status, [string]$FixedIn = '', [switch]$FixedInLatest, [switch]$Kev, [string[]]$References = @())
        $rec = New-VulnerabilityRecord -Id $Id -Score $Score -Severity $Severity -Vector '' -Published '2026-09-30T10:00:00.000' -LastModified '2026-10-08T09:00:00.000' `
            -Summary ('Placeholder summary of ' + $Id) -References $References -Source 'NVD'
        $rec.Status = $Status
        $rec.FixedIn = $FixedIn
        $rec.FixedInLatest = [bool]$FixedInLatest
        $rec.Kev = [bool]$Kev
        $rec.Links = (Get-VulnerabilityLinks -CveId $Id -Kev:$Kev).Cve
        $rec
    }
    # A result as a lookup returns it. The module derives Mode and Examined from the version pair.
    $newResult = {
        param([string]$Cpe, [string]$From, [string]$To, [object[]]$Cves = @(), [object[]]$Unversioned = @(), [string]$QueriedAt = '2026-10-11T12:00:00Z', [string]$ErrorText = '')
        $fixed = @($Cves | Where-Object { $_.Status -eq 'Fixed' })
        $top = @($Cves | Sort-Object -Property Score -Descending | Select-Object -First 1)
        $data = [pscustomobject]@{
            Schema = 2; Cpe = $Cpe; From = $From; To = $To; Source = 'NVD'; QueriedAt = $QueriedAt; Error = $(if ($ErrorText) { $ErrorText } else { $null })
            Count = $Cves.Count; FixedCount = $fixed.Count; OpenCount = ($Cves.Count - $fixed.Count); UnversionedCount = $Unversioned.Count
            KevCount = @($Cves | Where-Object { $_.Kev }).Count; FixedKevCount = @($fixed | Where-Object { $_.Kev }).Count
            FixedInLatestCount = @($fixed | Where-Object { $_.FixedInLatest }).Count
            FixedEarlierCount = @($fixed | Where-Object { -not $_.FixedInLatest -and $_.FixedIn }).Count
            FixedUnknownCount = @($fixed | Where-Object { -not $_.FixedIn }).Count
            MaxCvss = $(if ($top.Count) { $top[0].Score } else { $null }); MaxSeverity = $(if ($top.Count) { $top[0].Severity } else { '' }); MaxEpss = $null
            NewestPublished = $(if ($Cves.Count) { '2026-09-30' } else { '' }); LastNvdUpdate = $(if ($Cves.Count) { '2026-10-08' } else { '' }); ProductRecords = 12
            Links = [ordered]@{ 'Vendor advisory' = $appAdvisory }
            Cves = @($Cves); UnversionedCves = @($Unversioned)
        }
        $result = ConvertTo-VulnerabilityResultObject -Result $data
        if (-not $result) { throw 'The vulnerability module does not read a result of layout 2.' }
        $result
    }
    # The head line of the window for a result.
    $headOf = {
        param($Result)
        $fix = ConvertTo-VulnerabilityFixText -Result $Result
        '{0}{1} Source {2}, checked {3}.{4}' -f (ConvertTo-VulnerabilityResultText -Result $Result), $(if ($fix) { ' ' + $fix + '.' } else { '' }), $Result.Source, $Result.QueriedAt,
            $(if ($Result.NewestPublished) { ' Newest CVE {0}, NVD updated {1}.' -f $Result.NewestPublished, $Result.LastNvdUpdate } else { '' })
    }
    $click = { param($Button) $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }
    $refreshNeeds = 'Refresh needs a CPE header and a ConfigMgr version or a latest version.'
    $notChecked = 'Not checked. Run Check CVEs for this app. A lookup needs the ConfigMgr version or the latest version.'

    # Two CVEs that the update fixes, one that it leaves, and one record that names no version.
    $fixedKev = & $newRecord -Id 'CVE-2099-0001' -Score 9.8 -Severity 'CRITICAL' -Status 'Fixed' -FixedIn '2.0' -FixedInLatest -Kev -References @('https://example.test/advisory/1')
    $fixedKev.Epss = 0.5
    $fixedEarlier = & $newRecord -Id 'CVE-2099-0002' -Score 7.5 -Severity 'HIGH' -Status 'Fixed' -FixedIn '1.5' -References @('http://example.test/advisory/2')
    $open = & $newRecord -Id 'CVE-2099-0003' -Score 5.3 -Severity 'MEDIUM' -Status 'Open'
    $unversioned = & $newRecord -Id 'CVE-2099-0004' -Score 6.1 -Severity 'MEDIUM' -Status 'Unversioned'
    $savedResult = & $newResult -Cpe $appCpe -From '1.0' -To '2.0' -Cves @($fixedKev, $fixedEarlier, $open) -Unversioned @($unversioned)
    Save-VulnerabilityCachedResult -Result $savedResult
    # The latest version alone: ConfigMgr holds no version of this application.
    $latestResult = & $newResult -Cpe $toolCpe -From '' -To '3.1' -Cves @(& $newRecord -Id 'CVE-2099-0005' -Score 4.3 -Severity 'MEDIUM' -Status 'Open')
    Save-VulnerabilityCachedResult -Result $latestResult

    function global:Get-CveReportFolder { Join-Path $dataRoot 'Reports' }
    function global:Add-LogLine { param([string]$Message) $script:Log += $Message }
    function global:Install-TitleBarDragFallback { param($Window) }
    function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
    # Refresh starts the background pipeline. The probe records the call and makes no lookup.
    function global:Test-PipelineBusy { [bool]$script:PipelineBusy }
    function global:Get-CveRunContext { @{ CveSession = $null; NvdApiKey = ''; CacheTtlMinutes = 360; KevCatalog = $null; ForceRefresh = $false } }
    function global:Invoke-MultiAppPipeline { param([string]$Operation, $Rows, $Context) $script:PipelineCalls += , @{ Operation = $Operation; Rows = @($Rows); Context = $Context } }
    $script:Log = @()
    $script:PipelineBusy = $false
    $script:PipelineCalls = @()
    $script:CveSessionResults = $null

    $appRow = [pscustomobject]@{ Script = 'package-exampleapp.ps1'; Application = 'Example App'; Vendor = 'Example Vendor'; CurrentVersion = '1.0'; LatestVersion = '2.0'; Status = 'Update available (fixes 2 of 3 CVEs, 1 in KEV)'; Cpe = $appCpe; CveCount = '3'; CveCvss = '9.8 CRITICAL'; CveKev = 'Yes'; CveChecked = '2026-10-11T12:00:00Z'; AdvisoryUrl = $appAdvisory }
    $latestRow = [pscustomobject]@{ Script = 'package-exampletool.ps1'; Application = 'Example Tool'; Vendor = 'Example Vendor'; CurrentVersion = ''; LatestVersion = '3.1'; Status = 'Ready'; Cpe = $toolCpe; CveCount = '1'; CveCvss = '4.3 MEDIUM'; CveKev = ''; CveChecked = '2026-10-11T12:00:00Z' }
    $uncheckedRow = [pscustomobject]@{ Script = 'package-exampleviewer.ps1'; Application = 'Example Viewer'; Vendor = 'Example Vendor'; CurrentVersion = '3.0'; LatestVersion = '3.1'; Status = 'Update available'; Cpe = $viewerCpe; CveCount = ''; CveCvss = ''; CveKev = ''; CveChecked = '' }
    # The cells of this row are from another version pair than the row holds.
    $otherPairRow = [pscustomobject]@{ Script = 'package-exampleapp.ps1'; Application = 'Example App'; Vendor = 'Example Vendor'; CurrentVersion = '1.0'; LatestVersion = '2.5'; Status = 'Update available'; Cpe = $appCpe; CveCount = '3'; CveCvss = ''; CveKev = ''; CveChecked = '2026-10-11T12:00:00Z' }
    $noVersionRow = [pscustomobject]@{ Script = 'package-exampleviewer.ps1'; Application = 'Example Viewer'; Vendor = 'Example Vendor'; CurrentVersion = ''; LatestVersion = ''; Status = 'Ready'; Cpe = $viewerCpe; CveCount = 'No version'; CveCvss = ''; CveKev = ''; CveChecked = '' }
    $noCpeRow = [pscustomobject]@{ Script = 'package-nocpe.ps1'; Application = 'NoCpe'; Vendor = 'V'; CurrentVersion = '1.0'; LatestVersion = '1.1'; Status = 'Update available'; Cpe = ''; CveCount = 'CPE missing'; CveCvss = ''; CveKev = ''; CveChecked = '' }

    $failures = @()
    $check = { param([string]$Name, [bool]$Ok) if (-not $Ok) { $script:failures += $Name }; '{0} {1}' -f $(if ($Ok) { 'PASS' } else { 'FAIL' }), $Name }

    & $check 'two results are in the test cache folder' (@(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'LocalAppData\AppPackager\vuln-results') -Filter '*.json' -File).Count -eq 2)

    Show-CveDetailDialog -Owner $null -Row $appRow -Probe {
        param($p)
        & $check 'title names the app' ($p.Title.Text -eq 'Example App (Example Vendor)')
        & $check 'head line counts the CVEs of the ConfigMgr version and what the update fixes' ($p.Range.Text -like '3 CVEs affect 1.0*1 in KEV*2.0 fixes 2 and leaves 1*')
        & $check 'head line is the result text, the fix text, the source and the dates' ($p.Range.Text -ceq (& $headOf $savedResult))
        & $check 'head line gives the fix text' ($p.Range.Text -like '* Fixed in latest 1, earlier 1. Source NVD, checked 2026-10-11T12:00:00Z. Newest CVE 2026-09-30, NVD updated 2026-10-08.')
        & $check 'note hidden' ($p.Note.Visibility -eq 'Collapsed')
        $links = @($p.ProductLinks.ItemsSource)
        & $check 'six product links and the vendor advisory' ($links.Count -eq 7 -and @($links | Where-Object { $_.Name -eq 'Vendor advisory' -and $_.Url -eq $appAdvisory }).Count -eq 1)
        & $check 'product links are https' (@($links | Where-Object { $_.Url -notmatch '^https://' }).Count -eq 0)
        & $check 'NVD product link names the ConfigMgr version' (@($links | Where-Object { $_.Name -eq 'NVD' -and $_.Url -like '*exampleapp%3A1.0%3A*' }).Count -eq 1)
        & $check 'Status column follows the CVE column' ([string]$p.Grid.Columns[0].Header -eq 'CVE' -and [string]$p.Grid.Columns[1].Header -eq 'Status')
        $rows = @($p.Grid.ItemsSource)
        & $check 'three CVE rows and one row without a version' ($rows.Count -eq 4)
        & $check 'affecting CVEs first, the record without a version last' ((@($rows | ForEach-Object { $_.Id }) -join ' ') -eq 'CVE-2099-0001 CVE-2099-0002 CVE-2099-0003 CVE-2099-0004')
        & $check 'Status is Fixed, Open or No version' ((@($rows | ForEach-Object { $_.Status }) -join '|') -ceq 'Fixed|Fixed|Open|No version')
        & $check 'KEV row first' ($rows[0].Kev -eq 'Yes' -and $rows[1].Kev -eq '')
        & $check 'KEV row carries the KEV link' (@($rows[0].Links | Where-Object Name -eq 'KEV').Count -eq 1 -and @($rows[1].Links | Where-Object Name -eq 'KEV').Count -eq 0)
        & $check 'every row has the six site links' (@($rows | Where-Object { @($_.Links | Where-Object { $_.Name -in 'NVD', 'CVE.org', 'EUVD', 'OpenCVE', 'Shodan CVEDB', 'cvedetails' }).Count -ne 6 }).Count -eq 0)
        & $check 'only an https reference becomes a link' (@($rows[0].Links | Where-Object { $_.Name -eq 'Reference' -and $_.Url -eq 'https://example.test/advisory/1' }).Count -eq 1 -and @($rows[1].Links | Where-Object Name -eq 'Reference').Count -eq 0)
        & $check 'fix version marks the latest version' ($rows[0].FixedIn -eq '2.0 (latest)' -and $rows[1].FixedIn -eq '1.5' -and $rows[2].FixedIn -eq '')
        & $check 'summary text present' ($rows[0].Summary -eq 'Placeholder summary of CVE-2099-0001')
        & $check 'cvss text' ($rows[0].Cvss -eq '9.8 CRITICAL' -and $rows[3].Cvss -eq '6.1 MEDIUM')
        & $check 'epss text' ($rows[0].Epss -match '^50' -and $rows[1].Epss -eq '')
        & $check 'dates are days' ($rows[0].Published -eq '2026-09-30' -and $rows[0].Modified -eq '2026-10-08')

        # Refresh: another run is active.
        $refresh = $p.Window.FindName('btnRefresh')
        $script:PipelineBusy = $true
        & $click $refresh
        & $check 'Refresh waits for another run' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'Another run is in progress. Wait for it to finish.' -and $script:PipelineCalls.Count -eq 0 -and $refresh.IsEnabled)
        # Refresh: the lookup starts.
        $script:PipelineBusy = $false
        & $click $refresh
        & $check 'Refresh starts one forced lookup of the row' ($script:PipelineCalls.Count -eq 1 -and $script:PipelineCalls[0].Operation -eq 'CheckCve' -and $script:PipelineCalls[0].Rows.Count -eq 1 -and
            [object]::ReferenceEquals($script:PipelineCalls[0].Rows[0], $appRow) -and $script:PipelineCalls[0].Context.ForceRefresh -eq $true)
        & $check 'Refresh shows the lookup state' ($p.Note.Text -eq 'Looking up...' -and -not $refresh.IsEnabled)
        # The run reports a new result: the update now fixes every CVE.
        $fresh = & $newResult -Cpe $appCpe -From '1.0' -To '2.0' -Cves @($fixedKev, $fixedEarlier) -QueriedAt '2026-10-12T08:00:00Z'
        & $script:PipelineCalls[0].Context.CveDone ([pscustomobject]@{ CveResults = @{ 'package-exampleapp.ps1' = $fresh } })
        & $check 'the window shows the result of the Refresh run' ($p.Range.Text -ceq (& $headOf $fresh) -and $p.Range.Text -like '2 CVEs affect 1.0*fixes all*checked 2026-10-12T08:00:00Z*' -and @($p.Grid.ItemsSource).Count -eq 2)
        & $check 'Refresh is enabled again' ($refresh.IsEnabled -and $p.Note.Visibility -eq 'Collapsed')
        # The run reports a failed lookup: the cached result stays.
        & $click $refresh
        $appRow.CveCount = 'Error'
        $failed = & $newResult -Cpe $appCpe -From '1.0' -To '2.0' -ErrorText 'The lookup failed.' -QueriedAt '2026-10-12T09:00:00Z'
        & $script:PipelineCalls[1].Context.CveDone ([pscustomobject]@{ CveResults = @{ 'package-exampleapp.ps1' = $failed } })
        & $check 'a failed Refresh keeps the last good result' ($p.Range.Text -ceq (& $headOf $savedResult) -and @($p.Grid.ItemsSource).Count -eq 4)
        & $check 'a failed Refresh says so' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'The lookup failed; the last good result is shown. See the log.' -and $refresh.IsEnabled)
        $appRow.CveCount = '3'
        $p.Window.Close()
    }

    # A lookup of this session comes before the cached result of the same pair.
    $sessionResult = & $newResult -Cpe $appCpe -From '1.0' -To '2.0' -Cves @($open) -QueriedAt '2026-10-12T10:00:00Z'
    $script:CveSessionResults = @{ 'package-exampleapp.ps1' = $sessionResult }
    Show-CveDetailDialog -Owner $null -Row $appRow -Probe {
        param($p)
        & $check 'a session result comes before the cache' ($p.Range.Text -ceq (& $headOf $sessionResult) -and $p.Range.Text -like '1 CVE affects 1.0*fixes none*')
        & $check 'head line has no fix text when the update fixes none' ($p.Range.Text -notlike '*Fixed in latest*')
        $rows = @($p.Grid.ItemsSource)
        & $check 'one open CVE row from the session result' ($rows.Count -eq 1 -and $rows[0].Status -ceq 'Open' -and $rows[0].Id -eq 'CVE-2099-0003')
        $p.Window.Close()
    }
    $script:CveSessionResults = $null

    Show-CveDetailDialog -Owner $null -Row $latestRow -Probe {
        param($p)
        & $check 'head line names the latest version when ConfigMgr holds none' ($p.Range.Text -ceq (& $headOf $latestResult) -and $p.Range.Text -like '1 CVE affects the latest version 3.1.*')
        $rows = @($p.Grid.ItemsSource)
        & $check 'the CVE of the latest version is Open' ($rows.Count -eq 1 -and $rows[0].Status -ceq 'Open' -and $rows[0].FixedIn -eq '')
        & $check 'NVD product link names the latest version' (@(@($p.ProductLinks.ItemsSource) | Where-Object { $_.Name -eq 'NVD' -and $_.Url -like '*exampletool%3A3.1%3A*' }).Count -eq 1)
        # One version is enough for Refresh.
        $before = $script:PipelineCalls.Count
        & $click $p.Window.FindName('btnRefresh')
        & $check 'Refresh accepts a row with the latest version alone' ($script:PipelineCalls.Count -eq ($before + 1) -and [object]::ReferenceEquals($script:PipelineCalls[$before].Rows[0], $latestRow) -and $p.Note.Text -eq 'Looking up...')
        $p.Window.Close()
    }

    # NVD holds no record of the product: the head line says so and gives no count.
    $unknownResult = & $newResult -Cpe $viewerCpe -From '3.0' -To '3.1' -QueriedAt '2026-10-12T11:00:00Z'
    $unknownResult.ProductRecords = 0
    $unknownRow = [pscustomobject]@{ Script = 'package-exampleviewer.ps1'; Application = 'Example Viewer'; Vendor = 'Example Vendor'; CurrentVersion = '3.0'; LatestVersion = '3.1'; Status = 'Update available (product not in NVD)'; Cpe = $viewerCpe; CveCount = 'Not in NVD'; CveCvss = ''; CveKev = ''; CveChecked = '2026-10-12T11:00:00Z' }
    $script:CveSessionResults = @{ 'package-exampleviewer.ps1' = $unknownResult }
    Show-CveDetailDialog -Owner $null -Row $unknownRow -Probe {
        param($p)
        & $check 'head line says that NVD holds no record of the product' ($p.Range.Text -ceq (& $headOf $unknownResult) -and $p.Range.Text -like 'NVD holds no CVE record under the product name of the CPE header.*Source NVD, checked 2026-10-12T11:00:00Z.')
        & $check 'head line gives no count for a product NVD does not know' ($p.Range.Text -notlike '*No CVE affects*')
        & $check 'no rows and no note for a product NVD does not know' (@($p.Grid.ItemsSource).Count -eq 0 -and $p.Note.Visibility -eq 'Collapsed')
        $p.Window.Close()
    }
    $script:CveSessionResults = $null
    # The same row with no saved result: the text of the cell is not a reason.
    Show-CveDetailDialog -Owner $null -Row $unknownRow -Probe {
        param($p)
        & $check 'a Not in NVD cell with no saved result shows as not checked' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -ceq $notChecked)
        $p.Window.Close()
    }

    Show-CveDetailDialog -Owner $null -Row $uncheckedRow -Probe {
        param($p)
        & $check 'head line of a row with no result gives both versions' ($p.Range.Text -eq 'ConfigMgr 3.0, latest 3.1')
        & $check 'note of a row that was never checked' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -ceq $notChecked)
        & $check 'no rows without a result' (@($p.Grid.ItemsSource).Count -eq 0)
        $links = @($p.ProductLinks.ItemsSource)
        & $check 'six product links without a result' ($links.Count -eq 6 -and @($links | Where-Object { $_.Name -eq 'NVD' -and $_.Url -like '*exampleviewer%3A3.0%3A*' }).Count -eq 1)
        $p.Window.Close()
    }

    Show-CveDetailDialog -Owner $null -Row $otherPairRow -Probe {
        param($p)
        & $check 'a result of another version pair is not shown' ($p.Range.Text -eq 'ConfigMgr 1.0, latest 2.5' -and @($p.Grid.ItemsSource).Count -eq 0)
        & $check 'a count with no result gives the not-checked note' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -ceq $notChecked)
        $p.Window.Close()
    }

    Show-CveDetailDialog -Owner $null -Row $noVersionRow -Probe {
        param($p)
        & $check 'head line of a row with no version' ($p.Range.Text -eq 'ConfigMgr -, latest -')
        & $check 'note of a row with no version' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'No version')
        $before = $script:PipelineCalls.Count
        & $click $p.Window.FindName('btnRefresh')
        & $check 'Refresh needs one version' ($p.Note.Text -eq $refreshNeeds -and $script:PipelineCalls.Count -eq $before)
        $p.Window.Close()
    }

    Show-CveDetailDialog -Owner $null -Row $noCpeRow -Probe {
        param($p)
        & $check 'not-checked note shown' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'CPE missing')
        & $check 'no rows without a result' (@($p.Grid.ItemsSource).Count -eq 0)
        & $check 'no product links without a CPE' (@($p.ProductLinks.ItemsSource).Count -eq 0)
        $before = $script:PipelineCalls.Count
        & $click $p.Window.FindName('btnRefresh')
        & $check 'Refresh needs a CPE header' ($p.Note.Text -eq $refreshNeeds -and $script:PipelineCalls.Count -eq $before)
        $p.Window.Close()
    }

    $paths = Write-CveReportForRows -Rows @($appRow, $noCpeRow) -Name 'probe'
    & $check 'html written' (Test-Path -LiteralPath $paths.Html)
    $cveLines = @(Import-Csv -LiteralPath $paths.CvesCsv)
    & $check 'cves csv has the three CVEs and the record without a version' ($cveLines.Count -eq 4 -and (@($cveLines | ForEach-Object { $_.Status }) -join '|') -ceq 'Fixed|Fixed|Open|Unversioned')
    $appLines = @(Import-Csv -LiteralPath $paths.ApplicationsCsv)
    $appLine = $appLines | Where-Object Application -eq 'Example App'
    & $check 'applications csv counts the CVEs and the fixed CVEs' ($appLine.ExaminedVersion -eq '1.0' -and $appLine.CVEs -eq '3' -and $appLine.FixedByUpdate -eq '2' -and $appLine.Open -eq '1' -and $appLine.Unversioned -eq '1' -and $appLine.Note -eq '')
    & $check 'applications csv notes the missing CPE' (($appLines | Where-Object Application -eq 'NoCpe').Note -eq 'CPE missing')

    Open-ExternalUrl -Url 'http://example.test/'
    & $check 'http link refused' (@($script:Log | Where-Object { $_ -like 'Link not opened*' }).Count -eq 1)

    $cells = Get-CveGridCells -Summary (Get-VulnerabilitySummary -Result $savedResult) -Cpe $appCpe -CurrentVersion '1.0' -LatestVersion '2.0'
    & $check 'grid cells from the summary of the saved result' ($cells.Count -eq '3' -and $cells.Fixed -eq '2 of 3' -and $cells.Text -eq 'fixes 2 of 3 CVEs, 1 in KEV' -and $cells.Cvss -eq '9.8 CRITICAL' -and $cells.Kev -eq 'Yes' -and $cells.Checked -eq '2026-10-11T12:00:00Z')
    $cells = Get-CveGridCells -Summary (Get-VulnerabilitySummary -Result $savedResult) -Cpe $appCpe -CurrentVersion '1.0' -LatestVersion '2.5'
    & $check 'grid cells are empty for another version pair' ($cells.Count -eq '' -and $cells.Checked -eq '')
    $cells = Get-CveGridCells -Summary @{ Schema = 2; Cpe = 'cpe:2.3:a:x:y'; Mode = 'Deployed'; QueriedAt = 'q'; Count = 2; MaxCvss = 9.8; MaxSeverity = 'CRITICAL'; KevCount = 0 } -Cpe 'cpe:2.3:a:x:y'
    & $check 'grid cells from a hashtable summary' ($cells.Count -eq '2' -and $cells.Cvss -eq '9.8 CRITICAL' -and $cells.Kev -eq '' -and $cells.Fixed -eq '' -and $cells.Text -eq '2 CVEs' -and $cells.Checked -eq 'q')
    $cells = Get-CveGridCells -Summary @{ Cpe = 'cpe:2.3:a:x:y'; QueriedAt = 'q'; Count = 2; MaxCvss = 9.8; MaxSeverity = 'CRITICAL'; KevCount = 0 } -Cpe 'cpe:2.3:a:x:y'
    & $check 'grid cells are empty for a summary of the former layout' ($cells.Count -eq '' -and $cells.Cvss -eq '' -and $cells.Checked -eq '')
    $cells = Get-CveGridCells -Summary @{ Schema = 2; Cpe = 'cpe:2.3:a:x:y'; QueriedAt = 'q'; Error = 'boom' } -Cpe 'cpe:2.3:a:x:y'
    & $check 'grid cells show Error' ($cells.Count -eq 'Error' -and $cells.Checked -eq 'q')

    if ($failures.Count) { throw ('CVE dialog probe failed: ' + ($failures -join '; ')) }
    'CVE dialog probe passed.'
}
finally {
    $env:LOCALAPPDATA = $savedLocalAppData
    Remove-Item -LiteralPath $dataRoot -Recurse -Force -ErrorAction SilentlyContinue
}
