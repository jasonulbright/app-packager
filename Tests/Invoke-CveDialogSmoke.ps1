#Requires -Version 5.1
# Headless probe for the CVE detail window. Builds the real window from a
# fixture lookup result; checks the header line, the product link strip, the
# CVE rows with their link sets, the not-checked state, and the export.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerVulnerability.psd1') -Force

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('Show-CveDetailDialog', 'ConvertTo-CveDialogRows', 'Get-CveGridCells', 'Get-CveReportRows', 'Write-CveReportForRows', 'Open-ExternalUrl')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw "function $name not found" }
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$dataRoot = Join-Path $env:TEMP ('ap-cve-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $dataRoot -Force)

try {
    $fixtures = Join-Path $PSScriptRoot 'Fixtures\vuln'
    $cves = @(ConvertFrom-NvdResponse -Json ([System.IO.File]::ReadAllText((Join-Path $fixtures 'nvd-putty-0.78-0.83.json'))))
    foreach ($c in $cves) { $c.Links = (Get-VulnerabilityLinks -CveId $c.Id).Cve }
    $cves[0].Kev = $true
    $cves[0].Links = (Get-VulnerabilityLinks -CveId $cves[0].Id -Kev).Cve
    $script:Result = [pscustomobject]@{
        Cpe = 'cpe:2.3:a:putty:putty:*:*:*:*:*:*:*:*'; From = '0.78'; To = '0.83'; Source = 'NVD'; QueriedAt = '2026-10-11T12:00:00Z'
        Count = $cves.Count; MaxCvss = 5.9; MaxSeverity = 'MEDIUM'; KevCount = 1; MaxEpss = $null; Error = $null
        Links = (Get-VulnerabilityLinks -Cpe 'cpe:2.3:a:putty:putty:*:*:*:*:*:*:*:*' -VersionStart '0.78').Product
        Cves = $cves
    }
    function global:Get-CveCachedResult { param($Row) if ([string]$Row.Script -eq 'package-putty.ps1') { $script:Result } else { $null } }
    function global:Get-CveReportFolder { Join-Path $dataRoot 'Reports' }
    function global:Add-LogLine { param([string]$Message) $script:Log += $Message }
    function global:Install-TitleBarDragFallback { param($Window) }
    function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
    $script:Log = @()

    $putty = [pscustomobject]@{ Script = 'package-putty.ps1'; Application = 'PuTTY'; Vendor = 'Simon Tatham'; CurrentVersion = '0.78'; LatestVersion = '0.83'; Status = 'Update available'; Cpe = 'cpe:2.3:a:putty:putty:*:*:*:*:*:*:*:*'; CveCount = '5'; CveCvss = '5.9 MEDIUM'; CveKev = 'Yes'; CveChecked = '2026-10-11T12:00:00Z' }
    $nocpe = [pscustomobject]@{ Script = 'package-nocpe.ps1'; Application = 'NoCpe'; Vendor = 'V'; CurrentVersion = '1.0'; LatestVersion = '1.1'; Status = 'Update available'; Cpe = ''; CveCount = 'CPE missing'; CveCvss = ''; CveKev = ''; CveChecked = '' }

    $failures = @()
    $check = { param([string]$Name, [bool]$Ok) if (-not $Ok) { $script:failures += $Name }; '{0} {1}' -f $(if ($Ok) { 'PASS' } else { 'FAIL' }), $Name }

    Show-CveDetailDialog -Owner $null -Row $putty -Probe {
        param($p)
        & $check 'title names the app' ($p.Title.Text -eq 'PuTTY (Simon Tatham)')
        & $check 'range line counts and KEV' ($p.Range.Text -like '5 CVEs, 1 in KEV fixed by updating from 0.78 to 0.83.*NVD*')
        & $check 'note hidden' ($p.Note.Visibility -eq 'Collapsed')
        $links = @($p.ProductLinks.ItemsSource)
        & $check 'six product links' ($links.Count -eq 6)
        & $check 'product links are https' (@($links | Where-Object { $_.Url -notmatch '^https://' }).Count -eq 0)
        $rows = @($p.Grid.ItemsSource)
        & $check 'five CVE rows' ($rows.Count -eq 5)
        & $check 'KEV row first' ($rows[0].Kev -eq 'Yes' -and $rows[0].Id -eq $script:Result.Cves[0].Id)
        & $check 'KEV row carries the KEV link' (@($rows[0].Links | Where-Object Name -eq 'KEV').Count -eq 1)
        & $check 'every row has the six site links' (@($rows | Where-Object { @($_.Links | Where-Object { $_.Name -in 'NVD', 'CVE.org', 'EUVD', 'OpenCVE', 'Shodan CVEDB', 'cvedetails' }).Count -ne 6 }).Count -eq 0)
        & $check 'summary text present' (-not [string]::IsNullOrWhiteSpace($rows[0].Summary))
        & $check 'cvss text' ($rows[0].Cvss -match '^\d+\.\d [A-Z]+$')
        $p.Window.Close()
    }

    Show-CveDetailDialog -Owner $null -Row $nocpe -Probe {
        param($p)
        & $check 'not-checked note shown' ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'CPE missing')
        & $check 'no rows without a result' (@($p.Grid.ItemsSource).Count -eq 0)
        & $check 'no product links without a CPE' (@($p.ProductLinks.ItemsSource).Count -eq 0)
        $p.Window.Close()
    }

    $paths = Write-CveReportForRows -Rows @($putty, $nocpe) -Name 'probe'
    & $check 'html written' (Test-Path -LiteralPath $paths.Html)
    & $check 'cves csv has five rows' (@(Import-Csv -LiteralPath $paths.CvesCsv).Count -eq 5)
    & $check 'applications csv notes the missing CPE' ((Import-Csv -LiteralPath $paths.ApplicationsCsv | Where-Object Application -eq 'NoCpe').Note -eq 'CPE missing')

    Open-ExternalUrl -Url 'http://example.test/'
    & $check 'http link refused' (@($script:Log | Where-Object { $_ -like 'Link not opened*' }).Count -eq 1)

    $cells = Get-CveGridCells -Summary @{ QueriedAt = 'q'; Count = 2; MaxCvss = 9.8; MaxSeverity = 'CRITICAL'; KevCount = 0 } -Cpe 'cpe:2.3:a:x:y'
    & $check 'grid cells from a hashtable summary' ($cells.Count -eq '2' -and $cells.Cvss -eq '9.8 CRITICAL' -and $cells.Kev -eq '' -and $cells.Checked -eq 'q')
    $cells = Get-CveGridCells -Summary @{ QueriedAt = 'q'; Error = 'boom' } -Cpe 'cpe:2.3:a:x:y'
    & $check 'grid cells show Error' ($cells.Count -eq 'Error')

    if ($failures.Count) { throw ('CVE dialog probe failed: ' + ($failures -join '; ')) }
    'CVE dialog probe passed.'
}
finally {
    Remove-Item -LiteralPath $dataRoot -Recurse -Force -ErrorAction SilentlyContinue
}
