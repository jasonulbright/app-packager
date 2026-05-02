# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

function Read-MonitorConfig {
    param([Parameter(Mandatory)][string]$ConfigPath)

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        # First-run bootstrap: write a default config so the repo can ship
        # JSON-free. User edits the file afterward to point at their site.
        $default = [ordered]@{
            SchemaVersion = 1
            MECM = [ordered]@{
                SiteCode   = 'MCM'
                ServerFQDN = 'sccm01.contoso.com'
            }
            DownloadRoot = 'c:\temp\vm'
            NVD = [ordered]@{
                ApiKey             = ''
                RateLimitPerWindow = 5
                WindowSeconds      = 30
                CacheTtlMinutes    = 360
            }
            Report = [ordered]@{
                OutputFolder    = ''
                FileNamePattern = 'VersionMonitor-{0}.html'
                KeepReportDays  = 90
            }
            Logging = [ordered]@{
                LogFolder    = ''
                KeepLogDays  = 30
            }
            Notifications = [ordered]@{
                Enabled     = $false
                DropFolder  = ''
                WebhookUrl  = ''
                OnlyOnStale = $true
            }
        }
        $parent = Split-Path -Parent $ConfigPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        $default | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
        Write-Host ("[Read-MonitorConfig] Wrote default config: {0}" -f $ConfigPath)
    }
    $json = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 -ErrorAction Stop
    $cfg = $json | ConvertFrom-Json
    if (-not $cfg.SchemaVersion) { throw "Invalid config: missing SchemaVersion" }

    if (-not $cfg.DownloadRoot) { $cfg | Add-Member -NotePropertyName DownloadRoot -NotePropertyValue 'c:\temp\vm' -Force }
    if (-not $cfg.NVD) { $cfg | Add-Member -NotePropertyName NVD -NotePropertyValue ([pscustomobject]@{ ApiKey = ''; RateLimitPerWindow = 5; WindowSeconds = 30; CacheTtlMinutes = 360 }) -Force }
    if (-not $cfg.NVD.RateLimitPerWindow) { $cfg.NVD | Add-Member -NotePropertyName RateLimitPerWindow -NotePropertyValue 5 -Force }
    if (-not $cfg.NVD.WindowSeconds) { $cfg.NVD | Add-Member -NotePropertyName WindowSeconds -NotePropertyValue 30 -Force }
    if (-not $cfg.NVD.CacheTtlMinutes) { $cfg.NVD | Add-Member -NotePropertyName CacheTtlMinutes -NotePropertyValue 360 -Force }

    return $cfg
}

# ---------------------------------------------------------------------------
# Packager Discovery  (reads metadata + CPE/URL headers from packager scripts)
# ---------------------------------------------------------------------------

function Get-PackagerMetadata {
    param([Parameter(Mandatory)][string]$Path)

    $meta = [ordered]@{
        Vendor          = $null
        App             = $null
        CMName          = $null
        VendorUrl       = $null
        CPE             = $null
        ReleaseNotesUrl = $null
        DownloadPageUrl = $null
    }

    $lines = Get-Content -LiteralPath $Path -TotalCount 200 -ErrorAction Stop

    foreach ($line in $lines) {
        $l = $line.TrimStart([char]0xFEFF)

        if (-not $meta.Vendor          -and $l -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$')          { $meta.Vendor          = $Matches[1].Trim(); continue }
        if (-not $meta.App             -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')             { $meta.App             = $Matches[1].Trim(); continue }
        if (-not $meta.CMName          -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$')          { $meta.CMName          = $Matches[1].Trim(); continue }
        if (-not $meta.VendorUrl       -and $l -match '^\s*(?:#\s*)?VendorUrl\s*:\s*(.+?)\s*$')       { $meta.VendorUrl       = $Matches[1].Trim(); continue }
        if (-not $meta.CPE             -and $l -match '^\s*(?:#\s*)?CPE\s*:\s*(.+?)\s*$')             { $meta.CPE             = $Matches[1].Trim(); continue }
        if (-not $meta.ReleaseNotesUrl -and $l -match '^\s*(?:#\s*)?ReleaseNotesUrl\s*:\s*(.+?)\s*$') { $meta.ReleaseNotesUrl = $Matches[1].Trim(); continue }
        if (-not $meta.DownloadPageUrl -and $l -match '^\s*(?:#\s*)?DownloadPageUrl\s*:\s*(.+?)\s*$') { $meta.DownloadPageUrl = $Matches[1].Trim(); continue }

        if (-not $meta.App             -and $l -match '^\s*(?:#\s*)?Application\s*:\s*(.+?)\s*$')     { $meta.App             = $Matches[1].Trim(); continue }
    }

    if (-not $meta.CMName) { $meta.CMName = $meta.App }

    return [pscustomobject]@{
        Vendor          = $meta.Vendor
        Application     = $meta.App
        CMName          = $meta.CMName
        VendorUrl       = $meta.VendorUrl
        CPE             = $meta.CPE
        ReleaseNotesUrl = $meta.ReleaseNotesUrl
        DownloadPageUrl = $meta.DownloadPageUrl
        Script          = (Split-Path -Leaf $Path)
        FullPath        = $Path
    }
}

function Get-PackagerScripts {
    param([Parameter(Mandatory)][string]$PackagersRoot)

    if (-not (Test-Path -LiteralPath $PackagersRoot)) {
        throw "Packagers root not found: $PackagersRoot"
    }
    $scripts = Get-ChildItem -LiteralPath $PackagersRoot -Filter 'package-*.ps1' -File -ErrorAction Stop
    $results = @()
    foreach ($s in $scripts) {
        try {
            $results += Get-PackagerMetadata -Path $s.FullName
        }
        catch {
            Write-Log ("Failed to read metadata from {0}: {1}" -f $s.Name, $_.Exception.Message) -Level WARN
        }
    }
    return $results
}

# ---------------------------------------------------------------------------
# MECM Queries
# ---------------------------------------------------------------------------

function Get-MecmApplicationVersions {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string[]]$CMNames
    )

    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        try {
            if ($env:SMS_ADMIN_UI_PATH) {
                $cmModule = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
                if (Test-Path -LiteralPath $cmModule) {
                    Import-Module $cmModule -Force -ErrorAction Stop
                }
            }
        } catch { }
    }
    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        throw "ConfigMgr PowerShell cmdlets not available in this session."
    }

    $savedLocation = Get-Location
    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
    }
    catch {
        throw ("Failed to connect to CM site PSDrive '{0}:'" -f $SiteCode)
    }

    $results = @{}
    foreach ($cmName in $CMNames) {
        try {
            $apps = @(Get-CMApplication -Name $cmName -ErrorAction SilentlyContinue)
            if (-not $apps -or $apps.Count -eq 0) {
                $apps = @(Get-CMApplication -Name ("{0}*" -f $cmName) -ErrorAction SilentlyContinue)
            }

            if (-not $apps -or $apps.Count -eq 0) {
                $results[$cmName] = [pscustomobject]@{ Found = $false; DisplayName = $null; SoftwareVersion = $null; MatchCount = 0 }
                continue
            }

            $exact = $apps | Where-Object { $_.LocalizedDisplayName -eq $cmName -or $_.Name -eq $cmName }
            if ($exact -and @($exact).Count -gt 0) {
                $chosen = $exact | Select-Object -First 1
            }
            else {
                $parsable = @()
                $nonParsable = @()
                foreach ($a in $apps) {
                    try { $null = [version]$a.SoftwareVersion; $parsable += $a }
                    catch { $nonParsable += $a }
                }
                if ($parsable.Count -gt 0) {
                    $chosen = $parsable | Sort-Object { [version]$_.SoftwareVersion } -Descending | Select-Object -First 1
                }
                else {
                    $chosen = $nonParsable | Sort-Object Name -Descending | Select-Object -First 1
                }
            }

            $results[$cmName] = [pscustomobject]@{
                Found           = $true
                DisplayName     = $chosen.LocalizedDisplayName
                SoftwareVersion = $chosen.SoftwareVersion
                MatchCount      = $apps.Count
            }
        }
        catch {
            Write-Log ("MECM query failed for '{0}': {1}" -f $cmName, $_.Exception.Message) -Level WARN
            $results[$cmName] = [pscustomobject]@{ Found = $false; DisplayName = $null; SoftwareVersion = $null; MatchCount = 0 }
        }
    }

    Set-Location $savedLocation
    return $results
}

# ---------------------------------------------------------------------------
# Vendor Version Check
# ---------------------------------------------------------------------------

function Invoke-VendorVersionCheck {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$DownloadRoot = 'c:\temp\vm',
        [int]$TimeoutSeconds = 120
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -SiteCode "{1}" -DownloadRoot "{2}" -GetLatestVersionOnly' -f $PackagerPath, $SiteCode, $DownloadRoot
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    $null = $p.Start()
    $stderr = $p.StandardError.ReadToEnd()
    $stdout = $p.StandardOutput.ReadToEnd()

    if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
        try { $p.Kill() } catch { }
        throw "Timed out after ${TimeoutSeconds}s"
    }

    if ($p.ExitCode -ne 0) {
        $msg = $stderr
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = $stdout }
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Packager returned exit code $($p.ExitCode)." }
        throw $msg.Trim()
    }

    $lines = @($stdout -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $lines -or $lines.Count -lt 1) {
        throw "No version output received."
    }

    $version = ([string]$lines[0]).Trim()

    # Strip build metadata suffixes (e.g. "11.0.30+7", "2025.12-2", "2026.03.0-212")
    $version = $version -replace '[+-]\d+$', ''

    if ($version -notmatch '^\d+(\.\d+){1,3}$') {
        throw ("Unexpected version string: '{0}'" -f $version)
    }

    return $version
}

# ---------------------------------------------------------------------------
# Version Comparison
# ---------------------------------------------------------------------------

function Compare-Versions {
    param(
        [AllowEmptyString()][string]$MecmVersion,
        [AllowEmptyString()][string]$VendorVersion
    )

    if ([string]::IsNullOrWhiteSpace($MecmVersion) -or [string]::IsNullOrWhiteSpace($VendorVersion)) {
        return [pscustomobject]@{ Status = 'Unknown'; MecmParsed = $null; VendorParsed = $null }
    }

    try {
        $vm = [version]$MecmVersion
        $vv = [version]$VendorVersion
        $cmp = $vm.CompareTo($vv)
        $status = if ($cmp -ge 0) { 'Current' } else { 'Stale' }
        return [pscustomobject]@{ Status = $status; MecmParsed = $vm; VendorParsed = $vv }
    }
    catch {
        if ($MecmVersion -eq $VendorVersion) {
            return [pscustomobject]@{ Status = 'Current'; MecmParsed = $null; VendorParsed = $null }
        }
        return [pscustomobject]@{ Status = 'Unknown'; MecmParsed = $null; VendorParsed = $null }
    }
}

# ---------------------------------------------------------------------------
# NVD CVE Lookup
# ---------------------------------------------------------------------------

function Read-NvdCache {
    param([Parameter(Mandatory)][string]$CachePath)
    if (-not (Test-Path -LiteralPath $CachePath)) { return @{} }
    try {
        $json = Get-Content -LiteralPath $CachePath -Raw -Encoding UTF8
        $cache = @{}
        $obj = $json | ConvertFrom-Json
        foreach ($prop in $obj.PSObject.Properties) {
            $cache[$prop.Name] = $prop.Value
        }
        return $cache
    }
    catch { return @{} }
}

function Write-NvdCache {
    param(
        [Parameter(Mandatory)][string]$CachePath,
        [Parameter(Mandatory)][hashtable]$Cache
    )
    $Cache | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $CachePath -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Invoke-NvdCveQuery {
    param(
        [Parameter(Mandatory)][string]$CPE,
        [string]$VersionStart,
        [string]$VersionEnd,
        [string]$ApiKey
    )

    $baseUrl = 'https://services.nvd.nist.gov/rest/json/cves/2.0'
    $url = "{0}?virtualMatchString={1}&resultsPerPage=100" -f $baseUrl, [uri]::EscapeDataString($CPE)
    if ($VersionStart) {
        $url += "&versionStart={0}&versionStartType=including" -f [uri]::EscapeDataString($VersionStart)
    }
    if ($VersionEnd) {
        $url += "&versionEnd={0}&versionEndType=excluding" -f [uri]::EscapeDataString($VersionEnd)
    }
    if ($ApiKey) { $url += "&apiKey=$ApiKey" }

    $tmpFile = Join-Path $env:TEMP ("nvd-{0}.json" -f (Get-Random))
    try {
        $curlArgs = @('-L', '--fail', '--silent', '--show-error', '-o', $tmpFile, $url)
        & curl.exe @curlArgs 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "NVD API call failed (curl exit code $LASTEXITCODE)"
        }

        $json = Get-Content -LiteralPath $tmpFile -Raw -Encoding UTF8
        $response = $json | ConvertFrom-Json

        $cves = @()
        foreach ($vuln in $response.vulnerabilities) {
            $cve = $vuln.cve
            $score = $null; $severity = $null

            if ($cve.metrics.cvssMetricV31 -and $cve.metrics.cvssMetricV31.Count -gt 0) {
                $score = $cve.metrics.cvssMetricV31[0].cvssData.baseScore
                $severity = $cve.metrics.cvssMetricV31[0].cvssData.baseSeverity
            }
            elseif ($cve.metrics.cvssMetricV30 -and $cve.metrics.cvssMetricV30.Count -gt 0) {
                $score = $cve.metrics.cvssMetricV30[0].cvssData.baseScore
                $severity = $cve.metrics.cvssMetricV30[0].cvssData.baseSeverity
            }
            elseif ($cve.metrics.cvssMetricV2 -and $cve.metrics.cvssMetricV2.Count -gt 0) {
                $score = $cve.metrics.cvssMetricV2[0].cvssData.baseScore
                $severity = $cve.metrics.cvssMetricV2[0].baseSeverity
            }

            $cves += [pscustomobject]@{
                Id       = $cve.id
                Score    = $score
                Severity = $severity
                Url      = "https://nvd.nist.gov/vuln/detail/{0}" -f $cve.id
            }
        }

        $maxCve = $cves | Sort-Object Score -Descending | Select-Object -First 1
        return [pscustomobject]@{
            CVECount    = $cves.Count
            MaxCVSS     = if ($maxCve) { $maxCve.Score } else { $null }
            MaxSeverity = if ($maxCve) { $maxCve.Severity } else { $null }
            CVEs        = $cves
            Error       = $null
        }
    }
    catch {
        return [pscustomobject]@{
            CVECount    = 0
            MaxCVSS     = $null
            MaxSeverity = $null
            CVEs        = @()
            Error       = $_.Exception.Message
        }
    }
    finally {
        if (Test-Path -LiteralPath $tmpFile) { Remove-Item -LiteralPath $tmpFile -ErrorAction SilentlyContinue }
    }
}

function Invoke-NvdBatchQuery {
    param(
        [Parameter(Mandatory)][PSCustomObject[]]$StaleApps,
        [string]$ApiKey,
        [int]$RateLimit = 5,
        [int]$WindowSeconds = 30,
        [string]$CachePath,
        [int]$CacheTtlMinutes = 360
    )

    $cache = @{}
    if ($CachePath) { $cache = Read-NvdCache -CachePath $CachePath }
    $now = Get-Date

    $requestTimes = New-Object System.Collections.Generic.Queue[datetime]
    $results = @{}

    foreach ($app in $StaleApps) {
        $cpe = $app.CPE
        if ([string]::IsNullOrWhiteSpace($cpe)) {
            $results[$app.Script] = [pscustomobject]@{ CVECount = 0; MaxCVSS = $null; MaxSeverity = $null; CVEs = @(); Error = $null }
            continue
        }

        $cacheKey = "{0}|{1}|{2}" -f $cpe, $app.MecmVersion, $app.VendorVersion
        if ($cache.ContainsKey($cacheKey) -and $cache[$cacheKey].QueriedAt) {
            try {
                $cachedAt = [datetime]::Parse($cache[$cacheKey].QueriedAt)
                if (($now - $cachedAt).TotalMinutes -lt $CacheTtlMinutes) {
                    Write-Log ("NVD cache hit for {0}" -f $app.Script)
                    $results[$app.Script] = $cache[$cacheKey].Result
                    continue
                }
            }
            catch { }
        }

        while ($requestTimes.Count -ge $RateLimit) {
            $oldest = $requestTimes.Peek()
            $elapsed = ($now - $oldest).TotalSeconds
            if ($elapsed -ge $WindowSeconds) {
                $null = $requestTimes.Dequeue()
            }
            else {
                $sleepMs = [int](($WindowSeconds - $elapsed + 0.5) * 1000)
                Write-Log ("NVD rate limit: sleeping {0:N1}s" -f ($sleepMs / 1000))
                Start-Sleep -Milliseconds $sleepMs
                $now = Get-Date
            }
        }

        Write-Log ("Querying NVD for {0} ({1})" -f $app.Script, $cpe)
        $queryParams = @{ CPE = $cpe; ApiKey = $ApiKey }
        if ($app.MecmVersion) { $queryParams['VersionStart'] = $app.MecmVersion }
        if ($app.VendorVersion) { $queryParams['VersionEnd'] = $app.VendorVersion }
        $result = Invoke-NvdCveQuery @queryParams
        $requestTimes.Enqueue((Get-Date))
        $now = Get-Date

        if ($result.Error) {
            Write-Log ("NVD query failed for {0}: {1}" -f $app.Script, $result.Error) -Level WARN
        }
        else {
            Write-Log ("NVD: {0} has {1} CVEs, max CVSS {2}" -f $app.Script, $result.CVECount, $result.MaxCVSS)
        }

        $results[$app.Script] = $result

        $cache[$cacheKey] = [pscustomobject]@{
            QueriedAt = (Get-Date -Format 'o')
            Result    = $result
        }
    }

    if ($CachePath) { Write-NvdCache -CachePath $CachePath -Cache $cache }
    return $results
}

# ---------------------------------------------------------------------------
# HTML Report
# ---------------------------------------------------------------------------

function Export-VersionMonitorHtml {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$Results,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Vendor Version Monitor Report'
    )

    $currentCount = @($Results | Where-Object { $_.Status -eq 'Current' }).Count
    $staleCount   = @($Results | Where-Object { $_.Status -eq 'Stale' }).Count
    $errorCount   = @($Results | Where-Object { $_.Status -eq 'Error' -or $_.Status -like 'Error*' }).Count
    $unknownCount = @($Results | Where-Object { $_.Status -eq 'Unknown' -or $_.Status -eq 'Not in MECM' }).Count
    $totalCount   = $Results.Count

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; color: #333; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 16px; font-size: 0.95em; }',
        '.summary .count { font-weight: bold; }',
        '.summary .stale-count { color: #D83B01; }',
        '.summary .error-count { color: #A4262C; }',
        'table { border-collapse: collapse; margin-top: 8px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 10px; text-align: left; font-size: 0.85em; white-space: nowrap; }',
        'td { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; font-size: 0.9em; vertical-align: middle; white-space: nowrap; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        'tr.stale { background: #FFF0F0; }',
        'tr.stale:nth-child(even) { background: #FFE8E8; }',
        'tr.error { background: #FFF4F4; }',
        '.badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 0.8em; font-weight: bold; color: #fff; }',
        '.badge-current { background: #107C10; }',
        '.badge-stale { background: #D83B01; }',
        '.badge-unknown { background: #797775; }',
        '.badge-error { background: #A4262C; }',
        '.badge-notinmecm { background: #797775; }',
        '.cve-pill { display: inline-block; padding: 2px 8px; border-radius: 10px; background: #A4262C; color: #fff; font-size: 0.8em; font-weight: bold; text-decoration: none; }',
        '.cve-pill:hover { background: #8B0000; }',
        '.cvss-critical { color: #8B0000; font-weight: bold; }',
        '.cvss-high { color: #D83B01; font-weight: bold; }',
        '.cvss-medium { color: #CA5010; }',
        '.cvss-low { color: #107C10; }',
        '.links a { margin-right: 8px; color: #0078D4; text-decoration: none; font-size: 0.85em; }',
        '.links a:hover { text-decoration: underline; }',
        '.mono { font-family: Consolas, monospace; font-size: 0.9em; }',
        '</style>'
    ) -join "`r`n"

    $summaryHtml = (
        "<div class='summary'>Generated: {0} | " +
        "Applications: <span class='count'>{1}</span> | " +
        "Current: <span class='count'>{2}</span> | " +
        "Stale: <span class='count stale-count'>{3}</span> | " +
        "Unknown: <span class='count'>{4}</span> | " +
        "Error: <span class='count error-count'>{5}</span></div>"
    ) -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $totalCount, $currentCount, $staleCount, $unknownCount, $errorCount

    $headerRow = '<tr><th>Application</th><th>Publisher</th><th>MECM Version</th><th>Vendor Version</th><th>Status</th><th>CVEs</th><th>Max CVSS</th><th>Links</th></tr>'

    $rows = @()
    foreach ($r in ($Results | Sort-Object @{Expression={
        switch ($_.Status) { 'Stale' { 0 } 'Error' { 1 } 'Unknown' { 2 } 'Not in MECM' { 3 } default { 4 } }
    }}, @{Expression='Application'})) {
        $rowClass = switch -Wildcard ($r.Status) {
            'Stale'  { ' class="stale"' }
            'Error*' { ' class="error"' }
            default  { '' }
        }

        $badgeClass = switch -Wildcard ($r.Status) {
            'Current'      { 'badge-current' }
            'Stale'        { 'badge-stale' }
            'Not in MECM'  { 'badge-notinmecm' }
            'Error*'       { 'badge-error' }
            default        { 'badge-unknown' }
        }
        $statusBadge = "<span class='badge {0}'>{1}</span>" -f $badgeClass, [System.Net.WebUtility]::HtmlEncode($r.Status)

        $cveHtml = ''
        if ($null -ne $r.CVECount -and $r.CVECount -gt 0) {
            $cveUrl = if ($r.MaxCVEUrl) { $r.MaxCVEUrl } else { '#' }
            $cveHtml = "<a class='cve-pill' href='{0}' target='_blank' title='View highest-severity CVE'>{1}</a>" -f [System.Net.WebUtility]::HtmlEncode($cveUrl), $r.CVECount
        }
        elseif ($r.CVECount -eq 0 -and $r.Status -eq 'Stale') {
            $cveHtml = '0'
        }
        elseif ($null -ne $r.CVEError) {
            $cveHtml = '<span style="color:#999">N/A</span>'
        }

        $cvssHtml = ''
        if ($null -ne $r.MaxCVSS -and $r.MaxCVSS -gt 0) {
            $cvssClass = if ($r.MaxCVSS -ge 9.0) { 'cvss-critical' } elseif ($r.MaxCVSS -ge 7.0) { 'cvss-high' } elseif ($r.MaxCVSS -ge 4.0) { 'cvss-medium' } else { 'cvss-low' }
            $cveUrl = if ($r.MaxCVEUrl) { $r.MaxCVEUrl } else { '#' }
            $cvssHtml = "<a class='{0}' href='{1}' target='_blank' style='text-decoration:none' title='View CVE detail'>{2:N1} ({3})</a>" -f $cvssClass, [System.Net.WebUtility]::HtmlEncode($cveUrl), $r.MaxCVSS, $r.MaxSeverity
        }

        $links = @()
        if ($r.ReleaseNotesUrl) { $links += "<a href='{0}' target='_blank'>Release Notes</a>" -f [System.Net.WebUtility]::HtmlEncode($r.ReleaseNotesUrl) }
        if ($r.DownloadPageUrl) { $links += "<a href='{0}' target='_blank'>Download</a>" -f [System.Net.WebUtility]::HtmlEncode($r.DownloadPageUrl) }
        $linksHtml = "<div class='links'>{0}</div>" -f ($links -join '')

        $mecmVer = if ($r.MecmVersion) { $r.MecmVersion } else { '-' }
        $vendorVer = if ($r.VendorVersion) { $r.VendorVersion } else { '-' }

        $rows += "<tr{0}><td>{1}</td><td>{2}</td><td class='mono'>{3}</td><td class='mono'>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>" -f `
            $rowClass,
            [System.Net.WebUtility]::HtmlEncode($r.Application),
            [System.Net.WebUtility]::HtmlEncode($r.Publisher),
            [System.Net.WebUtility]::HtmlEncode($mecmVer),
            [System.Net.WebUtility]::HtmlEncode($vendorVer),
            $statusBadge,
            $cveHtml,
            $cvssHtml,
            $linksHtml
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8">',
        ("<title>{0}</title>" -f [System.Net.WebUtility]::HtmlEncode($ReportTitle)),
        $css,
        '</head><body>',
        ("<h1>{0}</h1>" -f [System.Net.WebUtility]::HtmlEncode($ReportTitle)),
        $summaryHtml,
        '<table>',
        ("<thead>{0}</thead>" -f $headerRow),
        '<tbody>',
        ($rows -join "`r`n"),
        '</tbody></table>',
        '</body></html>'
    ) -join "`r`n"

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }
    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8

    return $OutputPath
}

# ---------------------------------------------------------------------------
# Notifications
# ---------------------------------------------------------------------------

function Send-ReportNotification {
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)][PSCustomObject]$NotificationConfig,
        [int]$StaleCount = 0,
        [int]$TotalCount = 0
    )

    if (-not $NotificationConfig.Enabled) { return }

    if ($NotificationConfig.OnlyOnStale -and $StaleCount -eq 0) {
        Write-Log "No stale applications - skipping notification"
        return
    }

    if ($NotificationConfig.DropFolder -and (Test-Path -LiteralPath $NotificationConfig.DropFolder)) {
        try {
            Copy-Item -LiteralPath $ReportPath -Destination $NotificationConfig.DropFolder -Force -ErrorAction Stop
            Write-Log ("Report copied to drop folder: {0}" -f $NotificationConfig.DropFolder)
        }
        catch {
            Write-Log ("Failed to copy report to drop folder: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
    elseif ($NotificationConfig.DropFolder) {
        Write-Log ("Drop folder not accessible: {0}" -f $NotificationConfig.DropFolder) -Level WARN
    }

    if ($NotificationConfig.WebhookUrl) {
        Write-Log "Webhook notification not yet implemented - URL configured but skipped"
    }
}
