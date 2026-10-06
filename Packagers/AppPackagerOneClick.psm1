#Requires -Version 5.1
<#
    One Click planning and reporting. Every function here works on plain
    objects and files: no WPF, no site, no server. The GUI builds the plan
    window and the progress grid from Get-OneClickPlan and writes the run
    report through Write-OneClickReport.
#>

Set-StrictMode -Version 2.0

$script:OneClickDestinations = @('ConfigMgr', 'WSUS', 'Intune')

# ---------------------------------------------------------------------------
# Destinations
# ---------------------------------------------------------------------------

function Get-OneClickDestinationNames {
    <#
    .SYNOPSIS
        The destinations in the order a run publishes them.
    #>
    return @($script:OneClickDestinations)
}

function ConvertTo-OneClickDestinationSet {
    <#
    .SYNOPSIS
        Normalizes any destination description to the three flags.
    .DESCRIPTION
        Accepts a deployment target name (MECM, MECMAndIntune, IntuneOnly,
        MECMAndWSUS, WSUSOnly), an object with ConfigMgr, WSUS and Intune
        properties, or a list of destination names.
    #>
    param($Value)

    $set = [ordered]@{ ConfigMgr = $false; WSUS = $false; Intune = $false }
    if ($null -eq $Value) { return [pscustomobject]$set }

    if ($Value -is [string]) {
        switch ($Value) {
            'MECM'          { $set.ConfigMgr = $true }
            'MECMAndIntune' { $set.ConfigMgr = $true; $set.Intune = $true }
            'IntuneOnly'    { $set.Intune = $true }
            'MECMAndWSUS'   { $set.ConfigMgr = $true; $set.WSUS = $true }
            'WSUSOnly'      { $set.WSUS = $true }
            default         { $set.ConfigMgr = $true }
        }
        return [pscustomobject]$set
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [System.Collections.IDictionary]) {
        foreach ($name in $Value) {
            $key = Resolve-OneClickDestinationName -Name ([string]$name)
            if ($key) { $set[$key] = $true }
        }
        return [pscustomobject]$set
    }

    foreach ($key in @($set.Keys)) {
        $alias = @($key) + $(if ($key -eq 'WSUS') { @('Wsus') } elseif ($key -eq 'ConfigMgr') { @('MECM', 'Mecm') } else { @() })
        foreach ($name in $alias) {
            $prop = $null
            if ($Value -is [System.Collections.IDictionary]) { if ($Value.Contains($name)) { $prop = $Value[$name] } }
            elseif ($Value.PSObject.Properties[$name]) { $prop = $Value.PSObject.Properties[$name].Value }
            if ($null -ne $prop) { $set[$key] = [bool]$prop; break }
        }
    }
    return [pscustomobject]$set
}

function Resolve-OneClickDestinationName {
    param([string]$Name)
    switch -Regex ($Name) {
        '^(ConfigMgr|MECM)$' { return 'ConfigMgr' }
        '^WSUS$'             { return 'WSUS' }
        '^Intune$'           { return 'Intune' }
    }
    return $null
}

function Get-OneClickDestinations {
    <#
    .SYNOPSIS
        The destinations One Click publishes one application to.
    .DESCRIPTION
        A row with its own selection in AppFlow.Destinations wins. Any other
        row uses AppFlow.DefaultDestinations, or the saved deployment target
        of an older preferences file.
        Returns the flags and whether they came from the row.
    #>
    param(
        [Parameter(Mandatory)]$Prefs,
        [Parameter(Mandatory)][string]$PackagerName
    )

    $own = $null
    $appFlow = $null
    if ($Prefs.PSObject.Properties['AppFlow']) { $appFlow = $Prefs.AppFlow }
    if ($appFlow -and $appFlow.PSObject.Properties['Destinations'] -and $appFlow.Destinations) {
        $map = $appFlow.Destinations
        if ($map -is [System.Collections.IDictionary]) { if ($map.Contains($PackagerName)) { $own = $map[$PackagerName] } }
        elseif ($map.PSObject.Properties[$PackagerName]) { $own = $map.PSObject.Properties[$PackagerName].Value }
    }
    if ($null -ne $own) {
        $set = ConvertTo-OneClickDestinationSet -Value $own
        $set | Add-Member -NotePropertyName Source -NotePropertyValue 'Row' -Force
        return $set
    }

    if ($appFlow -and $appFlow.PSObject.Properties['DefaultDestinations'] -and $appFlow.DefaultDestinations) {
        $set = ConvertTo-OneClickDestinationSet -Value $appFlow.DefaultDestinations
    }
    else {
        $target = 'MECM'
        if ($Prefs.PSObject.Properties['Intune'] -and $Prefs.Intune -and $Prefs.Intune.PSObject.Properties['DeploymentTarget']) {
            $target = [string]$Prefs.Intune.DeploymentTarget
        }
        $set = ConvertTo-OneClickDestinationSet -Value $target
    }
    $set | Add-Member -NotePropertyName Source -NotePropertyValue 'Default' -Force
    return $set
}

function Test-OneClickDestinationInUse {
    <#
    .SYNOPSIS
        Whether a destination is selected under Systems in use and has the
        settings a publish needs.
    .OUTPUTS
        [pscustomobject] Ready, Reason.
    #>
    param(
        [Parameter(Mandatory)]$Prefs,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination
    )

    $systems = $null
    if ($Prefs.PSObject.Properties['Systems']) { $systems = $Prefs.Systems }
    $inUse = switch ($Destination) {
        'ConfigMgr' { -not $systems -or [bool]$systems.ConfigMgr }
        'WSUS'      { [bool]($systems -and $systems.Wsus) }
        'Intune'    { [bool]($systems -and $systems.Intune) }
    }
    if (-not $inUse) {
        return [pscustomobject]@{ Ready = $false; Reason = ('{0} is not in use' -f $Destination) }
    }
    switch ($Destination) {
        'ConfigMgr' {
            $console = [bool]($Prefs.PSObject.Properties['DetectedTools'] -and $Prefs.DetectedTools -and $Prefs.DetectedTools.ConfigMgrConsole -and $Prefs.DetectedTools.ConfigMgrConsole.Found)
            if (-not $console) { return [pscustomobject]@{ Ready = $false; Reason = 'ConfigMgr console not detected' } }
            if ([string]::IsNullOrWhiteSpace([string]$Prefs.SiteCode)) { return [pscustomobject]@{ Ready = $false; Reason = 'site code not set' } }
            if ([string]::IsNullOrWhiteSpace([string]$Prefs.FileShareRoot)) { return [pscustomobject]@{ Ready = $false; Reason = 'file share root not set' } }
        }
        'WSUS' {
            if (-not ($Prefs.PSObject.Properties['Wsus'] -and $Prefs.Wsus -and -not [string]::IsNullOrWhiteSpace([string]$Prefs.Wsus.ServerName))) {
                return [pscustomobject]@{ Ready = $false; Reason = 'WSUS server not set' }
            }
        }
        'Intune' {
            $intune = $null
            if ($Prefs.PSObject.Properties['Intune']) { $intune = $Prefs.Intune }
            $ready = [bool]($intune -and -not [string]::IsNullOrWhiteSpace([string]$intune.TenantId) -and
                -not [string]::IsNullOrWhiteSpace([string]$intune.ClientId) -and -not [string]::IsNullOrWhiteSpace([string]$intune.ClientSecretProtected))
            if (-not $ready) { return [pscustomobject]@{ Ready = $false; Reason = 'Intune credentials not set' } }
        }
    }
    return [pscustomobject]@{ Ready = $true; Reason = '' }
}

function Get-OneClickDestinationScope {
    <#
    .SYNOPSIS
        Where a publish lands, as the plan shows it: the ConfigMgr collection
        and DP group, the WSUS approval group, the Intune assignment.
    #>
    param(
        [Parameter(Mandatory)]$Prefs,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination
    )
    switch ($Destination) {
        'ConfigMgr' {
            $parts = @()
            $cd = $null
            if ($Prefs.PSObject.Properties['ContentDistribution']) { $cd = $Prefs.ContentDistribution }
            if ($cd -and [bool]$cd.AutoDistribute -and -not [string]::IsNullOrWhiteSpace([string]$cd.DPGroupName)) { $parts += ('distribute to DP group ' + $cd.DPGroupName) }
            if ($cd -and [bool]$cd.DeployToTestCollection -and -not [string]::IsNullOrWhiteSpace([string]$cd.TestCollectionName)) { $parts += ('deploy to collection ' + $cd.TestCollectionName) }
            if ($parts.Count -eq 0) { return 'application only, no distribution, no deployment' }
            return ($parts -join '; ')
        }
        'WSUS' {
            $group = ''
            if ($Prefs.PSObject.Properties['Wsus'] -and $Prefs.Wsus -and $Prefs.Wsus.PSObject.Properties['ApprovalGroup']) { $group = [string]$Prefs.Wsus.ApprovalGroup }
            if ([string]::IsNullOrWhiteSpace($group)) { return 'published, not approved' }
            return ('approved for group ' + $group.Trim())
        }
        'Intune' { return 'app created, not assigned' }
    }
}

# ---------------------------------------------------------------------------
# Publish history per destination
# ---------------------------------------------------------------------------

function Get-OneClickPublishedVersion {
    <#
    .SYNOPSIS
        The version last published to one destination, from a history entry.
    .OUTPUTS
        [pscustomobject] Version, At, Id; or $null when never published.
    #>
    param(
        $HistoryEntry,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination
    )
    if ($null -eq $HistoryEntry) { return $null }
    $published = $null
    if ($HistoryEntry -is [System.Collections.IDictionary]) { if ($HistoryEntry.Contains('Published')) { $published = $HistoryEntry['Published'] } }
    elseif ($HistoryEntry.PSObject.Properties['Published']) { $published = $HistoryEntry.Published }
    if ($null -eq $published) { return $null }
    $record = $null
    if ($published -is [System.Collections.IDictionary]) { if ($published.Contains($Destination)) { $record = $published[$Destination] } }
    elseif ($published.PSObject.Properties[$Destination]) { $record = $published.PSObject.Properties[$Destination].Value }
    if ($null -eq $record) { return $null }
    $get = { param($name) if ($record -is [System.Collections.IDictionary]) { if ($record.Contains($name)) { $record[$name] } } elseif ($record.PSObject.Properties[$name]) { $record.PSObject.Properties[$name].Value } }
    return [pscustomobject]@{
        Version = [string](& $get 'Version')
        At      = [string](& $get 'At')
        Id      = [string](& $get 'Id')
    }
}

function Set-OneClickPublishedVersion {
    <#
    .SYNOPSIS
        Records a publish against one destination in a history dictionary.
    .DESCRIPTION
        Mutates the hashtable that Read-PackagerHistory returns; the caller
        saves it. Entries read from JSON are objects; they are replaced by
        hashtables so that properties can be added.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$History,
        [Parameter(Mandatory)][string]$PackagerName,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination,
        [Parameter(Mandatory)][string]$Version,
        [string]$Id = '',
        [datetime]$At = (Get-Date)
    )
    if (-not $History.ContainsKey($PackagerName) -or $null -eq $History[$PackagerName]) {
        $History[$PackagerName] = @{ LastChecked = $null; LastStaged = $null; LastPackaged = $null; LastKnownVersion = $null; LastResult = $null }
    }
    $entry = $History[$PackagerName]
    if ($entry -isnot [hashtable]) {
        $h = @{}
        foreach ($p in $entry.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $entry = $h
        $History[$PackagerName] = $entry
    }
    $published = $null
    if ($entry.ContainsKey('Published')) { $published = $entry['Published'] }
    if ($null -eq $published) { $published = @{} }
    elseif ($published -isnot [hashtable]) {
        $h = @{}
        foreach ($p in $published.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $published = $h
    }
    $published[$Destination] = @{
        Version = $Version
        At      = $At.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Id      = $Id
    }
    $entry['Published'] = $published
    if ($entry.ContainsKey('NotSupported') -and $null -ne $entry['NotSupported']) {
        $refusals = ConvertTo-OneClickHashtable -Value $entry['NotSupported']
        $refusals.Remove($Destination)
        $entry['NotSupported'] = $refusals
    }
    return $History
}

function ConvertTo-OneClickHashtable {
    param($Value)
    if ($null -eq $Value) { return @{} }
    if ($Value -is [hashtable]) { return $Value }
    $h = @{}
    foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
    return $h
}

function Get-OneClickNotSupportedField {
    param($HistoryEntry, [string]$Destination, [string]$Field)
    if ($null -eq $HistoryEntry) { return '' }
    $refusals = $null
    if ($HistoryEntry -is [System.Collections.IDictionary]) { if ($HistoryEntry.Contains('NotSupported')) { $refusals = $HistoryEntry['NotSupported'] } }
    elseif ($HistoryEntry.PSObject.Properties['NotSupported']) { $refusals = $HistoryEntry.NotSupported }
    if ($null -eq $refusals) { return '' }
    $record = $null
    if ($refusals -is [System.Collections.IDictionary]) { if ($refusals.Contains($Destination)) { $record = $refusals[$Destination] } }
    elseif ($refusals.PSObject.Properties[$Destination]) { $record = $refusals.PSObject.Properties[$Destination].Value }
    if ($null -eq $record) { return '' }
    if ($record -is [System.Collections.IDictionary]) { return [string]$record[$Field] }
    if ($record.PSObject.Properties[$Field]) { return [string]$record.$Field }
    return ''
}

function Get-OneClickNotSupportedVersion {
    <#
    .SYNOPSIS
        The version a destination last refused, from a history entry.
    .OUTPUTS
        [string] The version; empty when the destination never refused.
    #>
    param(
        $HistoryEntry,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination
    )
    return (Get-OneClickNotSupportedField -HistoryEntry $HistoryEntry -Destination $Destination -Field 'Version')
}

function Get-OneClickNotSupportedReason {
    <#
    .SYNOPSIS
        The reason a destination gave for its last refusal, from a history entry.
    .OUTPUTS
        [string] The reason (for WSUS, "not supported: <code>"); empty when
        the destination never refused or recorded no reason.
    #>
    param(
        $HistoryEntry,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination
    )
    return (Get-OneClickNotSupportedField -HistoryEntry $HistoryEntry -Destination $Destination -Field 'Reason')
}

function Set-OneClickNotSupported {
    <#
    .SYNOPSIS
        Records that a destination refused one version of an application.
    .DESCRIPTION
        The plan and the run skip that destination for the same version
        until Force is set; a newer version is tried again. Mutates the
        hashtable that Read-PackagerHistory returns; the caller saves it.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$History,
        [Parameter(Mandatory)][string]$PackagerName,
        [Parameter(Mandatory)][ValidateSet('ConfigMgr', 'WSUS', 'Intune')][string]$Destination,
        [Parameter(Mandatory)][string]$Version,
        [string]$Reason = '',
        [datetime]$At = (Get-Date)
    )
    if (-not $History.ContainsKey($PackagerName) -or $null -eq $History[$PackagerName]) {
        $History[$PackagerName] = @{ LastChecked = $null; LastStaged = $null; LastPackaged = $null; LastKnownVersion = $null; LastResult = $null }
    }
    $entry = ConvertTo-OneClickHashtable -Value $History[$PackagerName]
    $History[$PackagerName] = $entry
    $refusals = @{}
    if ($entry.ContainsKey('NotSupported')) { $refusals = ConvertTo-OneClickHashtable -Value $entry['NotSupported'] }
    $refusals[$Destination] = @{
        Version = $Version
        At      = $At.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Reason  = $Reason
    }
    $entry['NotSupported'] = $refusals
    return $History
}

# ---------------------------------------------------------------------------
# Plan
# ---------------------------------------------------------------------------

function Get-OneClickCadenceDays {
    param($Prefs, [string]$PackagerName, [int]$HeaderDays = 0)
    $days = 7
    if ($HeaderDays -ge 1) { $days = $HeaderDays }
    $overrides = $null
    if ($Prefs.PSObject.Properties['AppFlow'] -and $Prefs.AppFlow -and $Prefs.AppFlow.PSObject.Properties['CadenceOverrides']) { $overrides = $Prefs.AppFlow.CadenceOverrides }
    if ($overrides) {
        $value = $null
        if ($overrides -is [System.Collections.IDictionary]) { if ($overrides.Contains($PackagerName)) { $value = $overrides[$PackagerName] } }
        elseif ($overrides.PSObject.Properties[$PackagerName]) { $value = $overrides.PSObject.Properties[$PackagerName].Value }
        $parsed = 0
        if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$parsed) -and $parsed -ge 1) { $days = $parsed }
    }
    return $days
}

function Get-OneClickPlan {
    <#
    .SYNOPSIS
        Builds the plan the One Click window shows before a run.
    .DESCRIPTION
        One plan row per tracked application. The row carries the
        destinations, the last version published to each, the planned steps
        and the reason for every skip. Nothing here runs a packager or
        contacts a server; the current version is the one the last check
        recorded, or empty when the run has to check first.
    .PARAMETER Apps
        Objects with Packager (base name), Application, Vendor, and
        optionally LatestVersion, CurrentVersion, CadenceDays (packager
        header), WsusUnsupported (bool).
    .PARAMETER History
        The dictionary Read-PackagerHistory returns.
    .PARAMETER Action
        Report, Stage or StageAndPackage.
    .PARAMETER Force
        Ignores the cadence and the duplicate guard.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Apps,
        [Parameter(Mandatory)]$Prefs,
        [hashtable]$History = @{},
        [ValidateSet('Report', 'Stage', 'StageAndPackage')][string]$Action = 'Report',
        [switch]$Force,
        [datetime]$Now = (Get-Date)
    )

    $tracked = @()
    if ($Prefs.PSObject.Properties['AppFlow'] -and $Prefs.AppFlow -and $Prefs.AppFlow.PSObject.Properties['Tracked']) { $tracked = @($Prefs.AppFlow.Tracked) }
    $trackedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $tracked) { [void]$trackedSet.Add([string]$t) }

    $readiness = @{}
    foreach ($d in $script:OneClickDestinations) { $readiness[$d] = Test-OneClickDestinationInUse -Prefs $Prefs -Destination $d }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($app in $Apps) {
        $name = [string]$app.Packager
        if (-not $trackedSet.Contains($name)) { continue }

        $entry = $null
        if ($History.ContainsKey($name)) { $entry = $History[$name] }
        $get = { param($obj, $prop) if ($null -eq $obj) { $null } elseif ($obj -is [System.Collections.IDictionary]) { if ($obj.Contains($prop)) { $obj[$prop] } } elseif ($obj.PSObject.Properties[$prop]) { $obj.PSObject.Properties[$prop].Value } }

        $latest = ''
        if ($app.PSObject.Properties['LatestVersion']) { $latest = [string]$app.LatestVersion }
        if ([string]::IsNullOrWhiteSpace($latest)) { $latest = [string](& $get $entry 'LastKnownVersion') }
        $current = ''
        if ($app.PSObject.Properties['CurrentVersion']) { $current = [string]$app.CurrentVersion }
        $headerDays = 0
        if ($app.PSObject.Properties['CadenceDays']) { $headerDays = [int]$app.CadenceDays }
        $wsusUnsupported = [bool]($app.PSObject.Properties['WsusUnsupported'] -and $app.WsusUnsupported)

        $destinations = Get-OneClickDestinations -Prefs $Prefs -PackagerName $name
        $lastPublished = [ordered]@{}
        $lastRefused = @{}
        foreach ($d in $script:OneClickDestinations) {
            $rec = Get-OneClickPublishedVersion -HistoryEntry $entry -Destination $d
            $lastPublished[$d] = $(if ($rec) { $rec.Version } else { '' })
            $lastRefused[$d] = Get-OneClickNotSupportedVersion -HistoryEntry $entry -Destination $d
        }

        $reasons = New-Object System.Collections.Generic.List[string]
        $publishTo = New-Object System.Collections.Generic.List[string]
        $notPublishable = New-Object System.Collections.Generic.List[string]
        $selected = @($script:OneClickDestinations | Where-Object { [bool]$destinations.$_ })

        $planned = ''
        $include = $true
        if ($selected.Count -eq 0) {
            $include = $false
            $planned = 'Skip'
            $reasons.Add('no destination selected')
        }
        else {
            $lastChecked = [string](& $get $entry 'LastChecked')
            $cadenceSkip = $false
            if ($Action -eq 'Report' -and -not $Force -and -not [string]::IsNullOrWhiteSpace($lastChecked)) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse($lastChecked, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
                    $days = Get-OneClickCadenceDays -Prefs $Prefs -PackagerName $name -HeaderDays $headerDays
                    if ($parsed.ToUniversalTime().AddDays($days) -gt $Now.ToUniversalTime()) {
                        $cadenceSkip = $true
                        $reasons.Add(('checked within the {0}-day cadence' -f $days))
                    }
                }
            }
            if ($cadenceSkip) {
                $include = $false
                $planned = 'Skip'
            }
            else {
                $steps = @('Check')
                if ($Action -in @('Stage', 'StageAndPackage')) { $steps += 'Stage' }
                if ($Action -eq 'StageAndPackage') {
                    foreach ($d in $selected) {
                        if (-not $readiness[$d].Ready) { $reasons.Add(('{0}: {1}' -f $d, $readiness[$d].Reason)); $notPublishable.Add($d); continue }
                        if ($d -eq 'WSUS' -and $wsusUnsupported) { $reasons.Add('WSUS: not supported'); $notPublishable.Add($d); continue }
                        if (-not $Force -and -not [string]::IsNullOrWhiteSpace($latest) -and $lastRefused[$d] -eq $latest) {
                            $reasons.Add(('{0}: not supported' -f $d)); continue
                        }
                        if (-not $Force -and -not [string]::IsNullOrWhiteSpace($latest) -and $lastPublished[$d] -eq $latest) {
                            $reasons.Add(('{0}: {1} already published' -f $d, $latest)); continue
                        }
                        $publishTo.Add($d)
                    }
                    if ($publishTo.Count -gt 0) { $steps += ('Publish to ' + ($publishTo -join '+')) }
                    elseif ($reasons.Count -gt 0 -and [string]::IsNullOrWhiteSpace($latest)) {
                        # The version is unknown until the check runs; publish
                        # decisions wait for it.
                        $steps += 'Publish after check'
                    }
                }
                $planned = ($steps -join ', ')
            }
        }

        $rows.Add([pscustomobject]@{
            Packager      = $name
            Application   = [string]$app.Application
            Vendor        = [string]$app.Vendor
            Version       = $latest
            Current       = $current
            ConfigMgr     = [bool]$destinations.ConfigMgr
            WSUS          = [bool]$destinations.WSUS
            Intune        = [bool]$destinations.Intune
            Source        = [string]$destinations.Source
            LastConfigMgr = $lastPublished['ConfigMgr']
            LastWSUS      = $lastPublished['WSUS']
            LastIntune    = $lastPublished['Intune']
            PublishTo     = $publishTo.ToArray()
            NotPublishable = $notPublishable.ToArray()
            Planned       = $planned
            Reason        = ($reasons -join '; ')
            Include       = $include
        })
    }
    return $rows.ToArray()
}

function Select-OneClickRunDestinations {
    <#
    .SYNOPSIS
        The destinations a run publishes to: the selected ones without those
        the plan reported as not ready or not supported.
    #>
    param([AllowEmptyCollection()][string[]]$Selected = @(), [AllowEmptyCollection()][string[]]$NotPublishable = @())
    return @($Selected | Where-Object { $NotPublishable -notcontains $_ })
}

function Get-OneClickPlanSummary {
    <#
    .SYNOPSIS
        The count line above the plan grid, and the prerequisites that block
        a destination for every row.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Plan,
        [Parameter(Mandatory)]$Prefs,
        [ValidateSet('Report', 'Stage', 'StageAndPackage')][string]$Action = 'Report'
    )
    $included = @($Plan | Where-Object { $_.Include })
    $counts = [ordered]@{}
    foreach ($d in $script:OneClickDestinations) {
        $counts[$d] = @($included | Where-Object { $_.PublishTo -contains $d }).Count
    }
    $line = '{0} application(s)' -f $included.Count
    if ($Plan.Count -ne $included.Count) { $line += (', {0} skipped' -f ($Plan.Count - $included.Count)) }
    if ($Action -eq 'StageAndPackage') {
        $line += ': ' + (($script:OneClickDestinations | ForEach-Object { '{0} to {1}' -f $counts[$_], $_ }) -join ', ')
    }
    else {
        $line += $(if ($Action -eq 'Stage') { ': check and stage, no publish' } else { ': check only, no publish' })
    }

    $blocking = New-Object System.Collections.Generic.List[string]
    if ($Action -eq 'StageAndPackage') {
        foreach ($d in $script:OneClickDestinations) {
            $wanted = @($Plan | Where-Object { [bool]$_.$d }).Count
            if ($wanted -eq 0) { continue }
            $state = Test-OneClickDestinationInUse -Prefs $Prefs -Destination $d
            if (-not $state.Ready) { $blocking.Add(('{0}: {1} ({2} application(s) select it)' -f $d, $state.Reason, $wanted)) }
        }
    }
    return [pscustomobject]@{
        Line     = $line
        Counts   = [pscustomobject]$counts
        Included = $included.Count
        Blocking = $blocking.ToArray()
    }
}

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

function Get-OneClickReportFolder {
    param([Parameter(Mandatory)][string]$LogFolder)
    return (Join-Path $LogFolder 'OneClick')
}

function Write-OneClickReport {
    <#
    .SYNOPSIS
        Writes the run report: one Markdown file and one JSON file per run.
    .PARAMETER Run
        Object with Started, Ended (datetime), Operator, Computer, Action,
        Force (bool), Scope (object with ConfigMgr, WSUS, Intune text),
        Canceled (bool).
    .PARAMETER Rows
        Plan rows extended by the run: Outcome (Published, Staged, Checked,
        Skipped, Not supported, Failed), Result per destination in
        ResultConfigMgr / ResultWSUS / ResultIntune (text with the identifier),
        IdConfigMgr / IdWSUS / IdIntune, DetailConfigMgr / DetailWSUS /
        DetailIntune (the full message behind a short result), and Reason.
    .OUTPUTS
        [pscustomobject] MarkdownPath, JsonPath.
    #>
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Folder
    )
    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $stamp = ([datetime]$Run.Started).ToString('yyyyMMdd-HHmmss')
    # Two runs can start in the same second.
    $name = 'one-click-{0}' -f $stamp
    $suffix = 1
    while ((Test-Path -LiteralPath (Join-Path $Folder ($name + '.md'))) -or (Test-Path -LiteralPath (Join-Path $Folder ($name + '.json')))) {
        $suffix++
        $name = 'one-click-{0}-{1}' -f $stamp, $suffix
    }
    $mdPath = Join-Path $Folder ($name + '.md')
    $jsonPath = Join-Path $Folder ($name + '.json')

    $get = { param($obj, $prop) if ($obj.PSObject.Properties[$prop]) { [string]$obj.PSObject.Properties[$prop].Value } else { '' } }
    # Application names and messages come from vendor pages and servers; a
    # Markdown viewer must show them as text, not as links or HTML.
    $esc = {
        param([string]$s)
        $s = (($s -replace '&', '&amp;') -replace '<', '&lt;') -replace '>', '&gt;'
        ($s -replace '([\\`*_\[\]|])', '\$1') -replace "`r?`n", ' '
    }

    $published = @($Rows | Where-Object { (& $get $_ 'Outcome') -eq 'Published' }).Count
    $failed = @($Rows | Where-Object { (& $get $_ 'Outcome') -eq 'Failed' }).Count
    $notSupported = @($Rows | Where-Object { (& $get $_ 'Outcome') -eq 'Not supported' }).Count
    $skipped = @($Rows | Where-Object { (& $get $_ 'Outcome') -eq 'Skipped' }).Count

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# One Click run ' + ([datetime]$Run.Started).ToString('yyyy-MM-dd HH:mm'))
    $lines.Add('')
    $lines.Add('| Field | Value |')
    $lines.Add('|---|---|')
    $lines.Add('| Started | ' + ([datetime]$Run.Started).ToString('yyyy-MM-dd HH:mm:ss') + ' |')
    $lines.Add('| Ended | ' + ([datetime]$Run.Ended).ToString('yyyy-MM-dd HH:mm:ss') + ' |')
    $lines.Add('| Operator | ' + (& $esc (& $get $Run 'Operator')) + ' |')
    $lines.Add('| Computer | ' + (& $esc (& $get $Run 'Computer')) + ' |')
    $lines.Add('| Action | ' + (& $get $Run 'Action') + $(if ($Run.PSObject.Properties['Force'] -and $Run.Force) { ' (force)' } else { '' }) + ' |')
    if ($Run.PSObject.Properties['Canceled'] -and $Run.Canceled) { $lines.Add('| Canceled | yes |') }
    $lines.Add('| Applications | ' + $Rows.Count + ' |')
    $lines.Add('| Published | ' + $published + ' |')
    $lines.Add('| Not supported | ' + $notSupported + ' |')
    $lines.Add('| Skipped | ' + $skipped + ' |')
    $lines.Add('| Failed | ' + $failed + ' |')
    $lines.Add('')
    if ($Run.PSObject.Properties['Scope'] -and $Run.Scope) {
        $lines.Add('## Scope')
        $lines.Add('')
        foreach ($d in $script:OneClickDestinations) {
            $text = & $get $Run.Scope $d
            if ($text) { $lines.Add(('- {0}: {1}' -f $d, $text)) }
        }
        $lines.Add('')
    }
    $lines.Add('## Applications')
    $lines.Add('')
    $lines.Add('| Application | Version | Outcome | ConfigMgr | WSUS | Intune | Reason |')
    $lines.Add('|---|---|---|---|---|---|---|')
    foreach ($row in $Rows) {
        $cells = @(
            (& $esc (& $get $row 'Application')),
            (& $esc (& $get $row 'Version')),
            (& $esc (& $get $row 'Outcome')),
            (& $esc (& $get $row 'ResultConfigMgr')),
            (& $esc (& $get $row 'ResultWSUS')),
            (& $esc (& $get $row 'ResultIntune')),
            (& $esc (& $get $row 'Reason'))
        )
        $lines.Add('| ' + ($cells -join ' | ') + ' |')
    }
    $lines.Add('')
    $details = @(foreach ($row in $Rows) {
        foreach ($d in $script:OneClickDestinations) {
            $text = & $get $row ('Detail' + $d)
            if ($text) { '- {0}, {1}: {2}' -f (& $esc (& $get $row 'Application')), $d, (& $esc $text) }
        }
    })
    if ($details.Count) {
        $lines.Add('## Details')
        $lines.Add('')
        foreach ($line in $details) { $lines.Add($line) }
        $lines.Add('')
    }
    $lines.Add('## Rollback')
    $lines.Add('')
    $rollback = [ordered]@{
        ConfigMgr = 'retire the application named in the ConfigMgr column.'
        WSUS      = 'expire the package ID in the WSUS column from WSUS Updates.'
        Intune    = 'unassign the app ID in the Intune column in the Intune admin center.'
    }
    $hasScope = [bool]($Run.PSObject.Properties['Scope'] -and $Run.Scope)
    foreach ($d in $script:OneClickDestinations) {
        if (-not $hasScope -or (& $get $Run.Scope $d)) { $lines.Add(('- {0}: {1}' -f $d, $rollback[$d])) }
    }
    $lines.Add('')
    [IO.File]::WriteAllText($mdPath, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))

    $json = [pscustomobject]@{
        Started      = ([datetime]$Run.Started).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Ended        = ([datetime]$Run.Ended).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Operator     = (& $get $Run 'Operator')
        Computer     = (& $get $Run 'Computer')
        Action       = (& $get $Run 'Action')
        Force        = [bool]($Run.PSObject.Properties['Force'] -and $Run.Force)
        Canceled     = [bool]($Run.PSObject.Properties['Canceled'] -and $Run.Canceled)
        Scope        = $(if ($Run.PSObject.Properties['Scope']) { $Run.Scope } else { $null })
        Applications = @($Rows | ForEach-Object {
            [pscustomobject]@{
                Packager        = (& $get $_ 'Packager')
                Application     = (& $get $_ 'Application')
                Version         = (& $get $_ 'Version')
                Outcome         = (& $get $_ 'Outcome')
                Reason          = (& $get $_ 'Reason')
                ConfigMgr       = [pscustomobject]@{ Result = (& $get $_ 'ResultConfigMgr'); Id = (& $get $_ 'IdConfigMgr'); Detail = (& $get $_ 'DetailConfigMgr') }
                WSUS            = [pscustomobject]@{ Result = (& $get $_ 'ResultWSUS'); Id = (& $get $_ 'IdWSUS'); Detail = (& $get $_ 'DetailWSUS') }
                Intune          = [pscustomobject]@{ Result = (& $get $_ 'ResultIntune'); Id = (& $get $_ 'IdIntune'); Detail = (& $get $_ 'DetailIntune') }
            }
        })
    }
    [IO.File]::WriteAllText($jsonPath, ($json | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
    return [pscustomobject]@{ MarkdownPath = $mdPath; JsonPath = $jsonPath }
}

function Get-OneClickReportList {
    <#
    .SYNOPSIS
        Earlier run reports, newest first.
    .OUTPUTS
        [pscustomobject] Started, Action, Applications, Published, Failed,
        MarkdownPath, JsonPath.
    #>
    param([Parameter(Mandatory)][string]$Folder)
    if (-not (Test-Path -LiteralPath $Folder)) { return @() }
    $items = foreach ($file in (Get-ChildItem -LiteralPath $Folder -Filter 'one-click-*.json' -File | Sort-Object Name -Descending)) {
        try {
            $data = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $apps = @($data.Applications)
            [pscustomobject]@{
                Started      = [datetime]::Parse([string]$data.Started, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime()
                Action       = [string]$data.Action
                Applications = $apps.Count
                Published    = @($apps | Where-Object { $_.Outcome -eq 'Published' }).Count
                Failed       = @($apps | Where-Object { $_.Outcome -eq 'Failed' }).Count
                MarkdownPath = [IO.Path]::ChangeExtension($file.FullName, '.md')
                JsonPath     = $file.FullName
            }
        }
        catch { }
    }
    return @($items)
}

Export-ModuleMember -Function @(
    'Get-OneClickDestinationNames', 'ConvertTo-OneClickDestinationSet', 'Get-OneClickDestinations',
    'Test-OneClickDestinationInUse', 'Get-OneClickDestinationScope',
    'Get-OneClickPublishedVersion', 'Set-OneClickPublishedVersion',
    'Get-OneClickNotSupportedVersion', 'Get-OneClickNotSupportedReason', 'Set-OneClickNotSupported',
    'Get-OneClickCadenceDays', 'Get-OneClickPlan', 'Get-OneClickPlanSummary', 'Select-OneClickRunDestinations',
    'Get-OneClickReportFolder', 'Write-OneClickReport', 'Get-OneClickReportList'
)
