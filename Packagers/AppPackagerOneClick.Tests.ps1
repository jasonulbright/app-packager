BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'AppPackagerOneClick.psd1') -Force

    function New-TestPrefs {
        param(
            [string]$Target = 'MECM',
            [string[]]$Tracked = @('package-a', 'package-b'),
            [hashtable]$Destinations = $null,
            [string[]]$InUse = @('ConfigMgr', 'Intune', 'Wsus'),
            [string]$WsusServer = 'wsus01',
            [string]$ApprovalGroup = 'Pilot',
            [hashtable]$CadenceOverrides = @{}
        )
        $dest = $null
        if ($Destinations) { $dest = [pscustomobject]$Destinations }
        [pscustomobject]@{
            SiteCode      = 'MCM'
            FileShareRoot = '\\server\share'
            Systems       = [pscustomobject]@{ ConfigMgr = ('ConfigMgr' -in $InUse); Intune = ('Intune' -in $InUse); Wsus = ('Wsus' -in $InUse) }
            DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $true } }
            Intune        = [pscustomobject]@{ DeploymentTarget = $Target; TenantId = 't'; ClientId = 'c'; ClientSecretProtected = 'p' }
            Wsus          = [pscustomobject]@{ ServerName = $WsusServer; ApprovalGroup = $ApprovalGroup }
            ContentDistribution = [pscustomobject]@{ AutoDistribute = $true; DPGroupName = 'All DPs'; DeployToTestCollection = $true; TestCollectionName = 'Pilot Devices' }
            AppFlow       = [pscustomobject]@{ Tracked = $Tracked; CadenceOverrides = [pscustomobject]$CadenceOverrides; Destinations = $dest }
        }
    }

    function New-TestApp {
        param([string]$Packager, [string]$Application = $Packager, [string]$Latest = '', [int]$CadenceDays = 0, [bool]$WsusUnsupported = $false)
        [pscustomobject]@{ Packager = $Packager; Application = $Application; Vendor = 'Vendor'; LatestVersion = $Latest; CadenceDays = $CadenceDays; WsusUnsupported = $WsusUnsupported }
    }
}

Describe 'ConvertTo-OneClickDestinationSet' {
    It 'maps every deployment target to its destinations' {
        (ConvertTo-OneClickDestinationSet 'MECM').ConfigMgr | Should -BeTrue
        (ConvertTo-OneClickDestinationSet 'MECM').WSUS | Should -BeFalse
        (ConvertTo-OneClickDestinationSet 'MECMAndWSUS').WSUS | Should -BeTrue
        (ConvertTo-OneClickDestinationSet 'WSUSOnly').ConfigMgr | Should -BeFalse
        (ConvertTo-OneClickDestinationSet 'IntuneOnly').Intune | Should -BeTrue
        (ConvertTo-OneClickDestinationSet 'MECMAndIntune').Intune | Should -BeTrue
    }
    It 'reads flags from an object, a hashtable with the Wsus alias, and a name list' {
        $fromObject = ConvertTo-OneClickDestinationSet ([pscustomobject]@{ ConfigMgr = $false; WSUS = $true; Intune = $false })
        $fromObject.WSUS | Should -BeTrue
        $fromObject.ConfigMgr | Should -BeFalse
        (ConvertTo-OneClickDestinationSet @{ Wsus = $true }).WSUS | Should -BeTrue
        $fromList = ConvertTo-OneClickDestinationSet @('MECM', 'Intune')
        $fromList.ConfigMgr | Should -BeTrue
        $fromList.Intune | Should -BeTrue
        $fromList.WSUS | Should -BeFalse
    }
    It 'treats nothing as no destination' {
        $none = ConvertTo-OneClickDestinationSet $null
        ($none.ConfigMgr -or $none.WSUS -or $none.Intune) | Should -BeFalse
    }
}

Describe 'Get-OneClickDestinations' {
    It 'uses the saved One Click destination when the row has no selection' {
        $set = Get-OneClickDestinations -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -PackagerName 'package-a'
        $set.ConfigMgr | Should -BeTrue
        $set.WSUS | Should -BeTrue
        $set.Source | Should -Be 'Default'
    }
    It 'prefers the saved default destinations over the old target' {
        $prefs = New-TestPrefs -Target 'MECM'
        $prefs.AppFlow | Add-Member -NotePropertyName DefaultDestinations -NotePropertyValue ([pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $true })
        $set = Get-OneClickDestinations -Prefs $prefs -PackagerName 'package-a'
        ($set.ConfigMgr -and $set.WSUS -and $set.Intune) | Should -BeTrue
    }
    It 'lets the row selection win over the default' {
        $prefs = New-TestPrefs -Target 'MECM' -Destinations @{ 'package-a' = [pscustomobject]@{ ConfigMgr = $false; WSUS = $true; Intune = $false } }
        $set = Get-OneClickDestinations -Prefs $prefs -PackagerName 'package-a'
        $set.ConfigMgr | Should -BeFalse
        $set.WSUS | Should -BeTrue
        $set.Source | Should -Be 'Row'
        (Get-OneClickDestinations -Prefs $prefs -PackagerName 'package-b').Source | Should -Be 'Default'
    }
}

Describe 'Test-OneClickDestinationInUse' {
    It 'is ready when the system is in use and configured' {
        (Test-OneClickDestinationInUse -Prefs (New-TestPrefs) -Destination 'ConfigMgr').Ready | Should -BeTrue
        (Test-OneClickDestinationInUse -Prefs (New-TestPrefs) -Destination 'WSUS').Ready | Should -BeTrue
        (Test-OneClickDestinationInUse -Prefs (New-TestPrefs) -Destination 'Intune').Ready | Should -BeTrue
    }
    It 'names the system that is not in use' {
        $r = Test-OneClickDestinationInUse -Prefs (New-TestPrefs -InUse @('ConfigMgr')) -Destination 'WSUS'
        $r.Ready | Should -BeFalse
        $r.Reason | Should -Be 'WSUS is not in use'
    }
    It 'names the missing setting' {
        (Test-OneClickDestinationInUse -Prefs (New-TestPrefs -WsusServer '') -Destination 'WSUS').Reason | Should -Be 'WSUS server not set'
        $prefs = New-TestPrefs; $prefs.SiteCode = ''
        (Test-OneClickDestinationInUse -Prefs $prefs -Destination 'ConfigMgr').Reason | Should -Be 'site code not set'
    }
}

Describe 'Get-OneClickDestinationScope' {
    It 'describes where each destination lands' {
        Get-OneClickDestinationScope -Prefs (New-TestPrefs) -Destination 'ConfigMgr' | Should -Be 'distribute to DP group All DPs; deploy to collection Pilot Devices'
        Get-OneClickDestinationScope -Prefs (New-TestPrefs) -Destination 'WSUS' | Should -Be 'approved for group Pilot'
        Get-OneClickDestinationScope -Prefs (New-TestPrefs -ApprovalGroup '') -Destination 'WSUS' | Should -Be 'published, not approved'
        Get-OneClickDestinationScope -Prefs (New-TestPrefs) -Destination 'Intune' | Should -Be 'app created, not assigned'
    }
}

Describe 'publish history per destination' {
    It 'records and reads back a publish, converting JSON objects to hashtables' {
        $history = @{ 'package-a' = [pscustomobject]@{ LastChecked = '2026-01-01T00:00:00Z'; LastKnownVersion = '1.0' } }
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '1.0' -Id 'guid-1' -At ([datetime]'2026-02-03T04:05:06Z') | Out-Null
        $history['package-a'] | Should -BeOfType [hashtable]
        $history['package-a']['LastKnownVersion'] | Should -Be '1.0'
        $rec = Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'WSUS'
        $rec.Version | Should -Be '1.0'
        $rec.Id | Should -Be 'guid-1'
        $rec.At | Should -Be '2026-02-03T04:05:06Z'
        Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'Intune' | Should -BeNullOrEmpty
    }
    It 'creates the entry for an application without history' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-new' -Destination 'ConfigMgr' -Version '2.0' | Out-Null
        (Get-OneClickPublishedVersion -HistoryEntry $history['package-new'] -Destination 'ConfigMgr').Version | Should -Be '2.0'
    }
    It 'survives a JSON round trip' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'Intune' -Version '3.0' -Id 'app-id' | Out-Null
        $back = ($history | ConvertTo-Json -Depth 6) | ConvertFrom-Json
        (Get-OneClickPublishedVersion -HistoryEntry $back.'package-a' -Destination 'Intune').Id | Should -Be 'app-id'
    }
}

Describe 'Get-OneClickPlan' {
    It 'lists only tracked applications' {
        $plan = Get-OneClickPlan -Apps @((New-TestApp 'package-a'), (New-TestApp 'package-x')) -Prefs (New-TestPrefs)
        @($plan).Count | Should -Be 1
        $plan[0].Packager | Should -Be 'package-a'
    }
    It 'skips a row without a destination and says so' {
        $prefs = New-TestPrefs -Destinations @{ 'package-a' = [pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $false } }
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs $prefs -Action 'StageAndPackage')[0]
        $row.Include | Should -BeFalse
        $row.Planned | Should -Be 'Skip'
        $row.Reason | Should -Be 'no destination selected'
    }
    It 'plans check, stage and the publish destinations' {
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -Action 'StageAndPackage')[0]
        $row.Planned | Should -Be 'Check, Stage, Publish to ConfigMgr+WSUS'
        $row.PublishTo | Should -Be @('ConfigMgr', 'WSUS')
        $row.Include | Should -BeTrue
    }
    It 'drops a destination whose version is already published unless forced' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' | Out-Null
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $history -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -Be @('ConfigMgr')
        $row.Reason | Should -Be 'WSUS: 2.0 already published'
        $row.LastWSUS | Should -Be '2.0'
        $forced = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $history -Action 'StageAndPackage' -Force)[0]
        $forced.PublishTo | Should -Be @('ConfigMgr', 'WSUS')
    }
    It 'publishes a destination newly selected since the last publish' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '2.0' | Out-Null
        $prefs = New-TestPrefs -Destinations @{ 'package-a' = [pscustomobject]@{ ConfigMgr = $true; WSUS = $true; Intune = $false } }
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs $prefs -History $history -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -Be @('WSUS')
        $row.Planned | Should -Be 'Check, Stage, Publish to WSUS'
    }
    It 'drops a destination that is not in use and names it' {
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '1.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS' -InUse @('ConfigMgr')) -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -Be @('ConfigMgr')
        $row.Reason | Should -Be 'WSUS: WSUS is not in use'
    }
    It 'drops WSUS for a packager WSUS cannot carry' {
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '1.0' -WsusUnsupported $true) -Prefs (New-TestPrefs -Target 'WSUSOnly') -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -BeNullOrEmpty
        $row.Reason | Should -Be 'WSUS: not supported'
        $row.Planned | Should -Be 'Check, Stage'
    }
    It 'drops a destination that refused the same version unless forced, and tries a newer version' {
        $history = @{ 'package-a' = [pscustomobject]@{ LastKnownVersion = '2.0' } }
        Set-OneClickNotSupported -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' -Reason 'not supported: DetectionNotMappable' | Out-Null
        $back = ($history | ConvertTo-Json -Depth 6) | ConvertFrom-Json
        $entries = @{ 'package-a' = $back.'package-a' }
        Get-OneClickNotSupportedVersion -HistoryEntry $entries['package-a'] -Destination 'WSUS' | Should -Be '2.0'
        Get-OneClickNotSupportedReason -HistoryEntry $entries['package-a'] -Destination 'WSUS' | Should -Be 'not supported: DetectionNotMappable'
        Get-OneClickNotSupportedReason -HistoryEntry $entries['package-a'] -Destination 'ConfigMgr' | Should -Be ''
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $entries -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -Be @('ConfigMgr')
        $row.Reason | Should -Be 'WSUS: not supported'
        $forced = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $entries -Action 'StageAndPackage' -Force)[0]
        $forced.PublishTo | Should -Be @('ConfigMgr', 'WSUS')
        $newer = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.1') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $entries -Action 'StageAndPackage')[0]
        $newer.PublishTo | Should -Be @('ConfigMgr', 'WSUS')
    }
    It 'clears a refusal when the destination later publishes' {
        $history = @{}
        Set-OneClickNotSupported -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' | Out-Null
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' | Out-Null
        Get-OneClickNotSupportedVersion -HistoryEntry $history['package-a'] -Destination 'WSUS' | Should -Be ''
    }
    It 'waits for the check when the version is unknown and a destination is already published' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '1.0' | Out-Null
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs -Target 'MECM') -History @{} -Action 'StageAndPackage')[0]
        $row.Planned | Should -Be 'Check, Stage, Publish to ConfigMgr'
        $row.Version | Should -Be ''
    }
    It 'uses the last known version from history when the row has none' {
        $history = @{ 'package-a' = [pscustomobject]@{ LastKnownVersion = '5.5'; LastChecked = '2020-01-01T00:00:00Z' } }
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs) -History $history -Action 'Stage')[0]
        $row.Version | Should -Be '5.5'
        $row.Planned | Should -Be 'Check, Stage'
    }
    It 'skips a Report row inside its cadence and runs it when forced or overdue' {
        $now = [datetime]'2026-03-10T12:00:00Z'
        $history = @{ 'package-a' = [pscustomobject]@{ LastChecked = '2026-03-08T12:00:00Z' } }
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs) -History $history -Action 'Report' -Now $now)[0]
        $row.Include | Should -BeFalse
        $row.Reason | Should -Be 'checked within the 7-day cadence'
        (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs) -History $history -Action 'Report' -Now $now -Force)[0].Include | Should -BeTrue
        (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -CadenceDays 1) -Prefs (New-TestPrefs) -History $history -Action 'Report' -Now $now)[0].Include | Should -BeTrue
        $override = New-TestPrefs -CadenceOverrides @{ 'package-a' = 1 }
        (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -CadenceDays 30) -Prefs $override -History $history -Action 'Report' -Now $now)[0].Include | Should -BeTrue
    }
    It 'does not apply the cadence to Stage runs' {
        $history = @{ 'package-a' = [pscustomobject]@{ LastChecked = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } }
        (Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs) -History $history -Action 'Stage')[0].Include | Should -BeTrue
    }
    It 'returns an empty plan for no tracked applications' {
        @(Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs -Tracked @())).Count | Should -Be 0
    }
}

Describe 'Get-OneClickPlanSummary' {
    It 'counts the publishes per destination and the skipped rows' {
        $prefs = New-TestPrefs -Target 'MECMAndWSUS' -Tracked @('package-a', 'package-b', 'package-c') -Destinations @{ 'package-c' = [pscustomobject]@{ ConfigMgr = $false; WSUS = $false; Intune = $false } }
        $plan = Get-OneClickPlan -Apps @((New-TestApp 'package-a' -Latest '1'), (New-TestApp 'package-b' -Latest '1'), (New-TestApp 'package-c' -Latest '1')) -Prefs $prefs -Action 'StageAndPackage'
        $summary = Get-OneClickPlanSummary -Plan $plan -Prefs $prefs -Action 'StageAndPackage'
        $summary.Line | Should -Be '2 application(s), 1 skipped: 2 to ConfigMgr, 2 to WSUS, 0 to Intune'
        $summary.Counts.WSUS | Should -Be 2
        $summary.Blocking | Should -BeNullOrEmpty
    }
    It 'names a destination that rows select but that is not ready' {
        $prefs = New-TestPrefs -Target 'MECMAndWSUS' -WsusServer ''
        $plan = Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '1') -Prefs $prefs -Action 'StageAndPackage'
        $summary = Get-OneClickPlanSummary -Plan $plan -Prefs $prefs -Action 'StageAndPackage'
        $summary.Blocking | Should -Be @('WSUS: WSUS server not set (1 application(s) select it)')
    }
    It 'says check only for a Report run' {
        $plan = Get-OneClickPlan -Apps @(New-TestApp 'package-a') -Prefs (New-TestPrefs) -Action 'Report'
        (Get-OneClickPlanSummary -Plan $plan -Prefs (New-TestPrefs) -Action 'Report').Line | Should -Be '1 application(s): check only, no publish'
    }
}

Describe 'Write-OneClickReport and Get-OneClickReportList' {
    It 'writes the Markdown and JSON report and lists it' {
        $folder = Join-Path $TestDrive 'OneClick'
        $run = [pscustomobject]@{
            Started = [datetime]'2026-04-01T10:00:00'; Ended = [datetime]'2026-04-01T10:12:30'
            Operator = 'CONTOSO\jason'; Computer = 'WS01'; Action = 'StageAndPackage'; Force = $false; Canceled = $false
            Scope = [pscustomobject]@{ ConfigMgr = 'deploy to collection Pilot Devices'; WSUS = 'approved for group Pilot'; Intune = 'app created, not assigned' }
        }
        $rows = @(
            [pscustomobject]@{ Packager = 'package-a'; Application = 'App A'; Version = '2.0'; Outcome = 'Published'; ResultConfigMgr = 'published App A 2.0'; IdConfigMgr = 'App A 2.0'; ResultWSUS = 'published 1111-2222'; IdWSUS = '1111-2222'; ResultIntune = ''; IdIntune = ''; Reason = '' }
            [pscustomobject]@{ Packager = 'package-b'; Application = 'App | B'; Version = '1.0'; Outcome = 'Not supported'; ResultConfigMgr = ''; IdConfigMgr = ''; ResultWSUS = 'not supported: CustomInstall'; IdWSUS = ''; ResultIntune = ''; IdIntune = ''; Reason = 'WSUS: not supported' }
            [pscustomobject]@{ Packager = 'package-c'; Application = 'App C'; Version = ''; Outcome = 'Failed'; ResultConfigMgr = 'failed: exit code 1'; IdConfigMgr = ''; ResultWSUS = ''; IdWSUS = ''; ResultIntune = ''; IdIntune = ''; Reason = 'stage error' }
        )
        $paths = Write-OneClickReport -Run $run -Rows $rows -Folder $folder
        Test-Path $paths.MarkdownPath | Should -BeTrue
        $md = Get-Content $paths.MarkdownPath -Raw
        $md | Should -Match '\| Published \| 1 \|'
        $md | Should -Match '\| Not supported \| 1 \|'
        $md | Should -Match '\| Failed \| 1 \|'
        $md | Should -Match '- WSUS: approved for group Pilot'
        $md | Should -Match '\| App A \| 2\.0 \| Published \| published App A 2\.0 \| published 1111-2222 \|  \|  \|'
        $md | Should -Match 'App \\\| B'
        $md | Should -Match '## Rollback'
        $json = Get-Content $paths.JsonPath -Raw | ConvertFrom-Json
        $json.Applications.Count | Should -Be 3
        $json.Applications[0].WSUS.Id | Should -Be '1111-2222'
        $json.Operator | Should -Be 'CONTOSO\jason'

        $list = Get-OneClickReportList -Folder $folder
        @($list).Count | Should -Be 1
        $list[0].Published | Should -Be 1
        $list[0].Failed | Should -Be 1
        $list[0].Applications | Should -Be 3
        $list[0].MarkdownPath | Should -Be $paths.MarkdownPath
    }
    It 'carries the full refusal text that the result cell shortens to its code' {
        $run = [pscustomobject]@{ Started = [datetime]'2026-04-04T10:00:00'; Ended = [datetime]'2026-04-04T10:01:00'; Operator = 'o'; Computer = 'c'; Action = 'Stage and Publish' }
        $full = 'not supported: DetectionNotMappable: The detection only checks that a file or registry key exists.'
        $rows = @(
            [pscustomobject]@{ Packager = 'package-b'; Application = 'App B'; Version = '1.0'; Outcome = 'Not supported'; ResultWSUS = 'not supported: DetectionNotMappable'; IdWSUS = ''; DetailWSUS = $full; Reason = 'WSUS: not supported' }
            [pscustomobject]@{ Packager = 'package-a'; Application = 'App A'; Version = '2.0'; Outcome = 'Published'; ResultWSUS = 'published 1111-2222'; IdWSUS = '1111-2222'; DetailWSUS = ''; Reason = '' }
        )
        $paths = Write-OneClickReport -Run $run -Rows $rows -Folder (Join-Path $TestDrive 'r4')
        $md = Get-Content $paths.MarkdownPath -Raw
        $md | Should -Match '\| App B \| 1\.0 \| Not supported \|  \| not supported: DetectionNotMappable \|'
        $md | Should -Match ([regex]::Escape('## Details'))
        $md | Should -Match ([regex]::Escape('- App B, WSUS: ' + $full))
        $md | Should -Not -Match 'App A, WSUS'
        $json = Get-Content $paths.JsonPath -Raw | ConvertFrom-Json
        $json.Applications[0].WSUS.Detail | Should -Be $full
        $json.Applications[1].WSUS.Detail | Should -Be ''
    }
    It 'writes no Details section when no result has a longer message' {
        $run = [pscustomobject]@{ Started = [datetime]'2026-04-05T10:00:00'; Ended = [datetime]'2026-04-05T10:01:00'; Operator = 'o'; Computer = 'c'; Action = 'Stage and Publish' }
        $rows = @([pscustomobject]@{ Packager = 'package-a'; Application = 'App A'; Version = '2.0'; Outcome = 'Published'; ResultWSUS = 'published 1111-2222'; IdWSUS = '1111-2222'; Reason = '' })
        $paths = Write-OneClickReport -Run $run -Rows $rows -Folder (Join-Path $TestDrive 'r5')
        Get-Content $paths.MarkdownPath -Raw | Should -Not -Match '## Details'
    }
    It 'lists rollback steps only for the destinations in the scope' {
        $run = [pscustomobject]@{
            Started = [datetime]'2026-04-03T10:00:00'; Ended = [datetime]'2026-04-03T10:01:00'; Operator = 'o'; Computer = 'c'; Action = 'Stage and Publish'
            Scope = [pscustomobject]@{ ConfigMgr = 'application only, no distribution, no deployment'; WSUS = 'published, not approved' }
        }
        $paths = Write-OneClickReport -Run $run -Rows @() -Folder (Join-Path $TestDrive 'r3')
        $md = Get-Content $paths.MarkdownPath -Raw
        $md | Should -Match '- ConfigMgr: retire the application'
        $md | Should -Match '- WSUS: expire the package ID'
        $md | Should -Not -Match '- Intune: unassign'
    }
    It 'returns an empty list for a missing folder' {
        @(Get-OneClickReportList -Folder (Join-Path $TestDrive 'none')).Count | Should -Be 0
    }
    It 'names the file by the start time' {
        $run = [pscustomobject]@{ Started = [datetime]'2026-04-02T01:02:03'; Ended = [datetime]'2026-04-02T01:02:04'; Operator = 'o'; Computer = 'c'; Action = 'Report' }
        $paths = Write-OneClickReport -Run $run -Rows @() -Folder (Join-Path $TestDrive 'r2')
        Split-Path $paths.MarkdownPath -Leaf | Should -Be 'one-click-20260402-010203.md'
    }
}
