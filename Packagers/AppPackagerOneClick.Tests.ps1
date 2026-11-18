#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
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
    It 'keys each destination by its configured instance' {
        $prefs = New-TestPrefs
        Get-OneClickInstanceKey -Prefs $prefs -Destination 'ConfigMgr' | Should -Be 'mcm@'
        Get-OneClickInstanceKey -Prefs $prefs -Destination 'WSUS' | Should -Be 'wsus01:'
        Get-OneClickInstanceKey -Prefs $prefs -Destination 'Intune' | Should -Be 't'
        $prefs | Add-Member -NotePropertyName ProviderMachineName -NotePropertyValue 'CM01.contoso.com'
        $prefs.Wsus | Add-Member -NotePropertyName PortNumber -NotePropertyValue 8531
        Get-OneClickInstanceKey -Prefs $prefs -Destination 'ConfigMgr' | Should -Be 'mcm@cm01.contoso.com'
        Get-OneClickInstanceKey -Prefs $prefs -Destination 'WSUS' | Should -Be 'wsus01:8531'
        $keys = Get-OneClickInstanceKeys -Prefs $prefs
        $keys.Keys.Count | Should -Be 3
        $keys['Intune'] | Should -Be 't'
        Get-OneClickInstanceKey -Prefs (New-TestPrefs -WsusServer '') -Destination 'WSUS' | Should -Be ''
    }
    It 'ignores a receipt and a refusal written for another instance' {
        $history = @{}
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'Intune' -Version '3.0' -Id 'app-id' -Instance 'tenant-a' -Build 'p1|2|abc' | Out-Null
        Set-OneClickNotSupported -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '3.0' -Reason 'not supported: MSIX' -Instance 'wsus01:8530' | Out-Null
        $entry = $history['package-a']
        $rec = Get-OneClickPublishedVersion -HistoryEntry $entry -Destination 'Intune' -Instance 'tenant-a'
        $rec.Version | Should -Be '3.0'
        $rec.Instance | Should -Be 'tenant-a'
        $rec.Build | Should -Be 'p1|2|abc'
        Get-OneClickPublishedVersion -HistoryEntry $entry -Destination 'Intune' -Instance 'tenant-b' | Should -BeNullOrEmpty
        Get-OneClickNotSupportedVersion -HistoryEntry $entry -Destination 'WSUS' -Instance 'wsus01:8530' | Should -Be '3.0'
        Get-OneClickNotSupportedVersion -HistoryEntry $entry -Destination 'WSUS' -Instance 'wsus02:8530' | Should -Be ''
        Get-OneClickNotSupportedReason -HistoryEntry $entry -Destination 'WSUS' -Instance 'wsus02:8530' | Should -Be ''
        # The same answers after a JSON round trip.
        $back = ($history | ConvertTo-Json -Depth 6) | ConvertFrom-Json
        (Get-OneClickPublishedVersion -HistoryEntry $back.'package-a' -Destination 'Intune' -Instance 'tenant-a').Build | Should -Be 'p1|2|abc'
        Get-OneClickPublishedVersion -HistoryEntry $back.'package-a' -Destination 'Intune' -Instance 'tenant-b' | Should -BeNullOrEmpty
    }
    It 'honors a receipt written before instances were recorded, for any instance' {
        $history = @{ 'package-a' = [pscustomobject]@{ Published = [pscustomobject]@{ Intune = [pscustomobject]@{ Version = '1.0'; At = '2026-01-01T00:00:00Z'; Id = 'x' } } } }
        (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'Intune' -Instance 'tenant-b').Version | Should -Be '1.0'
        (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'Intune').Instance | Should -Be ''
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
    It 'keeps a destination whose receipt belongs to another instance' {
        $history = @{}
        $wsus02 = Get-OneClickInstanceKey -Prefs (New-TestPrefs -WsusServer 'wsus02') -Destination 'WSUS'
        Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' -Instance $wsus02 | Out-Null
        Set-OneClickNotSupported -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '2.0' -Reason 'not supported' -Instance 'xyz@other' | Out-Null
        $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $history -Action 'StageAndPackage')[0]
        $row.PublishTo | Should -Be @('ConfigMgr', 'WSUS')
        $row.LastWSUS | Should -Be ''
        $same = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS' -WsusServer 'wsus02') -History $history -Action 'StageAndPackage')[0]
        $same.PublishTo | Should -Be @('ConfigMgr')
        $same.LastWSUS | Should -Be '2.0'
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

Describe 'Write-OneClickReport with hostile text and a repeated start time' {
    BeforeAll {
        $script:ReportRun = [pscustomobject]@{
            Started = [datetime]'2026-04-01T10:00:00'; Ended = [datetime]'2026-04-01T10:01:00'
            Operator = 'op'; Computer = 'WS01'; Action = 'Report'; Force = $false; Canceled = $false
        }
        $script:ReportRows = @([pscustomobject]@{
            Application = '[Click](http://example.invalid) <img src=x>'; Version = '1.0'; Outcome = 'Failed'
            Reason = 'a|b *bold* <script>alert(1)</script>'
        })
    }

    It 'writes names and messages as text, not as links or HTML' {
        $paths = Write-OneClickReport -Run $script:ReportRun -Rows $script:ReportRows -Folder (Join-Path $TestDrive 'esc')
        $md = Get-Content -LiteralPath $paths.MarkdownPath -Raw
        $md | Should -Not -Match '<img|<script'
        $md | Should -Not -Match '(?<!\\)\[Click\]'
        $md | Should -Match '&lt;img src=x&gt;'
        $md | Should -Match 'a\\\|b \\\*bold\\\*'
    }

    It 'keeps both reports when two runs start in the same second' {
        $folder = Join-Path $TestDrive 'same-second'
        $first = Write-OneClickReport -Run $script:ReportRun -Rows $script:ReportRows -Folder $folder
        $second = Write-OneClickReport -Run $script:ReportRun -Rows $script:ReportRows -Folder $folder
        $second.MarkdownPath | Should -Not -Be $first.MarkdownPath
        $second.JsonPath | Should -Not -Be $first.JsonPath
        @(Get-ChildItem -LiteralPath $folder -Filter *.md).Count | Should -Be 2
    }
}

Describe 'Destinations a run publishes to' {
    It 'reports a destination that is not ready and one that the packager marks as not supported' {
        $prefs = New-TestPrefs -Target 'MECMAndWSUS' -WsusServer ''
        $plan = Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '1') -Prefs $prefs -Action 'StageAndPackage'
        $plan[0].NotPublishable | Should -Contain 'WSUS'

        $app = New-TestApp 'package-b' -Latest '1'
        $app | Add-Member -NotePropertyName WsusUnsupported -NotePropertyValue $true -Force
        $ready = New-TestPrefs -Target 'MECMAndWSUS'
        $row = (Get-OneClickPlan -Apps @($app) -Prefs $ready -Action 'StageAndPackage')[0]
        if ($row.WSUS) {
            $row.NotPublishable | Should -Contain 'WSUS'
            $row.Reason | Should -Match 'WSUS: not supported'
        }
    }

    It 'leaves such a destination out of the run although its box is checked' {
        Select-OneClickRunDestinations -Selected @('ConfigMgr', 'WSUS') -NotPublishable @('WSUS') | Should -Be @('ConfigMgr')
        Select-OneClickRunDestinations -Selected @('ConfigMgr', 'WSUS') -NotPublishable @() | Should -Be @('ConfigMgr', 'WSUS')
        @(Select-OneClickRunDestinations -Selected @('WSUS') -NotPublishable @('WSUS')).Count | Should -Be 0
    }
}

Describe 'Default cadence days' {
    It 'is 7 when the preferences carry no default' {
        Get-OneClickCadenceDays -Prefs (New-TestPrefs) -PackagerName 'package-a' | Should -Be 7
    }

    It 'takes the preference default when the header and the override set none' {
        $prefs = New-TestPrefs
        $prefs.AppFlow | Add-Member -NotePropertyName DefaultCadenceDays -NotePropertyValue 30
        Get-OneClickCadenceDays -Prefs $prefs -PackagerName 'package-a' | Should -Be 30
    }

    It 'lets the packager header outrank the default' {
        $prefs = New-TestPrefs
        $prefs.AppFlow | Add-Member -NotePropertyName DefaultCadenceDays -NotePropertyValue 30
        Get-OneClickCadenceDays -Prefs $prefs -PackagerName 'package-a' -HeaderDays 14 | Should -Be 14
    }

    It 'lets the per-app override outrank the header and the default' {
        $prefs = New-TestPrefs -CadenceOverrides @{ 'package-a' = 2 }
        $prefs.AppFlow | Add-Member -NotePropertyName DefaultCadenceDays -NotePropertyValue 30
        Get-OneClickCadenceDays -Prefs $prefs -PackagerName 'package-a' -HeaderDays 14 | Should -Be 2
    }

    It 'ignores a default below 1' {
        $prefs = New-TestPrefs
        $prefs.AppFlow | Add-Member -NotePropertyName DefaultCadenceDays -NotePropertyValue 0
        Get-OneClickCadenceDays -Prefs $prefs -PackagerName 'package-a' | Should -Be 7
    }
}

Describe 'Test collection kind in the destination scope' {
    It 'names a user collection' {
        $prefs = New-TestPrefs
        $prefs.ContentDistribution | Add-Member -NotePropertyName TestCollectionIsUser -NotePropertyValue $true
        Get-OneClickDestinationScope -Prefs $prefs -Destination 'ConfigMgr' | Should -Match 'deploy to user collection '
    }

    It 'keeps the device wording when the key is absent or false' {
        Get-OneClickDestinationScope -Prefs (New-TestPrefs) -Destination 'ConfigMgr' | Should -Match 'deploy to collection '
    }
}

Describe 'Supersedence-only pass of One Click' {
    BeforeAll {
        function New-PendingHistory {
            param([string]$Version = '2.0', [string]$Instance = '')
            $history = @{}
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version $Version -Id 'App A 2.0' -Instance $Instance `
                -SupersedenceNotApplied -SupersedenceReason 'access denied' | Out-Null
            return $history
        }
    }

    Context 'the pending state of a publish' {
        It 'keeps the version as published and marks it pending with the reason' {
            $history = New-PendingHistory
            $rec = Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'ConfigMgr'
            $rec.Version | Should -Be '2.0'
            $rec.Id | Should -Be 'App A 2.0'
            $rec.SupersedencePending | Should -BeTrue
            $rec.SupersedenceReason | Should -Be 'access denied'
        }
        It 'is not pending after a publish whose step succeeded' {
            $history = @{}
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '2.0' -Id 'App A 2.0' | Out-Null
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'ConfigMgr').SupersedencePending | Should -BeFalse
        }
        It 'clears the pending state when the version is published again' {
            $history = New-PendingHistory
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '2.0' -Id 'App A 2.0' | Out-Null
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'ConfigMgr').SupersedencePending | Should -BeFalse
        }
        It 'never marks WSUS or Intune as pending' {
            $history = @{}
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '2.0' -SupersedenceNotApplied | Out-Null
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'Intune' -Version '2.0' -SupersedenceNotApplied | Out-Null
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'WSUS').SupersedencePending | Should -BeFalse
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'Intune').SupersedencePending | Should -BeFalse
        }
        It 'survives a JSON round trip' {
            $history = New-PendingHistory
            $back = ($history | ConvertTo-Json -Depth 6 | ConvertFrom-Json)
            $rec = Get-OneClickPublishedVersion -HistoryEntry $back.'package-a' -Destination 'ConfigMgr'
            $rec.SupersedencePending | Should -BeTrue
            $rec.SupersedenceReason | Should -Be 'access denied'
        }
        It 'ends the pending state and keeps the rest of the record' {
            $history = New-PendingHistory -Instance 'mcm@cm01'
            Clear-OneClickSupersedencePending -History $history -PackagerName 'package-a' | Out-Null
            $rec = Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'ConfigMgr' -Instance 'mcm@cm01'
            $rec.SupersedencePending | Should -BeFalse
            $rec.SupersedenceReason | Should -Be ''
            $rec.Version | Should -Be '2.0'
            $rec.Id | Should -Be 'App A 2.0'
            $rec.Instance | Should -Be 'mcm@cm01'
        }
        It 'ends the pending state of a record read from JSON' {
            $back = (New-PendingHistory | ConvertTo-Json -Depth 6 | ConvertFrom-Json)
            $history = @{ 'package-a' = $back.'package-a' }
            Clear-OneClickSupersedencePending -History $history -PackagerName 'package-a' | Out-Null
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'ConfigMgr').SupersedencePending | Should -BeFalse
        }
        It 'does nothing for an application without history or without a ConfigMgr record' {
            { Clear-OneClickSupersedencePending -History @{} -PackagerName 'package-a' } | Should -Not -Throw
            $history = @{}
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'WSUS' -Version '1.0' | Out-Null
            { Clear-OneClickSupersedencePending -History $history -PackagerName 'package-a' } | Should -Not -Throw
            (Get-OneClickPublishedVersion -HistoryEntry $history['package-a'] -Destination 'WSUS').Version | Should -Be '1.0'
        }
    }

    Context 'the reason in the output of a publish' {
        It 'reads the reason of a publish that succeeded' {
            $text = "[2026-11-18 10:00:00] [INFO ] Created ConfigMgr application : App A 2.0`r`n[2026-11-18 10:00:05] [WARN ] Published, supersedence not applied: access denied`r`n[2026-11-18 10:00:06] [INFO ] Done"
            Get-OneClickSupersedenceFailure -Text $text | Should -Be 'access denied'
        }
        It 'gives a failed content step precedence: a failed publish is not a pending supersedence' {
            $text = '[WARN ] Published, supersedence not applied: access denied'
            Get-OneClickSupersedenceFailure -Text $text -ExitCode 1 | Should -Be ''
        }
        It 'gives nothing for a run without a failed step' {
            Get-OneClickSupersedenceFailure -Text '[INFO ] Supersedence                 : App A 2.0 supersedes App A 1.0' | Should -Be ''
            Get-OneClickSupersedenceFailure -Text '' | Should -Be ''
        }
        It 'says that the site gave no reason when the line has none' {
            Get-OneClickSupersedenceFailure -Text '[WARN ] Published, supersedence not applied:' | Should -Be 'the site gave no reason.'
        }
    }

    Context 'the plan' {
        It 'plans the pass and no publish for a version that is on the site with its step pending' {
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeTrue
            $row.PublishTo | Should -BeNullOrEmpty
            $row.Planned | Should -Be 'Check, Supersedence'
            $row.Reason | Should -Be 'ConfigMgr: 2.0 already published, supersedence pending'
            $row.Include | Should -BeTrue
        }
        It 'publishes the other destinations and runs the pass in the same row' {
            $history = New-PendingHistory
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECMAndWSUS') -History $history -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeTrue
            $row.PublishTo | Should -Be @('WSUS')
            $row.Planned | Should -Be 'Check, Stage, Publish to WSUS, Supersedence'
        }
        It 'publishes a newer version in full and runs no pass' {
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.1') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeFalse
            $row.PublishTo | Should -Be @('ConfigMgr')
        }
        It 'publishes again instead when Force is set' {
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'StageAndPackage' -Force)[0]
            $row.SupersedenceOnly | Should -BeFalse
            $row.PublishTo | Should -Be @('ConfigMgr')
        }
        It 'runs no pass for a version without a pending step' {
            $history = @{}
            Set-OneClickPublishedVersion -History $history -PackagerName 'package-a' -Destination 'ConfigMgr' -Version '2.0' | Out-Null
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History $history -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeFalse
            $row.Reason | Should -Be 'ConfigMgr: 2.0 already published'
        }
        It 'runs the pass only in a run that publishes' {
            $stage = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'Stage')[0]
            $stage.SupersedenceOnly | Should -BeFalse
            $stage.Planned | Should -Be 'Check, Stage'
            $report = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'Report' -Force)[0]
            $report.SupersedenceOnly | Should -BeFalse
        }
        It 'ignores a pending record that belongs to another site' {
            $history = New-PendingHistory -Instance 'xyz@other'
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History $history -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeFalse
            $row.PublishTo | Should -Be @('ConfigMgr')
        }
        It 'does not plan the pass for a destination that is not ready' {
            $prefs = New-TestPrefs -Target 'MECM'
            $prefs.SiteCode = ''
            $row = (Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs $prefs -History (New-PendingHistory) -Action 'StageAndPackage')[0]
            $row.SupersedenceOnly | Should -BeFalse
            $row.NotPublishable | Should -Contain 'ConfigMgr'
        }
        It 'names the passes in the count line' {
            $plan = @(Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -History (New-PendingHistory) -Action 'StageAndPackage')
            (Get-OneClickPlanSummary -Plan $plan -Prefs (New-TestPrefs -Target 'MECM') -Action 'StageAndPackage').Line | Should -Match ', 1 supersedence only$'
            $none = @(Get-OneClickPlan -Apps @(New-TestApp 'package-a' -Latest '2.0') -Prefs (New-TestPrefs -Target 'MECM') -Action 'StageAndPackage')
            (Get-OneClickPlanSummary -Plan $none -Prefs (New-TestPrefs -Target 'MECM') -Action 'StageAndPackage').Line | Should -Not -Match 'supersedence'
        }
    }

    Context 'the report' {
        It 'counts a pass and a publish whose step failed, and keeps the row as published' {
            $run = [pscustomobject]@{ Started = [datetime]'2026-04-05T10:00:00'; Ended = [datetime]'2026-04-05T10:01:00'; Operator = 'o'; Computer = 'c'; Action = 'StageAndPackage' }
            $rows = @(
                [pscustomobject]@{ Packager = 'package-a'; Application = 'App A'; Version = '2.0'; Outcome = 'Published'; ResultConfigMgr = 'published App A 2.0, supersedence not applied'; IdConfigMgr = 'App A 2.0'; DetailConfigMgr = 'Published, supersedence not applied: access denied'; Reason = '' }
                [pscustomobject]@{ Packager = 'package-b'; Application = 'App B'; Version = '1.1'; Outcome = 'Superseded'; ResultConfigMgr = 'supersedence applied: App B 1.1 supersedes App B 1.0'; IdConfigMgr = 'App B 1.1'; Reason = '' }
            )
            $paths = Write-OneClickReport -Run $run -Rows $rows -Folder (Join-Path $TestDrive 'OneClickSup')
            $md = Get-Content $paths.MarkdownPath -Raw
            $md | Should -Match '\| Published \| 1 \|'
            $md | Should -Match '\| Supersedence applied \| 1 \|'
            $md | Should -Match '\| Supersedence not applied \| 1 \|'
            $md | Should -Match 'App A, ConfigMgr: Published, supersedence not applied: access denied'
        }
        It 'leaves the two lines out of a report that has neither' {
            $run = [pscustomobject]@{ Started = [datetime]'2026-04-05T11:00:00'; Ended = [datetime]'2026-04-05T11:01:00'; Operator = 'o'; Computer = 'c'; Action = 'StageAndPackage' }
            $rows = @([pscustomobject]@{ Packager = 'package-a'; Application = 'App A'; Version = '2.0'; Outcome = 'Published'; ResultConfigMgr = 'published App A 2.0'; Reason = '' })
            $md = Get-Content (Write-OneClickReport -Run $run -Rows $rows -Folder (Join-Path $TestDrive 'OneClickSup2')).MarkdownPath -Raw
            $md | Should -Not -Match 'Supersedence'
        }
    }
}
