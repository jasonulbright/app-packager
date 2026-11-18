#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The window side of Rewrite Old Versions and Restore Old Versions: the buttons, the plan that the window shows, the run after the
    confirmation, the result rows, the log and the line of the target.

.DESCRIPTION
    The cases are the cases of RewriteOldVersionsPlansTests and RewriteLeftoversTests of App Packager Desktop (the plan of the window,
    the result rows, the same-rewrite check, the read before the write, the line of a restore). The engine is covered by
    SupersedenceRewrite.Tests.ps1. The window itself is covered by Invoke-RewriteOldVersionsSmoke.ps1.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:GuiPath = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:GuiPath, [ref]$null, [ref]$null)
    foreach ($name in 'Get-SidebarSystemState', 'Get-RewriteTargetForRow', 'Get-RewriteChildArgument', 'Get-WorkbenchProfileDownloadRoot', 'Get-UserErrorText') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:Xaml = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
    $script:Gui = Get-Content -LiteralPath $script:GuiPath -Raw
    $script:LineA = 'APD-Line:catalog:a|default|x64'
    $script:LineB = 'APD-Line:catalog:b|default|x64'

    function New-TestClause {
        param([string]$Expected = '1.0', [string]$LogicalName = 'Setting_1', [string]$Op = 'GreaterEquals')
        [pscustomobject]@{ Kind = 'File'; LogicalName = $LogicalName; Operator = $Op; ExpectedValue = $Expected; Path = 'C:\Program Files\App'; FileName = 'app.exe'; Is64Bit = $true; Hive = ''; Key = ''; ValueName = '' }
    }
    function New-TestItem {
        param([string]$App = 'App 1.0', [string]$Dt = 'DT', $Skip = $null, [string]$Expected = '1.0')
        [pscustomobject]@{ Application = $App; Version = '1.0'; DeploymentType = $Dt; Clause = (New-TestClause -Expected $Expected); Skip = $Skip; Rewrites = ($null -eq $Skip) }
    }
    function New-TestPlan {
        param([object[]]$Items = @(), [string[]]$Blocked = @())
        $rewriting = @($Items | Where-Object { $_.Rewrites })
        [pscustomobject]@{ Items = @($Items); Blocked = @($Blocked); ClauseCount = $rewriting.Count; ApplicationCount = $rewriting.Count }
    }
    function New-TestEntry {
        param([int]$Minute, [string]$App = 'App 1.0', [string]$Dt = 'DT', [string]$Line = $script:LineA, [bool]$Restored = $false, [string]$OldOp = 'GreaterEquals', [string]$NewOp = 'IsEquals')
        [pscustomobject]@{ At = [datetimeoffset]::new(2026, 11, 12, 10, $Minute, 0, [timespan]::Zero); Application = $App; DeploymentType = $Dt; OldOperator = $OldOp; NewOperator = $NewOp; LogicalName = 'Setting_1'; Restored = $Restored; Line = $Line }
    }
    function New-TestPrefs {
        param([bool]$ConfigMgr = $true, [bool]$ConsoleFound = $true, [string]$SiteCode = 'MCM')
        [pscustomobject]@{
            SiteCode      = $SiteCode
            Systems       = [pscustomobject]@{ ConfigMgr = $ConfigMgr; Intune = $false; Wsus = $false }
            DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $ConsoleFound } }
            Intune        = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = '' }
            Wsus          = [pscustomobject]@{ ServerName = '' }
            DiscoverySqlServer = ''; DiscoverySqlDatabase = ''
        }
    }
}

Describe 'The buttons' {
    It 'are in the sidebar XAML, with wrapped labels in Title Case' {
        $script:Xaml | Should -Match 'x:Name="btnRewriteOld"'
        $script:Xaml | Should -Match 'x:Name="btnRestoreOld"'
        $script:Xaml | Should -Match '<TextBlock Text="Rewrite Old Versions" TextWrapping="Wrap"/>'
        $script:Xaml | Should -Match '<TextBlock Text="Restore Old Versions" TextWrapping="Wrap"/>'
    }

    It 'are wired into the window, the theme list, the enable switch, the system state and a click handler' {
        foreach ($name in 'btnRewriteOld', 'btnRestoreOld') {
            $script:Gui | Should -Match ("\`${0}\s*=\s*\`$window\.FindName\('{0}'\)" -f $name)
            $script:Gui | Should -Match ('\$script:WorkflowButtons = @\([^)]*\${0}' -f $name)
            $script:Gui | Should -Match ('\${0}\.IsEnabled\s*=\s*\$Enabled' -f $name)
            $script:Gui | Should -Match ('\${0}\.Add_Click' -f $name)
        }
        $script:Gui | Should -Match "@\(\`$btnRewriteOld, 'Rewrite'\)"
        $script:Gui | Should -Match "@\(\`$btnRestoreOld, 'Restore'\)"
    }

    It 'are enabled when ConfigMgr is in use, the console is found and the site code is set, and say what they do' {
        $s = Get-SidebarSystemState -Prefs (New-TestPrefs)
        $s.RewriteEnabled | Should -BeTrue
        $s.RestoreEnabled | Should -BeTrue
        $s.RewriteToolTip | Should -Match 'from GreaterEquals to IsEquals'
        $s.RewriteToolTip | Should -Match 'Shows the plan first and keeps a log; Undo restores the old operator\. Check one application\.'
        $s.RestoreToolTip | Should -Match 'from the log of an earlier Rewrite Old Versions'
    }

    It 'are disabled with the reason of Publish to ConfigMgr: <Name>' -ForEach @(
        @{ Name = 'the console is missing'; ConfigMgr = $true; Console = $false; Site = 'MCM'; Reason = 'Needs the ConfigMgr console on this computer.' }
        @{ Name = 'ConfigMgr is not in use'; ConfigMgr = $false; Console = $true; Site = 'MCM'; Reason = 'ConfigMgr is not in use.' }
        @{ Name = 'the site code is empty'; ConfigMgr = $true; Console = $true; Site = ''; Reason = 'Set the site code in Options, Connections, ConfigMgr.' }
    ) {
        $s = Get-SidebarSystemState -Prefs (New-TestPrefs -ConfigMgr $ConfigMgr -ConsoleFound $Console -SiteCode $Site)
        $s.RewriteEnabled | Should -BeFalse
        $s.RestoreEnabled | Should -BeFalse
        $s.RewriteToolTip | Should -BeLike ('Rewrite Old Versions: ' + $Reason + '*')
        $s.RestoreToolTip | Should -BeLike ('Restore Old Versions: ' + $Reason + '*')
        $s.RewriteToolTip.Substring('Rewrite Old Versions: '.Length) | Should -Be ($s.MecmToolTip.Substring('Publish to ConfigMgr: '.Length))
    }
}

Describe 'The plan of the window for a rewrite' {
    It 'has no window when no clause can be rewritten' {
        $plan = New-TestPlan -Items @((New-TestItem -Skip 'not rewritten: no installed value'))
        New-SupersedenceRewriteChangePlan -Plan $plan -Restorable 0 | Should -BeNullOrEmpty
    }

    It 'lists one row per clause, with the confirmation, the notes and Undo' {
        $plan = New-TestPlan -Items @((New-TestItem), (New-TestItem -App 'App 0.9' -Skip 'the operator is Equals, not GreaterEquals')) -Blocked @("'Old 0.9' has a required deployment (All Users).")
        $change = New-SupersedenceRewriteChangePlan -Plan $plan -Restorable 2
        $change.Items.Count | Should -Be 1
        $change.Items[0].Target | Should -Be 'App 1.0 / DT'
        $change.Items[0].Detail | Should -Be 'GreaterEquals -> IsEquals'
        $change.Items[0].Action | Should -Be 'Rewrite clause'
        $change.HasUndo | Should -BeTrue
        $change.IsDestructive | Should -BeTrue
        $change.DestructiveConfirmation | Should -Be (Get-SupersedenceRewriteConfirmationText)
        $change.Title | Should -Be 'Rewrite Old Versions'
        $change.ActionName | Should -Be 'Rewrite 1 Clause'
        @($change.Notes | Where-Object { $_ -like "Blocked: 'Old 0.9'*" }).Count | Should -Be 1
        @($change.Notes | Where-Object { $_ -like 'Left: App 0.9 / DT: the operator is Equals*' }).Count | Should -Be 1
        @($change.Notes | Where-Object { $_ -like '2 earlier rewrite(s)*' }).Count | Should -Be 1
        $change.Notes[0] | Should -Match 'reads "not installed"'
    }

    It 'names the button for the number of clauses' {
        $plan = New-TestPlan -Items @((New-TestItem), (New-TestItem -Dt 'DT 2'))
        (New-SupersedenceRewriteChangePlan -Plan $plan).ActionName | Should -Be 'Rewrite 2 Clauses'
    }

    It 'does not mention a pending restore when none is pending' {
        $change = New-SupersedenceRewriteChangePlan -Plan (New-TestPlan -Items @((New-TestItem))) -Restorable 0
        @($change.Notes | Where-Object { $_ -like '*earlier rewrite*' }).Count | Should -Be 0
    }
}

Describe 'The plan of the window for a restore' {
    It 'names the operator that comes back' {
        $item = ConvertTo-SupersedenceRestoreChangeItem -Entry (New-TestEntry -Minute 1)
        $item.Action | Should -Be 'Restore clause'
        $item.Target | Should -Be 'App 1.0 / DT'
        $item.Detail | Should -Be 'IsEquals -> GreaterEquals'
    }

    It 'has no window when nothing is pending' {
        New-SupersedenceRestoreChangePlan -Pending @() | Should -BeNullOrEmpty
    }

    It 'is destructive, has no Undo and says what a device may read' {
        $change = New-SupersedenceRestoreChangePlan -Pending @((New-TestEntry -Minute 1), (New-TestEntry -Minute 2 -Dt 'DT 2'))
        $change.Title | Should -Be 'Restore Old Versions'
        $change.ActionName | Should -Be 'Restore 2 Clauses'
        $change.HasUndo | Should -BeFalse
        $change.IsDestructive | Should -BeTrue
        $change.DestructiveConfirmation | Should -Be (Get-SupersedenceRestoreConfirmationText)
        $change.DestructiveConfirmation | Should -Match 'may read it differently'
    }

    It 'lists the rewrites of this line that no restore cancels, and leaves the other lines' {
        $log = Join-Path $TestDrive 'pending.jsonl'
        $text = (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 1 -App 'A 1.0' -Line $script:LineA)) + "`n" +
        (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 2 -App 'B 1.0' -Line $script:LineB)) + "`n" +
        (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 3 -App 'C 1.0' -Line $script:LineA)) + "`n" +
        (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 4 -App 'C 1.0' -Line $script:LineA -Restored $true -OldOp 'IsEquals' -NewOp 'GreaterEquals')) + "`n"
        Set-Content -LiteralPath $log -Value $text -Encoding UTF8
        @(Get-SupersedenceRewritePending -LogPath $log -Marker $script:LineA | ForEach-Object { $_.Application }) | Should -Be @('A 1.0')
        @(Get-SupersedenceRewritePending -LogPath $log -Marker $script:LineB | ForEach-Object { $_.Application }) | Should -Be @('B 1.0')
    }

    It 'has nothing pending without a log' {
        @(Get-SupersedenceRewritePending -LogPath (Join-Path $TestDrive 'none.jsonl') -Marker $script:LineA).Count | Should -Be 0
    }
}

Describe 'The result rows' {
    It 'follow what was written, what failed and what was not reached' {
        $items = @('A / one', 'B / two', 'C / three') | ForEach-Object { $a, $d = $_ -split ' / '; [pscustomobject]@{ Action = 'Rewrite clause'; Target = $_; Detail = $null; Application = $a; DeploymentType = $d } }
        $outcome = [pscustomobject]@{ Written = @((New-TestEntry -Minute 1 -App 'A' -Dt 'one')); Failures = @('B / two: the site refused it') }
        $result = ConvertTo-SupersedenceChangeResult -Items $items -Outcome $outcome
        $result.Items[0].Succeeded | Should -BeTrue
        $result.Items[1].Succeeded | Should -BeFalse
        $result.Items[1].Message | Should -Be 'the site refused it'
        $result.Items[2].Skipped | Should -BeTrue
        $result.UndoAvailable | Should -BeTrue
        $result.Summary | Should -Be '1 changed, 1 failed, 1 skipped.'
        @($result.Items | ForEach-Object { $_.Outcome }) | Should -Be @('Changed', 'Failed', 'Skipped')
    }

    It 'say changed and failed when nothing is skipped, and name a change made in part' {
        $row = New-SupersedenceChangeRow -Item ([pscustomobject]@{ Action = 'a'; Target = 't' }) -Succeeded $false -Partial $true -Message 'x'
        Get-SupersedenceChangeSummary -Rows @($row) | Should -Be '0 changed, 1 failed. 1 failed change was made in part; Undo can reverse it.'
        Get-SupersedenceChangeSummary -Rows @($row, $row) | Should -Be '0 changed, 2 failed. 2 failed changes were made in part; Undo can reverse them.'
    }

    It 'offers no Undo when nothing changed' {
        $items = @([pscustomobject]@{ Action = 'Rewrite clause'; Target = 'A / one'; Detail = $null; Application = 'A'; DeploymentType = 'one' })
        (ConvertTo-SupersedenceChangeResult -Items $items -Outcome ([pscustomobject]@{ Written = @(); Failures = @('A / one: no') })).UndoAvailable | Should -BeFalse
    }
}

Describe 'The same rewrite' {
    It 'is the same for two plans with the same clauses, whatever the case of the application' {
        Test-SupersedenceSameRewrite -Shown (New-TestPlan -Items @((New-TestItem))) -Fresh (New-TestPlan -Items @((New-TestItem))) | Should -BeTrue
        Test-SupersedenceSameRewrite -Shown (New-TestPlan -Items @((New-TestItem))) -Fresh (New-TestPlan -Items @((New-TestItem -App 'APP 1.0'))) | Should -BeTrue
    }

    It 'is another rewrite when the expected value or the application changed: <Expected> <App>' -ForEach @(
        @{ Expected = '2.0'; App = 'App 1.0' }
        @{ Expected = '1.0'; App = 'Another 1.0' }
    ) {
        Test-SupersedenceSameRewrite -Shown (New-TestPlan -Items @((New-TestItem))) -Fresh (New-TestPlan -Items @((New-TestItem -App $App -Expected $Expected))) | Should -BeFalse
    }

    It 'is another rewrite when a clause is lost while the window is open' {
        $lost = New-TestPlan -Items @((New-TestItem -Skip 'the operator is Equals, not GreaterEquals'))
        Test-SupersedenceSameRewrite -Shown (New-TestPlan -Items @((New-TestItem))) -Fresh $lost | Should -BeFalse
    }
}

Describe 'The run after the confirmation' {
    BeforeAll {
        $script:Run = { param($Plan, $Replan, $Write) Invoke-SupersedenceRewriteRun -Plan $Plan -SiteCode 'MCM' -Marker $script:LineA -TitleKey 'catalog:a' -Settings (New-SupersedenceSettings) -LogPath (Join-Path $TestDrive 'run.jsonl') -Replan $Replan -Write $Write }
    }

    It 'writes nothing when the site changed after the plan was shown' {
        $shown = New-TestPlan -Items @((New-TestItem))
        $script:Wrote = $false
        $result = & $script:Run $shown { New-TestPlan -Items @((New-TestItem -Expected '2.0')) } { param($p) $script:Wrote = $true; [pscustomobject]@{ Written = @(); Failures = @() } }
        $script:Wrote | Should -BeFalse
        $result.Stop | Should -Be 'Failed'
        @($result.Items | Where-Object { -not $_.Skipped }).Count | Should -Be 0
        $result.Items[0].Message | Should -Match 'site changed'
        $result.UndoAvailable | Should -BeFalse
    }

    It 'writes nothing and gives the reason when the site cannot be read again' {
        $shown = New-TestPlan -Items @((New-TestItem))
        $script:Wrote = $false
        $result = & $script:Run $shown { throw "The deployments of 'App 1.0' could not be read." } { param($p) $script:Wrote = $true; [pscustomobject]@{ Written = @(); Failures = @() } }
        $script:Wrote | Should -BeFalse
        $result.Stop | Should -Be 'Failed'
        $result.Items[0].Message | Should -Match 'could not be read'
        $result.UndoAvailable | Should -BeFalse
    }

    It 'writes the shown plan when the plan is unchanged' {
        $shown = New-TestPlan -Items @((New-TestItem))
        $script:Written = $null
        $result = & $script:Run $shown { New-TestPlan -Items @((New-TestItem)) } { param($p) $script:Written = $p; [pscustomobject]@{ Written = @((New-TestEntry -Minute 1)); Failures = @() } }
        [object]::ReferenceEquals($script:Written, $shown) | Should -BeTrue
        $result.Items[0].Succeeded | Should -BeTrue
        $result.Stop | Should -Be 'Completed'
    }

    It 'lists the pending clauses of the line before a restore, and gives each one a row' {
        $log = Join-Path $TestDrive 'restore-run.jsonl'
        Set-Content -LiteralPath $log -Value ((ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 1 -App 'A 1.0' -Dt 'one')) + "`n" + (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 2 -App 'A 1.0' -Dt 'two'))) -Encoding UTF8
        $result = Invoke-SupersedenceRestoreRun -SiteCode 'MCM' -Marker $script:LineA -LogPath $log -Write { [pscustomobject]@{ Written = @((New-TestEntry -Minute 9 -App 'A 1.0' -Dt 'one' -Restored $true)); Failures = @('A 1.0 / two: access denied') } }
        @($result.Items | ForEach-Object { $_.Target }) | Should -Be @('A 1.0 / one', 'A 1.0 / two')
        $result.Items[0].Succeeded | Should -BeTrue
        $result.Items[1].Message | Should -Be 'access denied'
        $result.Summary | Should -Be '1 changed, 1 failed.'
    }
}

Describe 'A run that stopped' {
    BeforeAll {
        $script:StopItems = @('A / one', 'B / two') | ForEach-Object { $a, $d = $_ -split ' / '; [pscustomobject]@{ Action = 'Rewrite clause'; Target = $_; Detail = $null; Application = $a; DeploymentType = $d } }
    }

    It 'reports a clause with an open intent in the log as possibly changed, and the others as not reported' {
        $result = Get-SupersedenceStoppedResult -Items $script:StopItems -Entries @((New-TestEntry -Minute 1 -App 'A' -Dt 'one')) -Stop Cancelled
        $result.Stop | Should -Be 'Cancelled'
        $result.Items[0].PartialChange | Should -BeTrue
        $result.Items[0].Message | Should -Match 'may be changed; Undo restores it'
        $result.Items[1].PartialChange | Should -BeFalse
        $result.Items[1].Message | Should -Be 'No result: the run was cancelled before this change reported.'
        $result.UndoAvailable | Should -BeTrue
        $result.Summary | Should -Match '^The run was cancelled\. Changes made before the cancel stay in place\.'
    }

    It 'reports a clause whose intent a restore entry cancelled as not changed' {
        $result = Get-SupersedenceStoppedResult -Items $script:StopItems -Entries @((New-TestEntry -Minute 1 -App 'A' -Dt 'one'), (New-TestEntry -Minute 2 -App 'A' -Dt 'one' -Restored $true)) -Stop Failed
        $result.Items[0].PartialChange | Should -BeFalse
        $result.UndoAvailable | Should -BeFalse
        $result.Summary | Should -Match '^The run failed\. The log has the details\.'
    }

    It 'reports a restore entry as a clause that is back' {
        $result = Get-SupersedenceStoppedResult -Items $script:StopItems -Entries @((New-TestEntry -Minute 2 -App 'A' -Dt 'one' -Restored $true)) -Stop Cancelled -Kind Restore
        $result.Items[0].Succeeded | Should -BeTrue
        $result.Items[1].Succeeded | Should -BeFalse
        $result.Summary | Should -Match 'Before the cancel: 1 changed, 1 failed\.'
    }

    It 'reads only the entries that the run wrote' {
        $log = Join-Path $TestDrive 'since.jsonl'
        Set-Content -LiteralPath $log -Value ((ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 1)) + "`n" + (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry -Minute 5 -Dt 'DT 2'))) -Encoding UTF8
        $since = [datetimeoffset]::new(2026, 11, 12, 10, 3, 0, [timespan]::Zero)
        @(Get-SupersedenceLogEntrySince -LogPath $log -Since $since | ForEach-Object { $_.DeploymentType }) | Should -Be @('DT 2')
        @(Get-SupersedenceLogEntrySince -LogPath (Join-Path $TestDrive 'nolog.jsonl') -Since $since).Count | Should -Be 0
    }

    It 'takes the result of the child when it gave one, and reports a child with none' {
        $done = [pscustomobject]@{ Kind = 'Result'; Result = (New-SupersedenceChangeResult -Rows @((New-SupersedenceChangeRow -Item $script:StopItems[0] -Succeeded $true -Message 'Written to the site and the log.'))) }
        $viaJson = ($done | ConvertTo-Json -Depth 8 -Compress) | ConvertFrom-Json
        $ok = ConvertFrom-SupersedenceChildRun -Run ([pscustomobject]@{ Cancelled = $false; Result = $viaJson; ExitCode = 0; StdErr = '' }) -Items $script:StopItems
        $ok.Items[0].Succeeded | Should -BeTrue
        $ok.Stop | Should -Be 'Completed'
        $none = ConvertFrom-SupersedenceChildRun -Run ([pscustomobject]@{ Cancelled = $false; Result = $null; ExitCode = 1; StdErr = 'boom' }) -Items $script:StopItems
        $none.Stop | Should -Be 'Failed'
        $none.Summary | Should -Match 'boom'
        (ConvertFrom-SupersedenceChildRun -Run ([pscustomobject]@{ Cancelled = $true; Result = $null; ExitCode = 1; StdErr = '' }) -Items $script:StopItems).Stop | Should -Be 'Cancelled'
    }
}

Describe 'The line of the target' {
    BeforeAll {
        function New-ApplicationId { param($Kind, $ScriptPath) 'catalog:' + [IO.Path]::GetFileNameWithoutExtension($ScriptPath) }
        function Get-ApplicationDefinition { param($ApplicationId) [pscustomobject]@{ ActiveProfileId = $script:Active } }
        function Find-NewestStageManifestForPackager { param($PackagerPath, $DownloadRoot) $script:SeenRoot = $DownloadRoot; $script:ManifestPath }
        function Read-StageManifest { param($Path) if ($script:ManifestThrows) { throw 'bad manifest' }; $script:Manifest }
        $script:Row = [pscustomobject]@{ Application = 'App'; FullPath = 'C:\x\package-app.ps1'; Script = 'package-app.ps1' }
    }
    BeforeEach { $script:Active = ''; $script:ManifestPath = $null; $script:Manifest = $null; $script:ManifestThrows = $false; $script:SeenRoot = $null }

    It 'makes the marker of the application, the default profile and x64 for a row that was never staged' {
        $t = Get-RewriteTargetForRow -Row $script:Row -DownloadRoot 'C:\dl'
        $t.Ok | Should -BeTrue
        $t.Marker | Should -Be 'APD-Line:catalog:package-app|default|x64'
        $t.ApplicationId | Should -Be 'catalog:package-app'
    }

    It 'reads the marker from the stage manifest, as the publish wrote it' {
        $script:ManifestPath = 'C:\dl\app\stage-manifest.json'
        $script:Manifest = [pscustomobject]@{ ApplicationId = 'catalog:package-app'; ProfileId = 'managed'; Architecture = 'arm64' }
        $t = Get-RewriteTargetForRow -Row $script:Row -DownloadRoot 'C:\dl'
        $t.Marker | Should -Be 'APD-Line:catalog:package-app|managed|arm64'
    }

    It 'falls back to the application of the row when the manifest has no identifier' {
        $script:ManifestPath = 'C:\dl\app\stage-manifest.json'
        $script:Manifest = [pscustomobject]@{ Architecture = 'x86' }
        $t = Get-RewriteTargetForRow -Row $script:Row -DownloadRoot 'C:\dl'
        $t.Marker | Should -Be 'APD-Line:catalog:package-app|default|x86'
        $t.ApplicationId | Should -Be 'catalog:package-app'
    }

    It 'searches the folder of the active profile' {
        $script:Active = 'managed'
        [void](Get-RewriteTargetForRow -Row $script:Row -DownloadRoot 'C:\dl')
        $script:SeenRoot | Should -Be 'C:\dl\profiles\managed'
    }

    It 'names a manifest that cannot be read' {
        $script:ManifestPath = 'C:\dl\app\stage-manifest.json'
        $script:ManifestThrows = $true
        $t = Get-RewriteTargetForRow -Row $script:Row -DownloadRoot 'C:\dl'
        $t.Ok | Should -BeFalse
        $t.Error | Should -Match 'could not be read: bad manifest'
    }
}

Describe 'The child command line' {
    BeforeAll { $script:T = [pscustomobject]@{ ApplicationId = 'catalog:package-app'; Marker = 'APD-Line:catalog:package-app|default|x64' } }

    It 'reads the plan without -Yes' {
        $a = Get-RewriteChildArgument -Mode plan -Target $script:T -SiteCode 'MCM' -ScriptPath 'C:\ap\Invoke-AppPackagerBuild.ps1'
        $a | Should -Contain '-Json'
        $a | Should -Not -Contain '-Yes'
        $a[$a.IndexOf('-RewriteOldVersions') + 1] | Should -Be 'plan'
        $a[$a.IndexOf('-IdentityMarker') + 1] | Should -Be 'APD-Line:catalog:package-app|default|x64'
        $a[$a.IndexOf('-SiteCode') + 1] | Should -Be 'MCM'
        $a[$a.IndexOf('-File') + 1] | Should -Be 'C:\ap\Invoke-AppPackagerBuild.ps1'
    }

    It 'applies the shown plan with -Yes and names the provider' {
        $a = Get-RewriteChildArgument -Mode apply -Target $script:T -SiteCode 'MCM' -ProviderMachineName ' cm01 ' -ShownPlan 'C:\t\plan.json' -ScriptPath 'x.ps1'
        $a | Should -Contain '-Yes'
        $a[$a.IndexOf('-ShownPlan') + 1] | Should -Be 'C:\t\plan.json'
        $a[$a.IndexOf('-ProviderMachineName') + 1] | Should -Be 'cm01'
    }

    It 'restores with -Yes and no shown plan' {
        $a = Get-RewriteChildArgument -Mode restore -Target $script:T -SiteCode 'MCM' -ShownPlan 'C:\t\plan.json' -ScriptPath 'x.ps1'
        $a | Should -Contain '-Yes'
        $a | Should -Not -Contain '-ShownPlan'
    }

    It 'has the parameters on Invoke-AppPackagerBuild.ps1, and refuses them without -RewriteOldVersions' {
        $cli = Join-Path $script:RepoRoot 'Invoke-AppPackagerBuild.ps1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($cli, [ref]$null, [ref]$null)
        $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($n in 'Json', 'IdentityMarker', 'ShownPlan') { $names | Should -Contain $n }
        { & $cli -Application 'catalog:package-x' -Stage -Json } | Should -Throw '*apply to -RewriteOldVersions only*'
        { & $cli -Application 'catalog:package-x' -RewriteOldVersions plan -ShownPlan 'x.json' } | Should -Throw '*-ShownPlan applies to -RewriteOldVersions apply only*'
    }
}

Describe 'The window code' {
    It 'asks for the confirmation before it writes, and keeps the way back' {
        $text = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-RewriteOldVersionsDialog' }, $false)).Extent.Text
        $text | Should -Match 'if \(\$needsCheck -and \$chkConfirm\.IsChecked -ne \$true\) \{ return \}'
        $text | Should -Match '\$btnConfirm\.IsEnabled = \(\$review -and'
        $text | Should -Match 'Title="Rewrite Old Versions"'
    }

    It 'uses the engine and does not repeat it' {
        $start = ($script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-RewriteOldVersionsRun' }, $false)).Extent.Text
        $start | Should -Match 'Get-SupersedenceRewritePending'
        $start | Should -Match 'New-SupersedenceRewriteChangePlan'
        $start | Should -Match 'New-SupersedenceRestoreChangePlan'
        $start | Should -Not -Match 'Set-CMScriptDeploymentType'
        $start | Should -Not -Match 'Invoke-SupersedenceRewrite\b'
    }
}
