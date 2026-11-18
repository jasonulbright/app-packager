#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The optional rewrite of the version clause of an old application to Exact, and its log and restore.

.DESCRIPTION
    The cases are the cases of the tests of App Packager Desktop (SupersedenceRewriteTests and RewriteLeftoversTests). The reader
    reads the one-clause detections that the site stores. The plan blocks an application with a required deployment and rewrites
    only a GreaterEquals clause whose expected value is the version of the old application, and for which the Exact guard passes.
    The log is the only way back.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Fixtures = Join-Path $PSScriptRoot 'Fixtures\Supersedence'
    $script:RealFileXml = Get-Content -LiteralPath (Join-Path $script:Fixtures 'file-detection.xml') -Raw -Encoding UTF8
    $script:RealRegistryXml = Get-Content -LiteralPath (Join-Path $script:Fixtures 'registry-detection.xml') -Raw -Encoding UTF8
    $script:Allow = { param($app, $clause) [pscustomobject]@{ Result = 'ExactApplied'; Detail = $null } }

    function New-TestApp {
        param([string]$Title, [string]$Version, [string[]]$Dts, [bool]$Expired = $false)
        $i = 0
        [pscustomobject]@{ CiId = '1'; Title = $Title; SoftwareVersion = $Version; IsExpired = $Expired; DeploymentTypes = @($Dts | ForEach-Object { $i++; [pscustomobject]@{ CiId = "$i"; Name = $_ } }) }
    }
    function New-TestFileClause {
        param([string]$Op = 'GreaterEquals', [string]$Expected = '1.0.0', [string]$LogicalName = 'File_1', [string]$Path = 'C:\Vendor')
        [pscustomobject]@{ Kind = 'File'; LogicalName = $LogicalName; Operator = $Op; ExpectedValue = $Expected; Path = $Path; FileName = 'app.exe'; Is64Bit = $false; Hive = ''; Key = ''; ValueName = '' }
    }
    function New-TestPlan {
        param($App, $Detection, [string[]]$Required = @(), [scriptblock]$Guard = $script:Allow)
        $detections = @{ (Get-SupersedenceRewriteKey $App.Title $App.DeploymentTypes[0].Name) = $Detection }
        Get-SupersedenceRewritePlan -Old @($App) -Detections $detections -RequiredCollections @{ $App.Title = $Required } -Guard $Guard
    }
    function New-TestEntry {
        param([int]$Minute, [string]$App = 'Old 1.0', [string]$Dt = 'Old 1.0 - x64', [string]$OldOp = 'GreaterEquals', [string]$NewOp = 'IsEquals', [bool]$Restored = $false, [string]$Line = $null)
        [pscustomobject]@{ At = [datetimeoffset]::new(2026, 11, 12, 10, $Minute, 0, [timespan]::Zero); Application = $App; DeploymentType = $Dt; OldOperator = $OldOp; NewOperator = $NewOp; LogicalName = 'File_1'; Restored = $Restored; Line = $Line }
    }
}

Describe 'Reading what the site stores' {
    It 'reads the file clause of a real deployment type' {
        $detection = Read-SupersedenceStoredDetection $script:RealFileXml
        $detection.NotRewritableReason | Should -BeNullOrEmpty
        $clause = $detection.Clause
        $clause.Kind | Should -Be 'File'
        $clause.Operator | Should -Be 'GreaterEquals'
        $clause.ExpectedValue | Should -Be '1.0.0'
        $clause.Path | Should -Be 'C:\APDLab'
        $clause.FileName | Should -Be 'app.exe'
        $clause.Is64Bit | Should -BeFalse
        $clause.LogicalName | Should -BeLike 'File_*'
    }

    It 'reads the registry clause of a real deployment type' {
        $clause = (Read-SupersedenceStoredDetection $script:RealRegistryXml).Clause
        $clause.Kind | Should -Be 'Registry'
        $clause.Operator | Should -Be 'GreaterEquals'
        $clause.ExpectedValue | Should -Be '1.0.0'
        $clause.Hive | Should -Be 'HKEY_LOCAL_MACHINE'
        $clause.Key | Should -Be 'SOFTWARE\APDLab'
        $clause.ValueName | Should -Be 'DisplayVersion'
        $clause.Is64Bit | Should -BeTrue
        $clause.LogicalName | Should -BeLike 'RegSetting_*'
    }

    It 'does not rewrite a document that is not an enhanced detection: <Xml>' -ForEach @(
        @{ Xml = $null }, @{ Xml = '<<<not xml' }, @{ Xml = '<AppMgmtDigest><DeploymentType/></AppMgmtDigest>' }
    ) {
        $detection = Read-SupersedenceStoredDetection $Xml
        $detection.Clause | Should -BeNullOrEmpty
        $detection.NotRewritableReason | Should -Not -BeNullOrEmpty
    }

    It 'does not rewrite a detection with two clauses' {
        $two = $script:RealFileXml.Replace('&lt;/Settings&gt;', '&lt;File LogicalName="File_2"&gt;&lt;Path&gt;C:\x&lt;/Path&gt;&lt;Filter&gt;y&lt;/Filter&gt;&lt;/File&gt;&lt;/Settings&gt;')
        (Read-SupersedenceStoredDetection $two).NotRewritableReason | Should -Match 'more than one clause'
    }
}

Describe 'The plan' {
    It 'rewrites a clause that passes every check' {
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause); NotRewritableReason = $null })
        $plan.ClauseCount | Should -Be 1
        $plan.ApplicationCount | Should -Be 1
        @($plan.Blocked).Count | Should -Be 0
    }

    It 'blocks an application with a required deployment and names it' {
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause); NotRewritableReason = $null }) -Required @('All Workstations')
        $plan.ClauseCount | Should -Be 0
        @($plan.Blocked)[0] | Should -Match 'required deployment \(All Workstations\)'
    }

    It 'leaves a clause with operator <Op> and expected value <Expected> with the reason' -ForEach @(
        @{ Op = 'Equals'; Expected = '1.0.0'; Reason = 'the operator is Equals' }
        @{ Op = 'GreaterEquals'; Expected = '0.9.0'; Reason = 'is not the version of the application' }
    ) {
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause -Op $Op -Expected $Expected); NotRewritableReason = $null })
        $plan.ClauseCount | Should -Be 0
        @($plan.Items)[0].Skip | Should -Match $Reason
    }

    It 'does not rewrite a clause that fails the Exact guard and lists the reason' {
        $guard = { param($app, $clause) [pscustomobject]@{ Result = 'ExactBlocked'; Detail = 'no installed value is recorded for version 1.0.0' } }
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause); NotRewritableReason = $null }) -Guard $guard
        $plan.ClauseCount | Should -Be 0
        @($plan.Items)[0].Skip | Should -Match 'no installed value'
    }

    It 'leaves a detection that cannot be read with its reason' {
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = $null; NotRewritableReason = 'the detection has more than one clause or a connector, which is not rewritten' })
        @($plan.Items)[0].Skip | Should -Match 'connector'
    }

    It 'leaves a registry clause of another hive at plan time with its reason' {
        $clause = [pscustomobject]@{ Kind = 'Registry'; LogicalName = 'Reg_1'; Operator = 'GreaterEquals'; ExpectedValue = '1.0.0'; Path = ''; FileName = ''; Is64Bit = $true; Hive = 'HKEY_CLASSES_ROOT'; Key = 'K'; ValueName = 'V' }
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = $clause; NotRewritableReason = $null })
        $plan.ClauseCount | Should -Be 0
        @($plan.Items)[0].Skip | Should -Match 'HKEY_CLASSES_ROOT'
    }

    It 'never lets two applications that hold a type of the same name share a detection' {
        $one = New-TestApp 'Foo 1.0' '1.0' @('Install')
        $two = New-TestApp 'Foo 1.0 x' '1.0' @('Install')
        $detections = @{
            (Get-SupersedenceRewriteKey 'Foo 1.0' 'Install') = [pscustomobject]@{ Clause = (New-TestFileClause -LogicalName 'File_ONE' -Path 'C:\one' -Expected '1.0'); NotRewritableReason = $null }
            (Get-SupersedenceRewriteKey 'Foo 1.0 x' 'Install') = [pscustomobject]@{ Clause = (New-TestFileClause -LogicalName 'File_TWO' -Path 'C:\two' -Expected '1.0'); NotRewritableReason = $null }
        }
        $plan = Get-SupersedenceRewritePlan -Old @($one, $two) -Detections $detections -RequiredCollections @{} -Guard $script:Allow
        (@($plan.Items | ForEach-Object { $_.Clause.LogicalName }) -join ',') | Should -Be 'File_ONE,File_TWO'
    }

    It 'needs a recorded value for the version of the old application in the guard of the runner' {
        $settings = New-SupersedenceSettings
        $settings.InstalledValues[(Get-SupersedenceRecordKey 'catalog:vendor' '1.0.0')] = '1.0.0'
        $app = New-TestApp 'Old' '1.0.0' @('d')
        (Get-SupersedenceRewriteGuardVerdict -Settings $settings -TitleKey 'catalog:vendor' -Application $app -Clause (New-TestFileClause)).Result | Should -Be 'ExactApplied'
        # A value for another version is no value for this one; a value that differs blocks.
        (Get-SupersedenceRewriteGuardVerdict -Settings $settings -TitleKey 'catalog:vendor' -Application (New-TestApp 'Old' '0.9.0' @('d')) -Clause (New-TestFileClause -Expected '0.9.0')).Result | Should -Be 'ExactBlocked'
        $settings.InstalledValues[(Get-SupersedenceRecordKey 'catalog:vendor' '1.0.0')] = '1.0.0.5'
        (Get-SupersedenceRewriteGuardVerdict -Settings $settings -TitleKey 'catalog:vendor' -Application $app -Clause (New-TestFileClause)).Detail | Should -Match '1\.0\.0\.5'
    }

    It 'says in the confirmation that a device with a newer version reads not installed' {
        Get-SupersedenceRewriteConfirmationText | Should -Match 'not installed'
    }
}

Describe 'Which applications are old' {
    It 'counts every application that holds the newest version as current, also under another title and form' {
        $line = @((New-TestApp 'Foo 2.0' '2.0' @('d')), (New-TestApp 'Foo 2.0 (renamed)' '2.0.0' @('d')), (New-TestApp 'Foo 1.0' '1.0' @('d')))
        (@(Get-SupersedenceOldOnes -Line $line | ForEach-Object { $_.Title }) -join '|') | Should -Be 'Foo 1.0'
    }

    It 'leaves nothing for a single application, and only the older ones when the newest is retired' {
        @(Get-SupersedenceOldOnes -Line @(New-TestApp 'Foo 1.0' '1.0' @('d'))).Count | Should -Be 0
        $line = @((New-TestApp 'Foo 3.0' '3.0' @('d') $true), (New-TestApp 'Foo 2.0' '2.0' @('d')), (New-TestApp 'Foo 1.0' '1.0' @('d')))
        (@(Get-SupersedenceOldOnes -Line $line | ForEach-Object { $_.Title }) -join '|') | Should -Be 'Foo 1.0'
        @(Get-SupersedenceOldOnes -Line @(New-TestApp 'Foo x' 'not-a-version' @('d'))).Count | Should -Be 0
    }
}

Describe 'The log and the restore' {
    It 'round trips the log and skips a line that cannot be read' {
        $text = (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry 1)) + "`nnot json`n" + (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry 2 -App 'Other'))
        (@(Read-SupersedenceRewriteLog $text | ForEach-Object { $_.Application }) -join '|') | Should -Be 'Old 1.0|Other'
    }

    It 'keeps the date of an entry through the log' {
        $entry = @(Read-SupersedenceRewriteLog (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry 7)))[0]
        $entry.At.UtcDateTime | Should -Be ([datetime]::SpecifyKind([datetime]'2026-11-12T10:07:00', 'Utc'))
    }

    It 'writes back the first operator of each rewritten type and only for the line' {
        $log = @((New-TestEntry 1), (New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals'), (New-TestEntry 3 -App 'Other App'))
        $restore = @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0'))
        $restore.Count | Should -Be 1
        $restore[0].OldOperator | Should -Be 'GreaterEquals'
        $restore[0].Application | Should -Be 'Old 1.0'
    }

    It 'does not restore a type that was restored' {
        $log = @((New-TestEntry 1), (New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true))
        @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0')).Count | Should -Be 0
    }

    It 'restores a type that was rewritten again after a restore to the original operator' {
        $log = @((New-TestEntry 1), (New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true), (New-TestEntry 3))
        @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0'))[0].OldOperator | Should -Be 'GreaterEquals'
    }

    It 'takes the first rewrite entry and never a restore entry when the first line was lost' {
        $log = @((New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true), (New-TestEntry 3))
        @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0'))[0].OldOperator | Should -Be 'GreaterEquals'
        @(Get-SupersedenceRestoreEntry -Log @(New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true) -Applications @('Old 1.0')).Count | Should -Be 0
    }

    Context 'For one line' {
        BeforeAll {
            $script:LineA = 'APD-Line:a|default|x64'
            $script:LineB = 'APD-Line:b|default|x64'
        }
        It 'leaves the rewrites of another line' {
            $log = @((New-TestEntry 1 -App 'A 1.0' -Line $script:LineA), (New-TestEntry 2 -App 'B 1.0' -Line $script:LineB))
            (@(Get-SupersedenceRestoreEntry -Log $log -Applications @('A 1.0', 'B 1.0') -Line $script:LineA | ForEach-Object { $_.Application }) -join '|') | Should -Be 'A 1.0'
            (@(Get-SupersedenceRestoreEntry -Log $log -Applications @('A 1.0', 'B 1.0') -Line $script:LineB | ForEach-Object { $_.Application }) -join '|') | Should -Be 'B 1.0'
        }
        It 'keeps an entry from before the log carried the line open to every line' {
            $log = @(New-TestEntry 1)
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0') -Line $script:LineA).Count | Should -Be 1
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0') -Line $script:LineB).Count | Should -Be 1
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0')).Count | Should -Be 1
        }
        It 'restores every entry when no line is given' {
            $log = @((New-TestEntry 1 -App 'A 1.0' -Line $script:LineA), (New-TestEntry 2 -App 'B 1.0' -Line $script:LineB))
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('A 1.0', 'B 1.0')).Count | Should -Be 2
        }
        It 'lets a restore of an old entry cancel it in every line' {
            $log = @((New-TestEntry 1), (New-TestEntry 2 -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true))
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0') -Line $script:LineA).Count | Should -Be 0
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('Old 1.0') -Line $script:LineB).Count | Should -Be 0
        }
        It 'lets a rewrite of a line and its restore cancel each other in that line' {
            $log = @((New-TestEntry 1 -App 'A 1.0' -Line $script:LineA), (New-TestEntry 2 -App 'A 1.0' -Line $script:LineA -OldOp 'IsEquals' -NewOp 'GreaterEquals' -Restored $true))
            @(Get-SupersedenceRestoreEntry -Log $log -Applications @('A 1.0') -Line $script:LineA).Count | Should -Be 0
        }
        It 'reads an old log that has no line' {
            $text = (ConvertTo-SupersedenceRewriteLogLine (New-TestEntry 1 -App 'A 1.0' -Line $script:LineA)) + "`n" +
            '{"At":"2026-11-12T10:02:00+00:00","Application":"Old","DeploymentType":"DT","OldOperator":"GreaterEquals","NewOperator":"IsEquals","LogicalName":"S","Restored":false}'
            $entries = @(Read-SupersedenceRewriteLog $text)
            $entries[0].Line | Should -Be $script:LineA
            $entries[1].Line | Should -BeNullOrEmpty
        }
    }
}

Describe 'Writing the rewrite' {
    BeforeAll {
        InModuleScope AppPackagerCommon {
            function script:Set-CMScriptDeploymentType { param($ApplicationName, $DeploymentTypeName, $RemoveDetectionClause, $AddDetectionClause, $ErrorAction) }
            function script:New-CMDetectionClauseFile { param($Path, $FileName, $PropertyType, [switch]$Value, $ExpressionOperator, $ExpectedValue, [switch]$Is64Bit) }
            function script:New-CMDetectionClauseRegistryKeyValue { param($Hive, $KeyName, $ValueName, $PropertyType, [switch]$Value, $ExpressionOperator, $ExpectedValue, [switch]$Is64Bit) }
        }
    }
    AfterAll {
        InModuleScope AppPackagerCommon {
            foreach ($n in 'Set-CMScriptDeploymentType', 'New-CMDetectionClauseFile', 'New-CMDetectionClauseRegistryKeyValue') { Remove-Item -LiteralPath "function:script:$n" -ErrorAction SilentlyContinue }
        }
    }
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Set-Location -ModuleName AppPackagerCommon {}
        Mock New-CMDetectionClauseFile -ModuleName AppPackagerCommon { [pscustomobject]@{ Kind = 'NewClause' } }
        Mock New-CMDetectionClauseRegistryKeyValue -ModuleName AppPackagerCommon { [pscustomobject]@{ Kind = 'NewClause' } }
        Mock Set-CMScriptDeploymentType -ModuleName AppPackagerCommon {}
    }

    It 'removes the clause by its logical name and adds the same clause with the other operator, built off the site drive' {
        InModuleScope AppPackagerCommon {
            $clause = [pscustomobject]@{ Kind = 'File'; LogicalName = 'File_1'; Operator = 'GreaterEquals'; ExpectedValue = '1.0.0'; Path = 'C:\Vendor'; FileName = 'app.exe'; Is64Bit = $false; Hive = ''; Key = ''; ValueName = '' }
            Write-SupersedenceClauseOperator -SiteCode 'MCM' -ApplicationName 'Old 1.0' -DeploymentTypeName 'Old 1.0 - x64' -Clause $clause -NewOperator 'IsEquals'
        }
        Should -Invoke New-CMDetectionClauseFile -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Path -eq 'C:\Vendor' -and $FileName -eq 'app.exe' -and $PropertyType -eq 'Version' -and $ExpressionOperator -eq 'IsEquals' -and $ExpectedValue -eq '1.0.0' -and $Value
        }
        Should -Invoke Set-CMScriptDeploymentType -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $ApplicationName -eq 'Old 1.0' -and $DeploymentTypeName -eq 'Old 1.0 - x64' -and $RemoveDetectionClause -eq 'File_1' -and $null -ne $AddDetectionClause
        }
        Should -Invoke Set-Location -ModuleName AppPackagerCommon -ParameterFilter { $Path -eq 'C:' }
        Should -Invoke Set-Location -ModuleName AppPackagerCommon -ParameterFilter { $Path -eq 'MCM:' }
    }

    It 'keeps the hive, the key, the value and the view of a registry clause' {
        InModuleScope AppPackagerCommon {
            $clause = [pscustomobject]@{ Kind = 'Registry'; LogicalName = 'RegSetting_1'; Operator = 'GreaterEquals'; ExpectedValue = '1.0.0'; Path = ''; FileName = ''; Is64Bit = $true; Hive = 'HKEY_LOCAL_MACHINE'; Key = 'SOFTWARE\APDLab'; ValueName = 'DisplayVersion' }
            Write-SupersedenceClauseOperator -SiteCode 'MCM' -ApplicationName 'Reg' -DeploymentTypeName 'Reg - x64' -Clause $clause -NewOperator 'IsEquals'
        }
        Should -Invoke New-CMDetectionClauseRegistryKeyValue -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
            $Hive -eq 'LocalMachine' -and $KeyName -eq 'SOFTWARE\APDLab' -and $ValueName -eq 'DisplayVersion' -and $PropertyType -eq 'Version' -and $ExpressionOperator -eq 'IsEquals' -and $ExpectedValue -eq '1.0.0' -and $Is64Bit
        }
    }

    It 'refuses an operator other than the two' {
        { InModuleScope AppPackagerCommon { Write-SupersedenceClauseOperator -SiteCode 'MCM' -ApplicationName 'A' -DeploymentTypeName 'D' -Clause ([pscustomobject]@{}) -NewOperator 'NotEquals' } } | Should -Throw
    }

    It 'writes the intent to the log before the site changes, and a restore entry when the write fails' {
        $log = Join-Path $TestDrive 'rewrite.jsonl'
        Mock Set-CMScriptDeploymentType -ModuleName AppPackagerCommon { throw 'access denied' }
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause); NotRewritableReason = $null })
        $result = InModuleScope AppPackagerCommon -Parameters @{ Plan = $plan; Log = $log } { param($Plan, $Log) Invoke-SupersedenceRewrite -Plan $Plan -SiteCode 'MCM' -Marker 'APD-Line:x|default|x64' -LogPath $Log }
        @($result.Written).Count | Should -Be 0
        @($result.Failures)[0] | Should -Match 'access denied'
        $entries = @(Read-SupersedenceRewriteLog (Get-Content -LiteralPath $log -Raw))
        $entries.Count | Should -Be 2
        $entries[0].Restored | Should -BeFalse
        $entries[1].Restored | Should -BeTrue
        $entries[1].NewOperator | Should -Be 'GreaterEquals'
        $entries[0].Line | Should -Be 'APD-Line:x|default|x64'
    }

    It 'keeps the intent and reports the write when the site accepts it' {
        $log = Join-Path $TestDrive 'rewrite2.jsonl'
        $plan = New-TestPlan (New-TestApp 'Old 1.0' '1.0.0' @('Old 1.0 - x64')) ([pscustomobject]@{ Clause = (New-TestFileClause); NotRewritableReason = $null })
        $result = InModuleScope AppPackagerCommon -Parameters @{ Plan = $plan; Log = $log } { param($Plan, $Log) Invoke-SupersedenceRewrite -Plan $Plan -SiteCode 'MCM' -Marker 'APD-Line:x|default|x64' -LogPath $Log }
        @($result.Written).Count | Should -Be 1
        @($result.Failures).Count | Should -Be 0
        @(Read-SupersedenceRewriteLog (Get-Content -LiteralPath $log -Raw)).Count | Should -Be 1
    }
}
