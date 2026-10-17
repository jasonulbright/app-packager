#Requires -Version 5.1
# Headless probe for Discover unmanaged software: the prompt window and its
# four buttons, the choice applied to the preferences, the finish step for a
# failed, an empty and a found result, the background runspace start, time
# limit and stop, and the SQL override fields in Options, ConfigMgr
# Preferences. No SQL server is contacted and no window is shown.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path (Join-Path $root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($name in @('Show-DiscoveryDialog', 'Get-DiscoveryPromptText', 'Invoke-DiscoveryChoice', 'Enable-DiscoveredApplication',
                    'Add-DiscoveryDismissed', 'Resolve-DiscoveryButton', 'Complete-DiscoveryScan', 'Read-Preferences', 'Resolve-FirstRunCompleted', 'New-MecmPreferencesPanel',
                    'Invoke-DiscoveryScan', 'Start-DiscoveryRunspace', 'Get-DiscoveryRunspaceState', 'Stop-DiscoveryRunspace')) {
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
    if (-not $fn) { throw ('function not found: ' + $name) }
    # Global, the way the shell publishes its functions: a GetNewClosure
    # handler runs in a module whose scope chain ends at global.
    Set-Item -Path ('function:global:' + $name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
}

$fixtureDir = Join-Path $env:TEMP ('ap-discover-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $fixtureDir -Force)
$global:fixturePath = Join-Path $fixtureDir 'preferences.json'
function global:Get-PreferencesPath { $global:fixturePath }
function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
function global:Add-LogLine { param([string]$Message) $global:Log += $Message }
function global:Save-Preferences { param($Prefs) $global:Saves++ }
function global:Invoke-RefreshGrid { param([switch]$DiscardSiteResults) $global:Refreshes++ }
function global:Set-ActionButtonsEnabled { param([bool]$Enabled) $global:ButtonsEnabled = $Enabled }
foreach ($stub in 'Get-GitHubApiAuthStatus', 'Get-IconPackManifestPath', 'Get-IconPackStatusText', 'Get-IntuneWinToolCachePath', 'Install-IconPack',
                  'Install-IconPackFromFile', 'Install-IntuneWinAppUtil', 'Invoke-DetectIntuneWinAppUtil', 'Read-IconPackManifest') {
    Set-Item -Path ('function:global:' + $stub) -Value { }
}

$checks = 0
$failures = @()
$assert = {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { $script:checks += 1; Write-Host ('PASS  ' + $Message) }
    else { $script:failures += $Message; Write-Host ('FAIL  ' + $Message) }
}
$click = { param($Control) $Control.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }

$candidates = @(
    [pscustomobject]@{ Script = 'package-git.ps1'; CMName = 'Git'; Vendor = 'The Git Development Community'; ClientCount = 2; MatchedName = 'Git' }
    [pscustomobject]@{ Script = 'package-7zip.ps1'; CMName = '7-Zip'; Vendor = '7-Zip'; ClientCount = 4; MatchedName = '7-Zip 26.03 (x64 edition)' }
    [pscustomobject]@{ Script = 'package-hashtools.ps1'; CMName = 'HashTools'; Vendor = 'Binary Fortress Software'; ClientCount = 1; MatchedName = 'HashTools' }
)
$names = @{ 'package-git.ps1' = 'Git for Windows (x64)'; 'package-7zip.ps1' = '7-Zip' }

try {
    # --- The prompt window ---------------------------------------------------
    $result = Show-DiscoveryDialog -Owner $null -Candidates $candidates -StatusText 'Inventory from 4 clients, newest 2026-10-17.' -NameByScript $names -Probe {
        param($p)
        & $assert ($p.Prompt.Text -eq 'Discovered 3 apps. Enable console view?') 'the prompt reads Discovered 3 apps. Enable console view?'
        & $assert ($p.Status.Text -eq 'Inventory from 4 clients, newest 2026-10-17.') 'the status line shows the inventory text'
        $rows = @($p.Rows)
        & $assert ($rows.Count -eq 3) 'one row per candidate'
        & $assert ($rows[0].Script -eq 'package-7zip.ps1' -and $rows[0].ClientCount -eq 4) 'the app on the most clients comes first'
        & $assert ($rows[1].Name -eq 'Git for Windows (x64)' -and $rows[2].Name -eq 'HashTools') 'the name is the grid name, or the ConfigMgr name without one'
        & $assert ($rows[1].Vendor -eq 'The Git Development Community') 'the vendor column carries the vendor'
        & $assert (@($rows | Where-Object Selected).Count -eq 0) 'no row starts checked'
        & $assert (@($p.Grid.Columns | ForEach-Object { [string]$_.Header }) -join ',' -eq ',Application,Vendor,Clients,Inventory name') 'the columns are check, name, vendor, client count and inventory name'
        & $assert ($p.EnableSelected.Content -eq 'Enable selected' -and $p.EnableAll.Content -eq 'Enable all' -and $p.NotNow.Content -eq 'Not now' -and $p.Dismiss.Content -eq "Don't ask again for selected") 'the four buttons carry their labels'
        & $click $p.EnableSelected
        & $assert ($p.Note.Visibility -eq 'Visible' -and $p.State.Action -eq 'NotNow') 'Enable selected with nothing checked asks for a check and keeps the window'
        $rows[2].Selected = $true
        & $click $p.EnableSelected
    }
    & $assert ($result.Action -eq 'EnableSelected' -and (@($result.Scripts) -join ',') -eq 'package-hashtools.ps1') 'Enable selected returns the checked script'

    $result = Show-DiscoveryDialog -Owner $null -Candidates $candidates -Probe { param($p) & $click $p.EnableAll }
    & $assert ($result.Action -eq 'EnableAll' -and @($result.Scripts).Count -eq 3) 'Enable all returns every script'

    $result = Show-DiscoveryDialog -Owner $null -Candidates $candidates -Probe { param($p) & $click $p.NotNow }
    & $assert ($result.Action -eq 'NotNow' -and @($result.Scripts).Count -eq 0) 'Not now returns no script'

    $result = Show-DiscoveryDialog -Owner $null -Candidates $candidates -Probe {
        param($p)
        & $click $p.Dismiss
        & $assert ($p.Note.Visibility -eq 'Visible' -and $p.Note.Text -eq 'Check the apps to leave out of later discoveries. Nothing changed.' -and $p.State.Action -eq 'NotNow') "Don't ask again with nothing checked says so and keeps the window"
        & $click $p.NotNow
    }
    & $assert ($result.Action -eq 'NotNow' -and @($result.Scripts).Count -eq 0) "Don't ask again with nothing checked returns no script"

    $result = Show-DiscoveryDialog -Owner $null -Candidates $candidates -Probe {
        param($p)
        @($p.Rows | Where-Object { $_.Script -in 'package-7zip.ps1', 'package-hashtools.ps1' }) | ForEach-Object { $_.Selected = $true }
        & $click $p.Dismiss
    }
    & $assert ($result.Action -eq 'Dismiss' -and (@($result.Scripts | Sort-Object) -join ',') -eq 'package-7zip.ps1,package-hashtools.ps1') "Don't ask again returns only the checked scripts"

    $result = Show-DiscoveryDialog -Owner $null -Candidates @($candidates[0]) -Probe { param($p) & $assert ($p.Prompt.Text -eq 'Discovered 1 app. Enable console view?') 'one candidate reads Discovered 1 app' }

    # --- The choice applied ----------------------------------------------------
    $global:txtStatus = [pscustomobject]@{ Text = '' }
    $global:window = [pscustomobject]@{ Cursor = $null }
    $script:Prefs = [pscustomobject]@{ HiddenApplications = @('package-git.ps1', 'package-7zip.ps1', 'package-vlc.ps1'); DiscoveryDismissed = @() }
    $global:Saves = 0; $global:Refreshes = 0; $global:Log = @()
    Invoke-DiscoveryChoice -Choice ([pscustomobject]@{ Action = 'EnableAll'; Scripts = @('package-git.ps1', 'package-7zip.ps1') })
    & $assert ((@($script:Prefs.HiddenApplications) -join ',') -eq 'package-vlc.ps1') 'enabling removes the scripts from HiddenApplications'
    & $assert ($global:Saves -eq 1 -and $global:Refreshes -eq 1) 'enabling saves the preferences and refreshes the grid once'
    & $assert (@($global:Log | Where-Object { $_ -like '*2 app(s) added to the grid*' }).Count -eq 1) 'the log names the count added'

    $global:Saves = 0; $global:Refreshes = 0
    Invoke-DiscoveryChoice -Choice ([pscustomobject]@{ Action = 'Dismiss'; Scripts = @('package-vlc.ps1') })
    & $assert ((@($script:Prefs.DiscoveryDismissed) -join ',') -eq 'package-vlc.ps1') 'dismissing records the scripts'
    & $assert ((@($script:Prefs.HiddenApplications) -join ',') -eq 'package-vlc.ps1') 'dismissing keeps the app hidden'
    & $assert ($global:Saves -eq 1 -and $global:Refreshes -eq 0) 'dismissing saves without a grid refresh'

    $global:Saves = 0
    Invoke-DiscoveryChoice -Choice ([pscustomobject]@{ Action = 'NotNow'; Scripts = @() })
    & $assert ($global:Saves -eq 0) 'Not now saves nothing'

    # --- Don't ask again for selected, then the next scan ------------------------------
    Import-Module (Join-Path $root 'Packagers\AppPackagerDiscovery.psm1') -Force -DisableNameChecking
    $scanPackagers = @(
        [pscustomobject]@{ Script = 'package-git.ps1'; CMName = 'Git'; Vendor = 'Git'; ArpName = 'Git' }
        [pscustomobject]@{ Script = 'package-7zip.ps1'; CMName = '7-Zip'; Vendor = '7-Zip'; ArpName = '7-Zip*' }
        [pscustomobject]@{ Script = 'package-hashtools.ps1'; CMName = 'HashTools'; Vendor = 'Binary Fortress Software'; ArpName = 'HashTools' }
    )
    $snapshot = [pscustomobject]@{ Applications = @(); Rows = @(
            [pscustomobject]@{ DisplayName = 'Git'; Version = '2.51.0'; Architecture = 'x64'; ClientCount = 2 }
            [pscustomobject]@{ DisplayName = '7-Zip 26.03 (x64 edition)'; Version = '26.03'; Architecture = 'x64'; ClientCount = 4 }
            [pscustomobject]@{ DisplayName = 'HashTools'; Version = '5.1'; Architecture = 'x64'; ClientCount = 1 }
        ) }
    $script:Prefs = [pscustomobject]@{ HiddenApplications = @('package-git.ps1', 'package-7zip.ps1', 'package-hashtools.ps1'); DiscoveryDismissed = @(); DiscoveryEnabled = @() }
    $offered = @(Find-DiscoveryCandidate -Snapshot $snapshot -Packager $scanPackagers -HiddenScript $script:Prefs.HiddenApplications -DismissedScript $script:Prefs.DiscoveryDismissed)
    & $assert ($offered.Count -eq 3) 'the first scan offers the three hidden apps'
    $choice = Show-DiscoveryDialog -Owner $null -Candidates $offered -Probe {
        param($p)
        @($p.Rows | Where-Object { $_.Script -in 'package-7zip.ps1', 'package-hashtools.ps1' }) | ForEach-Object { $_.Selected = $true }
        & $click $p.Dismiss
    }
    $global:Saves = 0
    Invoke-DiscoveryChoice -Choice $choice
    & $assert ((@($script:Prefs.DiscoveryDismissed | Sort-Object) -join ',') -eq 'package-7zip.ps1,package-hashtools.ps1' -and $global:Saves -eq 1) 'only the checked apps land in DiscoveryDismissed'
    & $assert (@($script:Prefs.DiscoveryEnabled).Count -eq 0 -and @($script:Prefs.HiddenApplications).Count -eq 3) 'dismissing enables nothing and keeps the apps hidden'
    $offered = @(Find-DiscoveryCandidate -Snapshot $snapshot -Packager $scanPackagers -HiddenScript $script:Prefs.HiddenApplications -DismissedScript $script:Prefs.DiscoveryDismissed)
    & $assert ((@($offered | ForEach-Object { $_.Script }) -join ',') -eq 'package-git.ps1') 'the next scan offers the unchecked app again'

    $global:Saves = 0
    $choice = Show-DiscoveryDialog -Owner $null -Candidates $offered -Probe { param($p) & $click $p.Dismiss; & $click $p.NotNow }
    Invoke-DiscoveryChoice -Choice $choice
    & $assert ($global:Saves -eq 0 -and @($script:Prefs.DiscoveryDismissed).Count -eq 2) "Don't ask again with nothing checked changes no preference"
    $offered = @(Find-DiscoveryCandidate -Snapshot $snapshot -Packager $scanPackagers -HiddenScript $script:Prefs.HiddenApplications -DismissedScript $script:Prefs.DiscoveryDismissed)
    & $assert ($offered.Count -eq 1) 'after nothing checked the next scan offers the same app'

    # --- The finish step ---------------------------------------------------------
    $global:Shown = 0
    function global:Show-DiscoveryDialog { param($Owner, $Candidates, $StatusText, $NameByScript, $Probe) $global:Shown++; [pscustomobject]@{ Action = 'EnableAll'; Scripts = @($Candidates | ForEach-Object { $_.Script }) } }
    $global:DiscoveryNames = @{}
    $global:Log = @(); $global:ButtonsEnabled = $false
    Complete-DiscoveryScan -Result ([pscustomobject]@{ Ok = $false; Error = 'Login failed for the ConfigMgr site database on cm01.'; Candidates = @() })
    & $assert ($global:txtStatus.Text -eq 'Discover unmanaged software: Login failed for the ConfigMgr site database on cm01.') 'a SQL failure shows the plain message on the status line'
    & $assert (@($global:Log | Where-Object { $_ -like '*Login failed*' }).Count -eq 1) 'a SQL failure is logged'
    & $assert ($global:ButtonsEnabled -and $global:Shown -eq 0) 'a SQL failure releases the buttons and shows no prompt'

    Complete-DiscoveryScan -Result ([pscustomobject]@{ Ok = $true; Server = 'cm01'; Database = 'CM_MCM'; StatusText = 'No inventory data. Hardware inventory has no Add/Remove Programs rows.'; Candidates = @() })
    & $assert ($global:txtStatus.Text -like 'No inventory data*' -and $global:Shown -eq 0) 'no candidates puts the status text on the status line without a prompt'

    $script:Prefs = [pscustomobject]@{ HiddenApplications = @('package-git.ps1'); DiscoveryDismissed = @() }
    Complete-DiscoveryScan -Result ([pscustomobject]@{ Ok = $true; Server = 'cm01'; Database = 'CM_MCM'; StatusText = '1 app found.'; Candidates = @($candidates[0]) })
    & $assert ($global:Shown -eq 1 -and @($script:Prefs.HiddenApplications).Count -eq 0) 'a candidate opens the prompt and its choice is applied'

    # --- The background runspace ---------------------------------------------------
    # Without a site code and without both overrides the scan stops before any
    # SQL connection, so the real module runs end to end with no network.
    $modulePath = Join-Path $root 'Packagers\AppPackagerDiscovery.psm1'
    $scanText = [string]${function:Invoke-DiscoveryScan}
    $job = Start-DiscoveryRunspace -ModulePath $modulePath -ScanText $scanText -ScanArgs @{ SiteCode = ''; ServerOverride = 'sql01'; DatabaseOverride = ''; Packager = @(); HiddenScript = @(); DismissedScript = @() }
    $deadline = [DateTime]::UtcNow.AddSeconds(60)
    while ((Get-DiscoveryRunspaceState -Job $job -TimeoutSeconds 60) -eq 'Running' -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    & $assert ((Get-DiscoveryRunspaceState -Job $job -TimeoutSeconds 60) -eq 'Done') 'the runspace scan finishes'
    & $assert ($job.State.Result -and -not $job.State.Result.Ok -and [string]$job.State.Result.Error -like '*site code or SQL overrides are required*') 'the runspace scan imports the module and returns the plain target message'
    $runspace = $job.Runspace
    Stop-DiscoveryRunspace -Job $job
    & $assert ($null -eq $job.Shell -and $null -eq $job.Runspace) 'stopping a finished scan releases the PowerShell object and the runspace'
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ($runspace.RunspaceStateInfo.State -notin 'Closed', 'Broken' -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    & $assert ($runspace.RunspaceStateInfo.State -eq 'Closed') 'the finished runspace closes'

    $job = Start-DiscoveryRunspace -ModulePath $modulePath -ScanText 'param() Start-Sleep -Seconds 120' -ScanArgs @{}
    Start-Sleep -Milliseconds 500
    & $assert ((Get-DiscoveryRunspaceState -Job $job -TimeoutSeconds 600) -eq 'Running') 'a scan within the time limit is running'
    & $assert ((Get-DiscoveryRunspaceState -Job $job -TimeoutSeconds 600 -UtcNow ([DateTime]::UtcNow.AddSeconds(601))) -eq 'TimedOut') 'a scan past the time limit is timed out'
    $runspace = $job.Runspace
    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    Stop-DiscoveryRunspace -Job $job
    & $assert ($stopWatch.Elapsed.TotalSeconds -lt 10 -and $null -eq $job.Shell -and $null -eq $job.Runspace) 'stopping a running scan returns without waiting for it'
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ($runspace.RunspaceStateInfo.State -notin 'Closed', 'Broken' -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    & $assert ($runspace.RunspaceStateInfo.State -eq 'Closed') 'the stopped runspace closes'
    & $assert ((Get-DiscoveryRunspaceState -Job $null) -eq 'Done') 'no scan reads as done'
    Stop-DiscoveryRunspace -Job $null

    # --- Options, ConfigMgr Preferences ---------------------------------------------
    Set-Content -LiteralPath $global:fixturePath -Value '{ "DiscoverySqlServer": "sql01", "DiscoverySqlDatabase": "CM_ABC", "DiscoveryDismissed": ["package-git.ps1", "package-vlc.ps1"] }'
    $script:Prefs = Read-Preferences
    $panel = New-MecmPreferencesPanel
    $el = $panel.Element
    & $assert ($el.FindName('txtDiscoverySqlServer').Text -eq 'sql01' -and $el.FindName('txtDiscoverySqlDatabase').Text -eq 'CM_ABC') 'the panel shows the stored SQL overrides'
    & $assert ($el.FindName('txtDiscoveryDismissed').Text -eq '2 app(s)' -and $el.FindName('btnClearDiscoveryDismissed').IsEnabled) 'the panel counts the apps not offered'
    $el.FindName('txtDiscoverySqlServer').Text = ' cm01.contoso.com '
    $el.FindName('txtDiscoverySqlDatabase').Text = ''
    & $click $el.FindName('btnClearDiscoveryDismissed')
    & $panel.Commit
    & $assert ($script:Prefs.DiscoverySqlServer -eq 'cm01.contoso.com' -and $script:Prefs.DiscoverySqlDatabase -eq '') 'OK saves the trimmed server and an empty database'
    & $assert (@($script:Prefs.DiscoveryDismissed).Count -eq 0) 'Clear list empties the list on OK'
}
finally {
    Remove-Item -LiteralPath $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("`n{0} passed, {1} failed." -f $checks, $failures.Count)
exit ([int]($failures.Count -gt 0))
