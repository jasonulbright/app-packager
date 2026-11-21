# Window proof of the supersedence confirmation. Runs in the interactive session of a lab account.
# It needs two ConfigMgr applications of each proof title at version 1.0.0 on the site and the share \\cm01.contoso.com\Q92Proof$,
# which the orchestrating script of the proof makes and removes. It publishes only the proof applications.
# Three Package runs: Yes with Remember this answer, Skip for each application, and No in the light theme.
$dataDir = Join-Path $env:LOCALAPPDATA 'AppPackager'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
$dl = Join-Path $GuideRoot 'dl'
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
$q = 'C:\guide\q92'
New-Item -ItemType Directory $q, (Join-Path $dl '_cache') -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $dl '_cache\gfixture-setup.exe'), 'not a real installer')
$versionFile = Join-Path $q 'version.txt'

# The two proof packagers. They stage a file and read the version from the version file; the publish is the real one.
$template = @'
<#
Vendor: Q92 Proof Vendor
App: @TITLE@
CMName: @TITLE@
VendorUrl: https://contoso.invalid
CPE: cpe:2.3:a:contoso:@KEY@:*:*:*:*:*:*:*
ReleaseNotesUrl: https://contoso.invalid/notes
DownloadPageUrl: https://contoso.invalid/download
UpdateCadenceDays: 30
.SYNOPSIS
    Offline proof packager of the supersedence confirmation.
#>
param(
    [string]$SiteCode = "MCM", [string]$Comment = "", [string]$FileServerPath = "\\fileserver\sccm$", [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15, [int]$MaximumRuntimeMins = 30, [string]$ContentLayout = 'Nested', [string]$LogPath,
    [switch]$GetLatestVersionOnly, [switch]$StageOnly, [switch]$PackageOnly, [switch]$VerboseLog
)
Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog
$Version = (Get-Content -LiteralPath 'C:\guide\q92\version.txt' -Raw).Trim()
if ($GetLatestVersionOnly) { Write-Output $Version; exit 0 }
$InstallerFileName = "gfixture-setup.exe"
$BaseDownloadRoot = Join-Path $DownloadRoot "@KEY@"
function Invoke-StageApp {
    $contentPath = Join-Path $BaseDownloadRoot $Version
    Initialize-Folder -Path $contentPath
    Copy-Item -LiteralPath (Join-Path (Get-SharedDownloadCacheRoot -DownloadRoot $DownloadRoot) $InstallerFileName) -Destination (Join-Path $contentPath $InstallerFileName) -Force
    Write-ContentWrappers -OutputPath $contentPath -InstallPs1Content 'exit 0' -UninstallPs1Content 'exit 0'
    Write-StageManifest -Path (Join-Path $contentPath "stage-manifest.json") -ManifestData @{
        AppName = "@TITLE@"; Publisher = "Q92 Proof Vendor"; SoftwareVersion = $Version; InstallerFile = $InstallerFileName; InstallerType = "EXE"
        InstallArgs = "/S"; Architecture = "x64"; InstallationBehaviorType = "InstallForSystem"; LogonRequirementType = "WhetherOrNotUserLoggedOn"
        Detection = @{ Type = "RegistryKeyValue"; RegistryKeyRelative = "SOFTWARE\@KEY@"; ValueName = "DisplayVersion"; ExpectedValue = $Version; Operator = "IsEquals"; Is64Bit = $true }
    }
    Write-Log "Stage complete               : $contentPath"
}
function Invoke-PackageApp {
    $localContentPath = Join-Path $BaseDownloadRoot $Version
    $manifest = Read-StageManifest -Path (Join-Path $localContentPath "stage-manifest.json")
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) { throw "Network root path not accessible: $FileServerPath" }
    $networkContentPath = Join-Path (Join-Path $FileServerPath '@KEY@') $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath
    Write-Log "Network content path         : $networkContentPath"
    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest
    New-MECMApplicationFromManifest -Manifest $manifest -SiteCode $SiteCode -Comment $Comment -NetworkContentPath $networkContentPath -EstimatedRuntimeMins $EstimatedRuntimeMins -MaximumRuntimeMins $MaximumRuntimeMins
}
try {
    if ($StageOnly) { Invoke-StageApp; exit 0 }
    if ($PackageOnly) { Invoke-PackageApp | Out-Null; exit 0 }
    Invoke-StageApp; Invoke-PackageApp | Out-Null; exit 0
}
catch { Write-Log ("Failed: " + $_.Exception.Message) -Level ERROR; exit 1 }
'@
foreach ($p in @(@{ Title = 'Q92 Proof App'; Key = 'Q92ProofA' }, @{ Title = 'Q92 Proof Tool'; Key = 'Q92ProofB' })) {
    $text = $template.Replace('@TITLE@', $p.Title).Replace('@KEY@', $p.Key)
    [IO.File]::WriteAllText((Join-Path $appDir ('Packagers\package-' + $p.Key.ToLowerInvariant() + '.ps1')), $text, (New-Object System.Text.ASCIIEncoding))
}
# The window shows the proof packagers only; the preferences of the lab account are removed at the end.
Set-Content $versionFile '2.0.0'

$win = Start-App 'start-apppackager.ps1' 'AppPackager*v20??.??.??.????'
Size-Win $win 1440 900
$win = Find-Top 'AppPackager*v20??.??.??.????' 10
$winCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
function WinList { (@($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -notlike 'GlowWindow*' } | ForEach-Object { $_.Current.Name }) -join ' ; ' }
function Top-Dialog {
    @($root.FindAll('Descendants', $winCond)) | Where-Object { $_.Current.Name -notlike 'AppPackager*v20??.??.??.????' -and $_.Current.Name -notin 'Program Manager', '' -and $_.Current.Name -notlike 'Administrator*' -and $_.Current.Name -notlike 'GlowWindow*' -and $_.Current.Name -notlike '*Edge*' } | Select-Object -First 1
}
function Idle([int]$sec = 300) {
    Start-Sleep -Seconds 2
    if (-not (Wait-Until { $b = Get-El $win 'btnCheckLatest' -Quiet; $b -and $b.Current.IsEnabled } $sec)) { throw "Timeout: the app stayed busy for $sec seconds." }
    Start-Sleep -Seconds 2
}
function Status { Get-ElText (Get-El $win 'txtStatus') }
function Check-Row([int]$index) {
    $rows = @((Get-El $win 'dataGrid').FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem))))
    if ($rows.Count -le $index) { throw 'The grid has no row ' + $index }
    $box = $rows[$index].FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::CheckBox)))
    $toggle = $box.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    if ($toggle.Current.ToggleState -ne 'On') { $toggle.Toggle() }
    Start-Sleep -Milliseconds 400
    if ($toggle.Current.ToggleState -ne 'On') { throw 'The row is not checked.' }
}
# The dialog of one application: its text starts with the title of that application. A dialog that is still closing never matches.
function Wait-Supersedence([string]$Like = '*', [int]$sec = 180) {
    $found = $null
    if (-not (Wait-Until { $t = Top-Dialog; if ($t -and $t.Current.Name -eq 'Supersedence') { $i = Get-El $t 'txtIntro' -Quiet; if ($i -and (Get-ElText $i) -like $Like) { $script:found = $t; return $true } }; return $false } $sec)) { throw ('No supersedence dialog like ' + $Like + '. Windows: ' + (WinList) + ' Status: ' + (Status)) }
    $d = $script:found
    Start-Sleep -Milliseconds 800
    $d
}

Section 'setup' {
    $s = Find-Last 'Setup' 20
    Set-Text (Get-El $s 'cm.siteCode') $Capture.SiteCode; Set-Text (Get-El $s 'cm.provider') $Capture.SMSProvider
    Set-Text (Get-El $s 'cm.fileShare') '\\cm01.contoso.com\Q92Proof$'; Set-Text (Get-El $s 'pkg.downloadRoot') $dl
    Invoke-El (Get-El $s 'btnWizSave'); Start-Sleep -Seconds 3
    Set-Text (Get-El $win 'txtGridFilter') 'Q92 Proof'; Start-Sleep -Seconds 2
    Log ('grid status: ' + (Status))
}
Section 'options: the setting by application' {
    Invoke-El (Get-El $win 'btnOptions'); $o = Find-LastLike 'Options*' 15
    Size-Win $o 1400 1040
    $tree = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem))))
    $group = $tree | Where-Object { $_.Current.Name -eq 'Applications' } | Select-Object -First 1
    if ($group) { try { Expand-El $group } catch { $null = $_ }; Start-Sleep -Milliseconds 600 }
    $node = @($o.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem)))) | Where-Object { $_.Current.Name -eq 'All Applications' } | Select-Object -First 1
    if (-not $node) { throw 'The Options tree has no All Applications page.' }
    Select-El $node; Start-Sleep -Milliseconds 1200
    # The grid is virtualized: each proof row is realized through the item container pattern, then its check box is toggled.
    $grid = Get-El $o 'sup.apps'
    foreach ($name in 'Q92 Proof App', 'Q92 Proof Tool') {
        $item = $grid.GetCurrentPattern([System.Windows.Automation.ItemContainerPattern]::Pattern).FindItemByProperty($null, $AE::NameProperty, $name)
        if (-not $item) { throw "No row for $name in the grid of all applications." }
        try { $item.GetCurrentPattern([System.Windows.Automation.VirtualizedItemPattern]::Pattern).Realize() } catch { $null = $_ }
        Start-Sleep -Milliseconds 700
        $row = $grid.FindFirst('Descendants', (New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))))
        Toggle-El (Get-El $row 'appgrid.check')
    }
    # The default value of the bulk bar is the column Supersedence with the value Replace.
    Invoke-El (Get-El $o 'sup.apps.bulk.Checked'); Start-Sleep -Milliseconds 800
    Shot $o 'proof-all-applications'
    Invoke-El (Get-El $o 'btnOK'); Start-Sleep -Seconds 3
    Log ('options saved: ' + (Get-Content (Join-Path $appDir 'AppPackager.preferences.json') -Raw | Select-String 'catalog:package-q92proof[ab]' -AllMatches | ForEach-Object { $_.Matches.Count }))
}
# One Package run: both rows checked, Stage, then Publish to ConfigMgr.
function Package-Run([string]$Version, [scriptblock]$Answers) {
    Set-Content $versionFile $Version
    Set-Text (Get-El $win 'txtGridFilter') 'Q92 Proof'; Start-Sleep -Seconds 2
    Check-Row 0; Check-Row 1
    Invoke-El (Get-El $win 'btnStage'); Start-Sleep -Seconds 4; Idle 600; Log ("stage $Version status: " + (Status))
    Invoke-El (Get-El $win 'btnPackage'); Start-Sleep -Seconds 3
    & $Answers
    try { Idle 420 } catch { Log ("STUCK after publish $Version. Windows: " + (WinList) + " Status: " + (Status)); Shot-Screen ("fail-stuck-" + $Version); throw }
    Log ("publish $Version status: " + (Status))
}
Section 'run 1: Yes with Remember this answer' {
    Package-Run '2.0.0' {
        $d = Wait-Supersedence 'Q92 Proof App*'; Shot $d 'supersedence-confirm'; Dump-Tree $d 'supersedence-confirm'
        Toggle-El (Get-El $d 'chkRemember'); Invoke-El (Get-El $d 'Yes'); Start-Sleep -Seconds 2
        # The second application is not asked: the answer holds for the run. A dialog here is a failure.
        Start-Sleep -Seconds 8
        $again = Top-Dialog; $againText = $(if ($again -and $again.Current.Name -eq 'Supersedence') { $ti = Get-El $again 'txtIntro' -Quiet; if ($ti) { Get-ElText $ti } else { '' } } else { '' })
        if ($againText -like 'Q92 Proof Tool*') { Log ('ERROR: the second application was asked although Remember was checked. Windows: ' + (WinList)); Shot-Screen 'fail-second-dialog' }
    }
}
Section 'run 2: Skip for each application' {
    Package-Run '3.0.0' {
        foreach ($like in 'Q92 Proof App*', 'Q92 Proof Tool*') { $d = Wait-Supersedence $like; Log ("dialog of run 2: $like"); Invoke-El (Get-El $d 'Skip'); Start-Sleep -Seconds 3 }
    }
}
Section 'run 3: No in the light theme' {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 2
    Shot $win 'light-supersedence-main'
    Package-Run '4.0.0' {
        $d = Wait-Supersedence 'Q92 Proof App*'; Shot $d 'light-supersedence-confirm'; Invoke-El (Get-El $d 'No'); Start-Sleep -Seconds 2
    }
    Shot $win 'light-supersedence-run-stopped'
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
}
Section 'the sidebar has no rewrite, restore or apply button' {
    foreach ($id in 'btnApplySupersedence', 'btnRewriteOld', 'btnRestoreOld') { if (Get-El $win $id -Quiet) { Log "ERROR: $id is still in the window" } }
    foreach ($name in 'Apply Supersedence', 'Rewrite Old Versions', 'Restore Old Versions') { if (Get-El $win $name -Quiet) { Log "ERROR: $name is still in the window" } }
    Log 'sidebar checked'
}
Get-ChildItem -LiteralPath (Join-Path $appDir 'Logs') -Filter *.log -File -ErrorAction SilentlyContinue | Copy-Item -Destination $OutDir -Force
Stop-App 'start-apppackager.ps1'
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $dl -Recurse -Force -ErrorAction SilentlyContinue
