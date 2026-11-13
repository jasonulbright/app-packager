# Shared set-up of the Options window probes. Dot-source it from a probe that runs in an STA process.
#
# It loads the WPF assemblies and the modules, publishes every function of start-apppackager.ps1 to the
# global scope the way the shell does for its closure handlers, and points the files that the Options
# window reads and writes at a fixture folder, so that a probe never touches the repository.
param(
    [Parameter(Mandatory)][string]$Root,
    [Parameter(Mandatory)][string]$FixtureDir
)

$ErrorActionPreference = 'Stop'
$global:SharedStyleRootOverride = $Root
$script:SharedStyleRootOverride = $Root

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -Path (Join-Path $Root 'Lib\ControlzEx.dll')
Add-Type -Path (Join-Path $Root 'Lib\MahApps.Metro.dll')
Import-Module (Join-Path $Root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking -Global
foreach ($module in 'AppPackagerCommon', 'AppPackagerDiscovery', 'AppPackagerWsus', 'AppPackagerSigning') {
    Import-Module (Join-Path $Root ('Packagers\{0}.psd1' -f $module)) -Force -Global -DisableNameChecking -ErrorAction SilentlyContinue
}

$script:LauncherAst = $null
$tokens = $null
$errors = $null
$script:LauncherAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $Root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
foreach ($fn in $script:LauncherAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    $body = $fn.Body.Extent.Text
    Set-Item -Path ('function:global:' + $fn.Name) -Value ([scriptblock]::Create($body.Substring(1, $body.Length - 2)))
}

$global:PackagersRoot = Join-Path $Root 'Packagers'
$global:FixtureDir = $FixtureDir
$global:FixturePrefsPath = Join-Path $FixtureDir 'preferences.json'
function global:Get-PreferencesPath { $global:FixturePrefsPath }
function global:Get-CwaSwitchesPath { Join-Path $global:FixtureDir 'citrix-workspace-switches.json' }
function global:Get-TvHostConfigPath { Join-Path $global:FixtureDir 'teamviewer-host-config.json' }
function global:Get-M365CustomConfigPath { Join-Path $global:FixtureDir 'm365-configuration.user.xml' }
function global:Add-LogLine { param([string]$Message) $global:LogLines += $Message }
function global:Set-DialogChromeFromOwner { param($Dialog, $Owner) }
function global:Invoke-RefreshGrid { param([switch]$DiscardSiteResults) $global:RefreshCount++ }
function global:Update-SidebarForSystems { $global:SidebarCount++ }
function global:Update-UiNow { }
# The probe keeps what Save-Preferences gets and writes it to the fixture, the way the shell writes it to its own folder.
function global:Save-Preferences {
    param($Prefs)
    $global:SavedCount++
    $Prefs | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $global:FixturePrefsPath -Encoding UTF8
}
function global:Show-ThemedMessage {
    param($Owner, [string]$Title, [string]$Message, $Buttons, $Icon)
    $global:Messages += ('{0}: {1}' -f $Title, $Message)
    return 'OK'
}
# The detection and download helpers reach the network and the registry; the probe has no use for them.
foreach ($stub in 'Get-GitHubApiAuthStatus', 'Install-IconPack', 'Install-IconPackFromFile', 'Install-IntuneWinAppUtil', 'Invoke-DetectIntuneWinAppUtil') {
    Set-Item -Path ('function:global:' + $stub) -Value { }
}
function global:Get-IconPackRoot { param([string]$AppRoot) Join-Path (Join-Path $global:PackagersRoot 'Icons') '' }
$global:LogLines = @()
$global:Messages = @()
$global:SavedCount = 0
$global:RefreshCount = 0
$global:SidebarCount = 0

function global:New-OptionsTestOwner {
    $owner = New-Object MahApps.Metro.Controls.MetroWindow
    $owner.Width = 400
    $owner.Height = 300
    return $owner
}

function global:Get-OptionsSectionDefinitionById {
    param([Parameter(Mandatory)][string]$Id)
    return @(Get-OptionsSectionDefinition | Where-Object { $_.Id -eq $Id })[0]
}

function global:New-OptionsTestPage {
    param([Parameter(Mandatory)][string]$Id, [hashtable]$Context)
    if (-not $Context) { $Context = New-OptionsContext -Prefs $script:Prefs; $Context.Owner = $null }
    return (New-OptionsSectionPage -Definition (Get-OptionsSectionDefinitionById -Id $Id) -Context $Context)
}
