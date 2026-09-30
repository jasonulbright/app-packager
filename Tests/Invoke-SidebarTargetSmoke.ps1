#Requires -Version 5.1

<#
.SYNOPSIS
    Headless checks for the sidebar publish buttons' availability logic.

.DESCRIPTION
    Parses start-apppackager.ps1, extracts Get-SidebarSystemState from the AST,
    and asserts which publish buttons are available, and what their tooltips
    say, for each combination of configured systems. No WPF assemblies are
    loaded and no window is shown.

.EXAMPLE
    .\Tests\Invoke-SidebarTargetSmoke.ps1
#>

[CmdletBinding()]
param(
    [string]$ScriptPath = ''
)

$ErrorActionPreference = 'Stop'
# $PSScriptRoot is empty inside a parameter default under powershell.exe -File.
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path -Parent (Split-Path -Parent $PSCommandPath)) 'start-apppackager.ps1' }

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    foreach ($e in $errors) { Write-Host ("PARSE  {0}" -f $e.Message) -ForegroundColor Red }
    exit 1
}

foreach ($name in @('Get-SidebarSystemState')) {
    $fn = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true) | Select-Object -First 1
    if (-not $fn) { Write-Host ("MISSING function {0}" -f $name) -ForegroundColor Red; exit 1 }
    . ([scriptblock]::Create($fn.Extent.Text))
}

function New-TestPrefs {
    param([bool]$Console = $true, [string]$SiteCode = 'MCM', [bool]$Intune = $false, [string]$WsusServer = '', [string[]]$InUse = @('ConfigMgr', 'Intune', 'Wsus'))
    [pscustomobject]@{
        Systems       = [pscustomobject]@{ ConfigMgr = ('ConfigMgr' -in $InUse); Intune = ('Intune' -in $InUse); Wsus = ('Wsus' -in $InUse) }
        SiteCode      = $SiteCode
        DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $Console } }
        Intune        = [pscustomobject]@{
            TenantId              = $(if ($Intune) { 'tenant' } else { '' })
            ClientId              = $(if ($Intune) { 'client' } else { '' })
            ClientSecretProtected = $(if ($Intune) { 'protected' } else { '' })
            DeploymentTarget      = 'MECM'
        }
        Wsus          = [pscustomobject]@{ ServerName = $WsusServer }
    }
}

# Update-SidebarForSystems must stay the only writer of the publish buttons'
# availability and tooltips, and must be wired to every refresh point.
$source = Get-Content -Path $ScriptPath -Raw
$callCount = ([regex]::Matches($source, 'Update-SidebarForSystems')).Count

$mecmOnly  = Get-SidebarSystemState -Prefs (New-TestPrefs)
$noConsole = Get-SidebarSystemState -Prefs (New-TestPrefs -Console $false -WsusServer 'wsus01')
$noSite    = Get-SidebarSystemState -Prefs (New-TestPrefs -SiteCode '')
$all       = Get-SidebarSystemState -Prefs (New-TestPrefs -Intune $true -WsusServer 'wsus01')
$halfIntune = New-TestPrefs
$halfIntune.Intune.TenantId = 'tenant'
$half      = Get-SidebarSystemState -Prefs $halfIntune
$mecmOff   = Get-SidebarSystemState -Prefs (New-TestPrefs -WsusServer 'wsus01' -InUse @('Wsus'))
$wsusOff   = Get-SidebarSystemState -Prefs (New-TestPrefs -Intune $true -WsusServer 'wsus01' -InUse @('ConfigMgr', 'Intune'))

$cases = @(
    @{ Name = 'ConfigMgr only enables Check ConfigMgr and Publish to ConfigMgr'; Actual = { $mecmOnly.CheckMecmEnabled -and $mecmOnly.MecmEnabled }; Expected = $true }
    @{ Name = 'ConfigMgr only disables Publish to Intune';     Actual = { $mecmOnly.IntuneEnabled };      Expected = $false }
    @{ Name = 'ConfigMgr only disables Publish to WSUS and WSUS Updates'; Actual = { $mecmOnly.WsusEnabled -or $mecmOnly.WsusUpdatesEnabled }; Expected = $false }
    @{ Name = 'a disabled WSUS button says where to set the server'; Actual = { $mecmOnly.WsusToolTip -match 'WSUS server' -and $mecmOnly.WsusUpdatesToolTip -match 'WSUS server' }; Expected = $true }
    @{ Name = 'a disabled Intune button names the three credentials'; Actual = { $mecmOnly.IntuneToolTip -match 'Tenant ID' -and $mecmOnly.IntuneToolTip -match 'Client Secret' }; Expected = $true }
    @{ Name = 'no console disables both ConfigMgr buttons';     Actual = { $noConsole.CheckMecmEnabled -or $noConsole.MecmEnabled }; Expected = $false }
    @{ Name = 'no console names the console in the tooltip';    Actual = { $noConsole.MecmToolTip -match 'console' }; Expected = $true }
    @{ Name = 'no console still enables WSUS with a server';    Actual = { $noConsole.WsusEnabled -and $noConsole.WsusUpdatesEnabled }; Expected = $true }
    @{ Name = 'no site code disables Publish to ConfigMgr';     Actual = { $noSite.MecmEnabled };         Expected = $false }
    @{ Name = 'no site code names the site code';               Actual = { $noSite.MecmToolTip -match 'site code' }; Expected = $true }
    @{ Name = 'all systems enable every button';                Actual = { $all.CheckMecmEnabled -and $all.MecmEnabled -and $all.IntuneEnabled -and $all.WsusEnabled -and $all.WsusUpdatesEnabled }; Expected = $true }
    @{ Name = 'an enabled WSUS button describes the older-version rule'; Actual = { $all.WsusToolTip -match 'older version' }; Expected = $true }
    @{ Name = 'a tenant without client and secret keeps Intune disabled'; Actual = { $half.IntuneEnabled }; Expected = $false }
    @{ Name = 'ConfigMgr not in use disables both ConfigMgr buttons despite console and site code'; Actual = { $mecmOff.CheckMecmEnabled -or $mecmOff.MecmEnabled }; Expected = $false }
    @{ Name = 'ConfigMgr not in use says so and names Systems in use';   Actual = { $mecmOff.MecmToolTip -match 'not in use' -and $mecmOff.MecmToolTip -match 'Systems in use' }; Expected = $true }
    @{ Name = 'ConfigMgr not in use keeps WSUS enabled';               Actual = { $mecmOff.WsusEnabled -and $mecmOff.WsusUpdatesEnabled }; Expected = $true }
    @{ Name = 'WSUS not in use disables its buttons despite a server';  Actual = { $wsusOff.WsusEnabled -or $wsusOff.WsusUpdatesEnabled }; Expected = $false }
    @{ Name = 'WSUS not in use keeps ConfigMgr and Intune enabled';    Actual = { $wsusOff.MecmEnabled -and $wsusOff.IntuneEnabled }; Expected = $true }
    @{ Name = 'Setup and Options both save the systems in use';        Actual = { ([regex]::Matches($source, '\$prefsRef\.Systems\.ConfigMgr\s*=')).Count -eq 2 }; Expected = $true }
    @{ Name = 'the publish and One Click paths fall back to the packager default site code'; Actual = { ([regex]::Matches($source, 'if \(\[string\]::IsNullOrWhiteSpace\(\$siteCodeValue\)\) \{ \$siteCodeValue = ''MCM'' \}')).Count -ge 2 }; Expected = $true }
    @{ Name = 'Setup without ConfigMgr saves an empty site code'; Actual = { $source -match '\} else \{\s+#[^\n]*\n\s+#[^\n]*\n\s+\$prefsRef\.SiteCode = ''''\s+\}' }; Expected = $true }
    @{ Name = 'Check Latest runs without a site code when ConfigMgr is not in use'; Actual = { $source -match 'if \(\$script:Prefs\.Systems\.ConfigMgr\) \{\s+Add-LogLine -Message \"SiteCode is required' }; Expected = $true }
    @{ Name = 'the sidebar refresh is wired to launch, Options, Setup and the button reset'; Actual = { $callCount -ge 5 }; Expected = $true }
    @{ Name = 'the retired target-driven refresh is gone';      Actual = { $source -notmatch 'Update-SidebarForDeploymentTarget|Get-SidebarTargetState' }; Expected = $true }
    @{ Name = 'the sidebar refresh enables a button whose system became configured'; Actual = { $source -match '\$button\.IsEnabled = \(\[bool\]\$script:ActionButtonsEnabled -and \[bool\]\$state\[' }; Expected = $true }
    @{ Name = 'a running pipeline keeps the sidebar refresh from enabling buttons'; Actual = { $source -match 'function Set-ActionButtonsEnabled \{\s+param\(\[bool\]\$Enabled\)\s+\$script:ActionButtonsEnabled = \$Enabled' }; Expected = $true }
)

$failed = 0
foreach ($case in $cases) {
    $actual = & $case.Actual
    if ($actual -eq $case.Expected) {
        Write-Host ("PASS  {0}" -f $case.Name) -ForegroundColor Green
    } else {
        Write-Host ("FAIL  {0} (expected {1}, got {2})" -f $case.Name, $case.Expected, $actual) -ForegroundColor Red
        $failed++
    }
}

Write-Host ("`n{0} passed, {1} failed." -f ($cases.Count - $failed), $failed)
exit ([int]($failed -gt 0))
