BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-SidebarSystemState', 'Invoke-DiscoveryScan') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerDiscovery.psd1') -Force

    function New-TargetPrefs {
        param([string]$SiteCode, [string]$Server, [string]$Database)
        [pscustomobject]@{
            Systems              = [pscustomobject]@{ ConfigMgr = $true; Intune = $false; Wsus = $false }
            SiteCode             = $SiteCode
            DetectedTools        = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $false } }
            Intune               = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = '' }
            Wsus                 = [pscustomobject]@{ ServerName = '' }
            DiscoverySqlServer   = $Server
            DiscoverySqlDatabase = $Database
        }
    }
}

Describe 'Discover button rule and the scan agree on the SQL target' {
    BeforeEach {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { [pscustomobject]@{ SQLServerName = 'cm01'; SQLDatabaseName = 'CM_MCM' } }
        Mock -ModuleName AppPackagerDiscovery Get-DiscoveryProviderMachineName { 'cm01' }
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery { throw 'must not be called' }
        Mock Get-CMInventorySnapshot { [pscustomobject]@{ Rows = @(); ClientTotal = 0; InventoryNewest = $null; Applications = @() } }
    }

    It 'site code <SiteCode>, server <Server>, database <Database>: button enabled is <Enabled>' -ForEach @(
        @{ SiteCode = 'MCM'; Server = '';      Database = '';      Enabled = $true }
        @{ SiteCode = 'MCM'; Server = 'sql01'; Database = '';      Enabled = $true }
        @{ SiteCode = 'MCM'; Server = 'sql01'; Database = 'CM_X';  Enabled = $true }
        @{ SiteCode = '';    Server = 'sql01'; Database = 'CM_X';  Enabled = $true }
        @{ SiteCode = '  ';  Server = 'sql01'; Database = 'CM_X';  Enabled = $true }
        @{ SiteCode = '';    Server = 'sql01'; Database = '';      Enabled = $false }
        @{ SiteCode = '';    Server = '';      Database = 'CM_X';  Enabled = $false }
        @{ SiteCode = '';    Server = '';      Database = '';      Enabled = $false }
    ) {
        $sidebar = Get-SidebarSystemState -Prefs (New-TargetPrefs -SiteCode $SiteCode -Server $Server -Database $Database)
        $sidebar.DiscoverEnabled | Should -Be $Enabled

        $scan = Invoke-DiscoveryScan -SiteCode $SiteCode -ServerOverride $Server -DatabaseOverride $Database -Packager @() -HiddenScript @() -DismissedScript @()
        if ($Enabled) {
            $scan.Error | Should -BeNullOrEmpty
            $scan.Ok | Should -BeTrue
        }
        else {
            $scan.Ok | Should -BeFalse
            $scan.Error | Should -BeLike '*site code or SQL overrides are required*'
            Should -Invoke Get-CMInventorySnapshot -Times 0
        }
    }

    It 'uses both overrides as the target without a site code' {
        $scan = Invoke-DiscoveryScan -SiteCode '' -ServerOverride 'sql01' -DatabaseOverride 'CM_X' -Packager @() -HiddenScript @() -DismissedScript @()
        $scan.Server | Should -Be 'sql01'
        $scan.Database | Should -Be 'CM_X'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 0
    }

    It 'passes the stored SMS Provider to the site lookup' {
        $scan = Invoke-DiscoveryScan -SiteCode 'MCM' -ProviderMachineName ' cm02.contoso.com ' -Packager @() -HiddenScript @() -DismissedScript @()
        $scan.Ok | Should -BeTrue
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 1 -ParameterFilter { $ProviderMachineName -eq 'cm02.contoso.com' }
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoveryProviderMachineName -Times 0
    }
}
