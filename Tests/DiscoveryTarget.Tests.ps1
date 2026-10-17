BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-SidebarSystemState', 'Invoke-DiscoveryScan', 'Resolve-DiscoveryProviderName') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerDiscovery.psd1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psm1') -Force -DisableNameChecking

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

Describe 'Discover provider name from APP_PACKAGER_CM_PROVIDER' {
    It 'uses the environment variable when the preference is empty and no CMSite drive exists' {
        Mock Get-PSDrive { } -ParameterFilter { $PSProvider -eq 'CMSite' }
        $saved = $env:APP_PACKAGER_CM_PROVIDER
        try {
            $env:APP_PACKAGER_CM_PROVIDER = ' cm04.contoso.com '
            Resolve-DiscoveryProviderName -Preference '' -SiteCode 'MCM' | Should -Be 'cm04.contoso.com'
            Should -Invoke Get-PSDrive -Times 0 -ParameterFilter { $PSProvider -eq 'CMSite' }
        }
        finally {
            if ($null -eq $saved) { Remove-Item Env:\APP_PACKAGER_CM_PROVIDER -ErrorAction SilentlyContinue }
            else { $env:APP_PACKAGER_CM_PROVIDER = $saved }
        }
    }
}

Describe 'Discover provider name' {
    BeforeEach {
        Mock Resolve-CMProviderMachineName { $null }
    }

    It 'uses the stored preference, trimmed, without reading the site drive' {
        Mock Get-PSDrive { throw 'must not be called' }
        Resolve-DiscoveryProviderName -Preference ' cm02.contoso.com ' -SiteCode 'MCM' | Should -Be 'cm02.contoso.com'
        Should -Invoke Get-PSDrive -Times 0
        Should -Invoke Resolve-CMProviderMachineName -Times 0
    }

    It 'uses the shared provider resolution before the CMSite drives when the preference is empty' {
        Mock Resolve-CMProviderMachineName { ' cm03.contoso.com ' }
        Mock Get-PSDrive { @([pscustomobject]@{ Name = 'MCM'; Root = 'cm01.contoso.com' }) } -ParameterFilter { $PSProvider -eq 'CMSite' }
        Resolve-DiscoveryProviderName -Preference '' -SiteCode 'MCM' | Should -Be 'cm03.contoso.com'
        Should -Invoke Resolve-CMProviderMachineName -Times 1
        Should -Invoke Get-PSDrive -Times 0 -ParameterFilter { $PSProvider -eq 'CMSite' }
    }

    It 'uses the root of the site code CMSite drive when the preference is empty' {
        Mock Get-PSDrive {
            @([pscustomobject]@{ Name = 'LAB'; Root = 'lab01.contoso.com' }, [pscustomobject]@{ Name = 'MCM'; Root = ' cm01.contoso.com ' })
        } -ParameterFilter { $PSProvider -eq 'CMSite' }
        Resolve-DiscoveryProviderName -Preference '' -SiteCode 'MCM' | Should -Be 'cm01.contoso.com'
    }

    It 'uses the first CMSite drive when no drive has the site code' {
        Mock Get-PSDrive { @([pscustomobject]@{ Name = 'LAB'; Root = 'lab01.contoso.com' }) } -ParameterFilter { $PSProvider -eq 'CMSite' }
        Resolve-DiscoveryProviderName -Preference '   ' -SiteCode 'MCM' | Should -Be 'lab01.contoso.com'
    }

    It 'returns an empty name when no CMSite drive exists or the provider is not loaded' {
        Mock Get-PSDrive { } -ParameterFilter { $PSProvider -eq 'CMSite' }
        Resolve-DiscoveryProviderName -Preference '' -SiteCode 'MCM' | Should -Be ''
        Mock Get-PSDrive { throw 'Cannot find a provider with the name CMSite.' } -ParameterFilter { $PSProvider -eq 'CMSite' }
        Resolve-DiscoveryProviderName -Preference $null -SiteCode '' | Should -Be ''
    }
}
