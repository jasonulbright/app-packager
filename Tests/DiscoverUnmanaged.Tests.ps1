BeforeAll {
    $t = $null; $e = $null
    $script:GuiPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1'
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:GuiPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Read-Preferences', 'Resolve-FirstRunCompleted', 'Get-SidebarSystemState', 'Get-DiscoveryPackagerInput',
        'Invoke-DiscoveryScan', 'Enable-DiscoveredApplication', 'Add-DiscoveryDismissed', 'Get-DiscoveryPromptText') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:prefPath = Join-Path $TestDrive 'preferences.json'
    function Get-PreferencesPath { $script:prefPath }

    # Stand-ins for the discovery module; each test replaces them with mocks.
    function Resolve-CMDatabaseTarget { param([string]$SiteCode, [string]$ServerOverride, [string]$DatabaseOverride) }
    function Get-CMInventorySnapshot { param([string]$Server, [string]$Database, [int]$TimeoutSeconds = 30) }
    function Find-DiscoveryCandidate { param($Snapshot, [object[]]$Packager, [string[]]$HiddenScript, [string[]]$DismissedScript) }
    function Get-DiscoveryStatusText { param($Snapshot, [int]$CandidateCount) }

    function New-ScanPackager {
        @(
            [pscustomobject]@{ Script = 'package-7zip.ps1'; Vendor = '7-Zip'; App = '7-Zip'; CMName = '7-Zip'; ArpName = '7-Zip [0-9]*' }
            [pscustomobject]@{ Script = 'package-git.ps1'; Vendor = 'Git'; App = 'Git'; CMName = 'Git'; ArpName = 'Git' }
        )
    }
}

Describe 'Discovery preferences' {
    It 'defaults to an empty dismissed list and no SQL overrides' {
        Set-Content $script:prefPath -Value '{}'
        $p = Read-Preferences
        @($p.DiscoveryDismissed).Count | Should -Be 0
        $p.DiscoverySqlServer | Should -BeExactly ''
        $p.DiscoverySqlDatabase | Should -BeExactly ''
    }

    It 'has the same defaults when no preferences file exists' {
        Remove-Item $script:prefPath -ErrorAction SilentlyContinue
        $p = Read-Preferences
        @($p.DiscoveryDismissed).Count | Should -Be 0
        $p.DiscoverySqlServer | Should -BeExactly ''
        $p.DiscoverySqlDatabase | Should -BeExactly ''
    }

    It 'loads stored values, trims the overrides and drops entries that are not packager scripts' {
        Set-Content $script:prefPath -Value '{ "DiscoveryDismissed": ["package-git.ps1", "", "evil.ps1", "package-7zip.ps1", "package-git.ps1"], "DiscoverySqlServer": " cm01.contoso.com ", "DiscoverySqlDatabase": " CM_MCM " }'
        $p = Read-Preferences
        $p.DiscoveryDismissed | Should -Be @('package-git.ps1', 'package-7zip.ps1')
        $p.DiscoverySqlServer | Should -BeExactly 'cm01.contoso.com'
        $p.DiscoverySqlDatabase | Should -BeExactly 'CM_MCM'
    }

    It 'survives a save and load through JSON' {
        Set-Content $script:prefPath -Value '{}'
        $p = Read-Preferences
        $p.DiscoveryDismissed = @('package-git.ps1')
        $p.DiscoverySqlServer = 'sql01'
        $p.DiscoverySqlDatabase = 'CM_ABC'
        Set-Content $script:prefPath -Value ($p | ConvertTo-Json -Depth 5)
        $back = Read-Preferences
        $back.DiscoveryDismissed | Should -Be @('package-git.ps1')
        $back.DiscoverySqlServer | Should -BeExactly 'sql01'
        $back.DiscoverySqlDatabase | Should -BeExactly 'CM_ABC'
    }

    It 'is written by Save-Preferences, which serializes the whole preference object' {
        $source = [System.IO.File]::ReadAllText($script:GuiPath)
        $source | Should -Match '\$json = \$Prefs \| ConvertTo-Json -Depth 5'
    }
}

Describe 'Enabling discovered applications' {
    It 'removes the scripts from HiddenApplications without regard to case and keeps the rest' {
        $prefs = [pscustomobject]@{ HiddenApplications = @('package-git.ps1', 'package-7zip.ps1', 'package-vlc.ps1') }
        $removed = Enable-DiscoveredApplication -Prefs $prefs -Script @('PACKAGE-GIT.ps1', 'package-vlc.ps1', 'package-notthere.ps1')
        $removed | Should -Be 2
        $prefs.HiddenApplications | Should -Be @('package-7zip.ps1')
    }

    It 'leaves an empty array, not null, when the last hidden script is enabled' {
        $prefs = [pscustomobject]@{ HiddenApplications = @('package-git.ps1') }
        [void](Enable-DiscoveredApplication -Prefs $prefs -Script @('package-git.ps1'))
        , $prefs.HiddenApplications | Should -BeOfType [array]
        @($prefs.HiddenApplications).Count | Should -Be 0
    }

    It 'adds dismissed scripts once each' {
        $prefs = [pscustomobject]@{ DiscoveryDismissed = @('package-git.ps1') }
        $added = Add-DiscoveryDismissed -Prefs $prefs -Script @('package-git.ps1', 'package-7zip.ps1', 'PACKAGE-7ZIP.ps1')
        $added | Should -Be 1
        $prefs.DiscoveryDismissed | Should -Be @('package-git.ps1', 'package-7zip.ps1')
    }
}

Describe 'Discovery prompt text' {
    It 'reads "<Expected>" for <Count> candidate(s)' -ForEach @(
        @{ Count = 1; Expected = 'Discovered 1 app. Enable console view?' }
        @{ Count = 3; Expected = 'Discovered 3 apps. Enable console view?' }
    ) {
        Get-DiscoveryPromptText -Count $Count | Should -BeExactly $Expected
    }
}

Describe 'Discovery packager input' {
    It 'maps grid rows to the module packager shape' {
        $rows = @(
            [pscustomobject]@{ Script = 'package-git.ps1'; Vendor = 'Git'; Application = 'Git for Windows'; CMName = 'Git'; ArpName = 'Git'; Status = 'Ready' }
            [pscustomobject]@{ Script = 'package-bad.ps1'; Vendor = ''; Application = ''; CMName = ''; ArpName = ''; Status = 'Read error: x' }
        )
        $out = @(Get-DiscoveryPackagerInput -Rows $rows)
        $out.Count | Should -Be 1
        $out[0].Script | Should -Be 'package-git.ps1'
        $out[0].App | Should -Be 'Git for Windows'
        $out[0].Vendor | Should -Be 'Git'
        $out[0].CMName | Should -Be 'Git'
        $out[0].ArpName | Should -Be 'Git'
    }
}

Describe 'Invoke-DiscoveryScan' {
    BeforeEach {
        $script:snap = [pscustomobject]@{ Rows = @([pscustomobject]@{ DisplayName = 'Git'; Publisher = 'Git'; Version = '2.0'; Architecture = 'x64'; ClientCount = 3 }); ClientTotal = 4; InventoryNewest = (Get-Date); Applications = @() }
        Mock Resolve-CMDatabaseTarget { [pscustomobject]@{ Server = 'cm01'; Database = 'CM_MCM' } }
        Mock Get-CMInventorySnapshot { $script:snap }
        Mock Find-DiscoveryCandidate { @([pscustomobject]@{ Script = 'package-git.ps1'; CMName = 'Git'; Vendor = 'Git'; ClientCount = 3; MatchedName = 'Git' }) }
        Mock Get-DiscoveryStatusText { 'Inventory from 4 clients. 1 app found.' }
    }

    It 'returns the candidates and the status text' {
        $r = Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @('package-git.ps1') -DismissedScript @()
        $r.Ok | Should -BeTrue
        @($r.Candidates).Count | Should -Be 1
        $r.Candidates[0].Script | Should -Be 'package-git.ps1'
        $r.StatusText | Should -Be 'Inventory from 4 clients. 1 app found.'
        Should -Invoke Get-CMInventorySnapshot -Times 1 -ParameterFilter { $Server -eq 'cm01' -and $Database -eq 'CM_MCM' }
        Should -Invoke Get-DiscoveryStatusText -Times 1 -ParameterFilter { $CandidateCount -eq 1 }
    }

    It 'passes the hidden and dismissed lists to the matcher' {
        [void](Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @('package-git.ps1', 'package-7zip.ps1') -DismissedScript @('package-7zip.ps1'))
        Should -Invoke Find-DiscoveryCandidate -Times 1 -ParameterFilter {
            @($HiddenScript).Count -eq 2 -and @($DismissedScript) -contains 'package-7zip.ps1' -and @($Packager).Count -eq 2
        }
    }

    It 'passes the SQL overrides only when they are set' {
        [void](Invoke-DiscoveryScan -SiteCode 'MCM' -ServerOverride 'sql01' -DatabaseOverride 'CM_X' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @())
        Should -Invoke Resolve-CMDatabaseTarget -Times 1 -ParameterFilter { $SiteCode -eq 'MCM' -and $ServerOverride -eq 'sql01' -and $DatabaseOverride -eq 'CM_X' }
        [void](Invoke-DiscoveryScan -SiteCode 'MCM' -ServerOverride '' -DatabaseOverride '  ' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @())
        Should -Invoke Resolve-CMDatabaseTarget -Times 1 -ParameterFilter { [string]::IsNullOrEmpty($ServerOverride) -and [string]::IsNullOrEmpty($DatabaseOverride) }
    }

    It 'returns zero candidates with the status text and no error' {
        Mock Find-DiscoveryCandidate { }
        Mock Get-DiscoveryStatusText { 'No inventory data.' }
        $r = Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @()
        $r.Ok | Should -BeTrue
        @($r.Candidates).Count | Should -Be 0
        $r.StatusText | Should -Be 'No inventory data.'
        Should -Invoke Get-DiscoveryStatusText -Times 1 -ParameterFilter { $CandidateCount -eq 0 }
    }

    It 'returns the plain message when the SQL query fails' {
        Mock Get-CMInventorySnapshot { throw 'Login failed for the ConfigMgr site database on cm01.' }
        $r = Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @()
        $r.Ok | Should -BeFalse
        $r.Error | Should -Be 'Login failed for the ConfigMgr site database on cm01.'
        @($r.Candidates).Count | Should -Be 0
        Should -Invoke Find-DiscoveryCandidate -Times 0
    }

    It 'returns a plain message when the target has no server' {
        Mock Resolve-CMDatabaseTarget { [pscustomobject]@{ Server = ''; Database = 'CM_MCM' } }
        $r = Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @()
        $r.Ok | Should -BeFalse
        $r.Error | Should -Match 'SQL server'
        Should -Invoke Get-CMInventorySnapshot -Times 0
    }

    It 'returns a plain message when the discovery module is not loaded' {
        Mock Get-Command { $null } -ParameterFilter { $Name -eq 'Get-CMInventorySnapshot' }
        $r = Invoke-DiscoveryScan -SiteCode 'MCM' -Packager (New-ScanPackager) -HiddenScript @() -DismissedScript @()
        $r.Ok | Should -BeFalse
        $r.Error | Should -Match 'module'
    }
}

Describe 'Discover button availability' {
    BeforeAll {
        function New-SidebarPrefs {
            param([string]$SiteCode = 'MCM', [bool]$ConfigMgr = $true, [bool]$Console = $false, [string]$Server = '', [string]$Database = '')
            [pscustomobject]@{
                Systems              = [pscustomobject]@{ ConfigMgr = $ConfigMgr; Intune = $false; Wsus = $false }
                SiteCode             = $SiteCode
                DetectedTools        = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $Console } }
                Intune               = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = '' }
                Wsus                 = [pscustomobject]@{ ServerName = '' }
                DiscoverySqlServer   = $Server
                DiscoverySqlDatabase = $Database
            }
        }
    }

    It 'is enabled with a site code and no console' {
        $s = Get-SidebarSystemState -Prefs (New-SidebarPrefs)
        $s.DiscoverEnabled | Should -BeTrue
        $s.DiscoverToolTip | Should -Match 'hardware inventory'
    }

    It 'is disabled without a site code or a database override and says what to set' {
        $s = Get-SidebarSystemState -Prefs (New-SidebarPrefs -SiteCode '')
        $s.DiscoverEnabled | Should -BeFalse
        $s.DiscoverToolTip | Should -Match 'site code'
    }

    It 'is enabled without a site code when both SQL overrides are set' {
        $s = Get-SidebarSystemState -Prefs (New-SidebarPrefs -SiteCode '' -Server 'sql01' -Database 'CM_MCM')
        $s.DiscoverEnabled | Should -BeTrue
    }

    It 'is disabled when ConfigMgr is not in use' {
        $s = Get-SidebarSystemState -Prefs (New-SidebarPrefs -ConfigMgr $false)
        $s.DiscoverEnabled | Should -BeFalse
        $s.DiscoverToolTip | Should -Match 'not in use'
    }
}

Describe 'Discover wiring in the GUI' {
    BeforeAll {
        $script:Source = [System.IO.File]::ReadAllText($script:GuiPath)
        $script:Xaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\MainWindow.xaml'))
    }

    It 'places the Discover button right after Check ConfigMgr in the same style' {
        $script:Xaml | Should -Match '(?s)x:Name="btnCheckMECM".*?/>\s*(<!--[^>]*-->\s*)?<Button x:Name="btnDiscover"\s+Style="\{StaticResource WorkflowButtonStyle\}"'
        $script:Xaml | Should -Match '(?s)x:Name="btnDiscover".*?<TextBlock Text="Discover unmanaged software" TextWrapping="Wrap"/>\s*</Button>'
    }

    It 'imports the discovery module in the GUI and in the background runspace' {
        ([regex]::Matches($script:Source, "AppPackagerDiscovery\.psm1")).Count | Should -BeGreaterOrEqual 2
    }

    It 'disables and themes the Discover button with the other workflow buttons' {
        $script:Source | Should -Match '\$script:WorkflowButtons = @\([^)]*\$btnDiscover'
        $script:Source | Should -Match '\$btnDiscover\.IsEnabled\s+= \$Enabled'
        $script:Source | Should -Match "@\(\`$btnDiscover, 'Discover'\)"
    }
}
