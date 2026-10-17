BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'AppPackagerDiscovery.psd1') -Force

    if (-not ('DiscoveryTestSqlError' -as [type])) {
        Add-Type -TypeDefinition @'
public class DiscoveryTestSqlError : System.Exception {
    public int Number;
    public DiscoveryTestSqlError(int number, string message) : base(message) { Number = number; }
}
'@
    }

    function New-ArpTable {
        param([object[]]$Rows)
        $t = New-Object System.Data.DataTable
        foreach ($c in 'DisplayName', 'Publisher', 'Version', 'Architecture') { [void]$t.Columns.Add($c, [string]) }
        [void]$t.Columns.Add('ClientCount', [int])
        foreach ($r in $Rows) {
            $row = $t.NewRow()
            $row.DisplayName = $r[0]
            $row.Publisher = if ($null -eq $r[1]) { [DBNull]::Value } else { $r[1] }
            $row.Version = if ($null -eq $r[2]) { [DBNull]::Value } else { $r[2] }
            $row.Architecture = $r[3]
            $row.ClientCount = $r[4]
            $t.Rows.Add($row)
        }
        return , $t
    }

    function New-SnapshotDataSet {
        param([object[]]$ArpRows = @(), [int]$ClientTotal = 0, $Newest = [DBNull]::Value, [string[]]$Applications = @())
        $ds = New-Object System.Data.DataSet
        $ds.Tables.Add((New-ArpTable -Rows $ArpRows))
        $s = New-Object System.Data.DataTable
        [void]$s.Columns.Add('ClientTotal', [int])
        [void]$s.Columns.Add('InventoryNewest', [datetime])
        $row = $s.NewRow(); $row.ClientTotal = $ClientTotal; $row.InventoryNewest = if ($Newest -is [DBNull]) { $Newest } else { [datetime]$Newest }; $s.Rows.Add($row)
        $ds.Tables.Add($s)
        $a = New-Object System.Data.DataTable
        [void]$a.Columns.Add('DisplayName', [string])
        foreach ($n in $Applications) { $r = $a.NewRow(); $r.DisplayName = $n; $a.Rows.Add($r) }
        $ds.Tables.Add($a)
        return , $ds
    }

    function New-Snapshot {
        param([object[]]$Rows = @(), [int]$ClientTotal = 4, $InventoryNewest = (Get-Date '2026-10-15T02:55:30'), [string[]]$Applications = @())
        [pscustomobject]@{
            Rows            = @($Rows | ForEach-Object {
                    [pscustomobject]@{ DisplayName = $_[0]; Publisher = $_[1]; Version = $_[2]; Architecture = $_[3]; ClientCount = $_[4] }
                })
            ClientTotal     = $ClientTotal
            InventoryNewest = $InventoryNewest
            Applications    = $Applications
        }
    }

    function New-Packager {
        param([string]$Script, [string]$CMName, [string]$ArpName, [string]$Vendor = 'Vendor', [string]$App = 'App')
        [pscustomobject]@{ Script = $Script; Vendor = $Vendor; App = $App; CMName = $CMName; ArpName = $ArpName }
    }
}

Describe 'Resolve-CMDatabaseTarget' {
    It 'uses both overrides without asking the site' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'must not be called' }
        $t = Resolve-CMDatabaseTarget -SiteCode 'MCM' -ServerOverride ' sql01.contoso.com ' -DatabaseOverride ' CM_ALT '
        $t.Server | Should -Be 'sql01.contoso.com'
        $t.Database | Should -Be 'CM_ALT'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 0
    }

    It 'reads the server and database from the site definition on the SMS Provider' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition {
            [pscustomobject]@{ SQLServerName = 'CM01.contoso.com'; SQLDatabaseName = 'CM_MCM' }
        }
        $t = Resolve-CMDatabaseTarget -SiteCode 'MCM' -ProviderMachineName 'cm01.contoso.com'
        $t.Server | Should -Be 'CM01.contoso.com'
        $t.Database | Should -Be 'CM_MCM'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 1 -ParameterFilter {
            $ProviderMachineName -eq 'cm01.contoso.com' -and $SiteCode -eq 'MCM'
        }
    }

    It 'moves a named instance from the database name to the server' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition {
            [pscustomobject]@{ SQLServerName = 'SQL01.contoso.com'; SQLDatabaseName = 'CMINST\CM_ABC' }
        }
        $t = Resolve-CMDatabaseTarget -SiteCode 'ABC' -ProviderMachineName 'cm01'
        $t.Server | Should -Be 'SQL01.contoso.com\CMINST'
        $t.Database | Should -Be 'CM_ABC'
    }

    It 'lets the server override win over the site definition' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition {
            [pscustomobject]@{ SQLServerName = 'CM01.contoso.com'; SQLDatabaseName = 'CM_MCM' }
        }
        $t = Resolve-CMDatabaseTarget -SiteCode 'MCM' -ProviderMachineName 'cm01' -ServerOverride 'aglistener.contoso.com'
        $t.Server | Should -Be 'aglistener.contoso.com'
        $t.Database | Should -Be 'CM_MCM'
    }

    It 'defaults the database to CM_<SiteCode> when the site lookup fails and a server override exists' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'Access denied' }
        $t = Resolve-CMDatabaseTarget -SiteCode 'mcm' -ProviderMachineName 'cm01' -ServerOverride 'sql01'
        $t.Server | Should -Be 'sql01'
        $t.Database | Should -Be 'CM_MCM'
    }

    It 'throws a plain message when no server is known' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'RPC server unavailable' }
        { Resolve-CMDatabaseTarget -SiteCode 'MCM' -ProviderMachineName 'cm01' } |
            Should -Throw -ExpectedMessage '*SQL Server for site MCM*RPC server unavailable*'
    }

    It 'throws a plain message when no SMS Provider and no override exist' {
        InModuleScope AppPackagerDiscovery { Mock Get-DiscoveryProviderMachineName { $null } }
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'must not be called' }
        { Resolve-CMDatabaseTarget -SiteCode 'MCM' } | Should -Throw -ExpectedMessage '*SQL Server for site MCM*'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 0
    }

    It 'rejects a site code that is not three letters or digits' {
        { Resolve-CMDatabaseTarget -SiteCode "M';--" -ServerOverride 'x' } | Should -Throw
    }
}

Describe 'Connection and command shape' {
    It 'builds a Windows-auth, read-intent connection string without string concatenation' {
        InModuleScope AppPackagerDiscovery {
            $cs = New-DiscoveryConnectionString -Server 'sql01;User ID=sa' -Database 'CM_MCM;Integrated Security=false' -TimeoutSeconds 30
            $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $cs
            $b['Data Source'] | Should -Be 'sql01;User ID=sa'
            $b['Initial Catalog'] | Should -Be 'CM_MCM;Integrated Security=false'
            $b['Integrated Security'] | Should -BeTrue
            $b['User ID'] | Should -BeNullOrEmpty
            [string]$b['ApplicationIntent'] | Should -Be 'ReadOnly'
            $cs | Should -Not -Match 'Encrypt'
            $cs | Should -Not -Match 'Password'
        }
    }

    It 'binds every value as a typed SqlParameter and sets the command timeout' {
        InModuleScope AppPackagerDiscovery {
            $cn = New-Object System.Data.SqlClient.SqlConnection
            $cmd = New-DiscoverySqlCommand -Connection $cn -CommandText 'SELECT @A, @B' -Parameters ([ordered]@{ A = 'x86'; B = 31 }) -TimeoutSeconds 45
            $cmd.CommandTimeout | Should -Be 45
            $cmd.CommandText | Should -Be 'SELECT @A, @B'
            $cmd.Parameters.Count | Should -Be 2
            $cmd.Parameters['@A'].SqlDbType | Should -Be ([System.Data.SqlDbType]::NVarChar)
            $cmd.Parameters['@A'].Value | Should -Be 'x86'
            $cmd.Parameters['@B'].SqlDbType | Should -Be ([System.Data.SqlDbType]::Int)
            $cmd.Parameters['@B'].Value | Should -Be 31
            $cmd.Dispose(); $cn.Dispose()
        }
    }
}

Describe 'Get-CMInventorySnapshot' {
    BeforeEach {
        $script:Captured = $null
    }

    It 'sends a parameterized read-only query over the v_ views and passes the timeout' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            $script:Captured = [pscustomobject]@{ ConnectionString = $ConnectionString; CommandText = $CommandText; Parameters = $Parameters; TimeoutSeconds = $TimeoutSeconds }
            New-SnapshotDataSet
        }
        $null = Get-CMInventorySnapshot -Server 'CM01.contoso.com' -Database 'CM_MCM' -TimeoutSeconds 12
        $c = $script:Captured
        $c.TimeoutSeconds | Should -Be 12
        $c.ConnectionString | Should -Match 'Integrated Security=True'
        $c.ConnectionString | Should -Match 'ApplicationIntent=ReadOnly'
        $c.CommandText | Should -Not -Match 'CM01|CM_MCM'
        $c.CommandText | Should -Not -Match "'x86'|'x64'"
        $c.CommandText | Should -Match '@Arch32'
        $c.CommandText | Should -Match '@Arch64'
        $c.Parameters['Arch32'] | Should -Be 'x86'
        $c.Parameters['Arch64'] | Should -Be 'x64'
        $c.CommandText | Should -Not -Match '(?i)\b(insert|update|delete|merge|exec|execute|drop|alter|create|truncate|grant)\b'
        $targets = [regex]::Matches($c.CommandText, '(?i)\b(?:from|join)\s+([\w\.\[\]]+)') | ForEach-Object { $_.Groups[1].Value }
        $targets.Count | Should -BeGreaterThan 0
        foreach ($t in $targets) { $t | Should -Match '^(dbo\.)?v_' }
        $c.CommandText | Should -Match 'v_GS_ADD_REMOVE_PROGRAMS\b'
        $c.CommandText | Should -Match 'v_GS_ADD_REMOVE_PROGRAMS_64\b'
        $c.CommandText | Should -Match 'v_Applications\b'
    }

    It 'never selects machine names' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            $script:Captured = $CommandText
            New-SnapshotDataSet
        }
        $null = Get-CMInventorySnapshot -Server 's' -Database 'd'
        $script:Captured | Should -Not -Match '(?i)\bName0\b|Netbios|v_R_System|ResourceName|MachineName'
    }

    It 'merges 32-bit and 64-bit rows and maps the summary' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows @(
                , @('7-Zip 26.03', 'Igor Pavlov', '26.03', 'x86', 1)
                , @('7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64', 3)
                , @('Notepad++ (64-bit x64)', $null, $null, 'x64', 2)
            ) -ClientTotal 4 -Newest (Get-Date '2026-10-15T02:55:30') -Applications @('Audacity 3.7.9', 'CMTrace Open')
        }
        $s = Get-CMInventorySnapshot -Server 's' -Database 'd'
        @($s.Rows).Count | Should -Be 3
        @($s.Rows | Where-Object Architecture -eq 'x86').Count | Should -Be 1
        @($s.Rows | Where-Object Architecture -eq 'x64').Count | Should -Be 2
        ($s.Rows | Where-Object DisplayName -eq '7-Zip 26.03 (x64)').ClientCount | Should -Be 3
        ($s.Rows | Where-Object DisplayName -like 'Notepad*').Publisher | Should -Be ''
        $s.ClientTotal | Should -Be 4
        $s.InventoryNewest | Should -Be (Get-Date '2026-10-15T02:55:30')
        $s.Applications | Should -Be @('Audacity 3.7.9', 'CMTrace Open')
    }

    It 'returns only the contract properties' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows @(, @('A', 'P', '1', 'x86', 1)) -ClientTotal 1 -Newest (Get-Date)
        }
        $s = Get-CMInventorySnapshot -Server 's' -Database 'd'
        @($s.PSObject.Properties.Name) | Should -Be @('Rows', 'ClientTotal', 'InventoryNewest', 'Applications')
        @($s.Rows[0].PSObject.Properties.Name) | Should -Be @('DisplayName', 'Publisher', 'Version', 'Architecture', 'ClientCount')
    }

    It 'returns an empty snapshot for zero inventory rows' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery { New-SnapshotDataSet }
        $s = Get-CMInventorySnapshot -Server 's' -Database 'd'
        @($s.Rows).Count | Should -Be 0
        $s.ClientTotal | Should -Be 0
        $s.InventoryNewest | Should -BeNullOrEmpty
        @($s.Applications).Count | Should -Be 0
    }

    It 'reports a denied login in plain words' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(18456, "Login failed for user 'CONTOSO\jdoe'.")
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_MCM' } |
            Should -Throw -ExpectedMessage '*CM01 denied the Windows login*db_datareader*CM_MCM*'
    }

    It 'reports an anonymous login from a second remoting hop as a delegation problem' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(18456, "Login failed for user 'NT AUTHORITY\ANONYMOUS LOGON'.")
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_MCM' } |
            Should -Throw -ExpectedMessage '*anonymous*delegat*'
    }

    It 'reports a failed Windows authentication handshake in plain words' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(0, 'The target principal name is incorrect.  Cannot generate SSPI context.')
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_MCM' } |
            Should -Throw -ExpectedMessage '*Windows authentication to SQL Server CM01 failed*domain*'
    }

    It 'reports an unreachable server in plain words' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(53, 'A network-related or instance-specific error occurred while establishing a connection to SQL Server.')
        }
        { Get-CMInventorySnapshot -Server 'SQLX' -Database 'CM_MCM' } |
            Should -Throw -ExpectedMessage '*Cannot reach SQL Server SQLX*'
    }

    It 'reports a wrong database in plain words' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(4060, 'Cannot open database "CM_XYZ" requested by the login. The login failed.')
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_XYZ' } |
            Should -Throw -ExpectedMessage '*Database CM_XYZ*CM01*'
    }

    It 'reports a timeout in plain words' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [DiscoveryTestSqlError]::new(-2, 'Execution Timeout Expired.')
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_MCM' -TimeoutSeconds 5 } |
            Should -Throw -ExpectedMessage '*did not answer within 5 seconds*'
    }

    It 'finds the SQL error number on an inner exception' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            throw [System.Management.Automation.MethodInvocationException]::new('Exception calling "Open"', [DiscoveryTestSqlError]::new(18456, 'Login failed'))
        }
        { Get-CMInventorySnapshot -Server 'CM01' -Database 'CM_MCM' } | Should -Throw -ExpectedMessage '*denied the Windows login*'
    }
}

Describe 'Find-DiscoveryCandidate' {
    BeforeAll {
        $script:Rows = @(
            , @('7-Zip 26.03', 'Igor Pavlov', '26.03', 'x86', 1)
            , @('7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64', 3)
            , @('Notepad++ (64-bit x64)', 'Notepad++ Team', '8.8', 'x64', 2)
            , @('Mozilla Firefox (x64 en-US)', 'Mozilla', '144.0', 'x64', 4)
            , @('Mozilla Maintenance Service', 'Mozilla', '144.0', 'x64', 4)
            , @('Audacity 3.7.9', 'Audacity Team', '3.7.9', 'x64', 1)
            , @('Java 8 Update 471', 'Oracle', '8.0.4710', 'x86', 1)
            , @('Java 8 Update 471 (64-bit)', 'Oracle', '8.0.4710', 'x64', 2)
        )
    }

    It 'merges 32-bit and 64-bit rows into one candidate with the summed client count' {
        $snap = New-Snapshot -Rows $script:Rows
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip *') -HiddenScript 'package-7zip.ps1' -DismissedScript @())
        $c.Count | Should -Be 1
        $c[0].Script | Should -Be 'package-7zip.ps1'
        $c[0].CMName | Should -Be '7-Zip'
        $c[0].ClientCount | Should -Be 4
        $c[0].MatchedName | Should -Be '7-Zip 26.03 (x64)'
    }

    It 'matches a wildcard case-insensitively' {
        $snap = New-Snapshot -Rows $script:Rows
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-npp.ps1' 'Notepad++' 'NOTEPAD++*') -HiddenScript 'package-npp.ps1')
        $c.Count | Should -Be 1
        $c[0].ClientCount | Should -Be 2
    }

    It 'matches any pattern of a multi-pattern ArpName' {
        $snap = New-Snapshot -Rows $script:Rows
        $p = New-Packager 'package-java8.ps1' 'Java 8' 'Java 8 Update * (64-bit) | Java 8 Update ?*'
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript 'package-java8.ps1')
        $c.Count | Should -Be 1
        $c[0].ClientCount | Should -Be 3
    }

    It 'counts a row once when more than one pattern matches it' {
        $snap = New-Snapshot -Rows $script:Rows
        $p = New-Packager 'package-ff.ps1' 'Firefox' 'Mozilla Firefox* | Mozilla Firefox (x64*'
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript 'package-ff.ps1')
        $c[0].ClientCount | Should -Be 4
    }

    It 'skips an ArpName of none or an empty ArpName' {
        $snap = New-Snapshot -Rows $script:Rows
        $pk = @(
            (New-Packager 'package-portable.ps1' 'Portable' 'none')
            (New-Packager 'package-portable2.ps1' 'Portable2' ' NONE ')
            (New-Packager 'package-blank.ps1' 'Blank' '')
            [pscustomobject]@{ Script = 'package-noprop.ps1'; Vendor = 'V'; App = 'A'; CMName = 'NoProp' }
        )
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager $pk -HiddenScript @('package-portable.ps1', 'package-portable2.ps1', 'package-blank.ps1', 'package-noprop.ps1')).Count | Should -Be 0
    }

    It 'skips a dismissed packager' {
        $snap = New-Snapshot -Rows $script:Rows
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*') -HiddenScript 'package-7zip.ps1' -DismissedScript 'PACKAGE-7ZIP.PS1').Count | Should -Be 0
    }

    It 'skips a packager that is not hidden' {
        $snap = New-Snapshot -Rows $script:Rows
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*') -HiddenScript 'package-other.ps1').Count | Should -Be 0
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*') -HiddenScript @()).Count | Should -Be 0
    }

    It 'skips a packager whose CMName is already a ConfigMgr Application, ignoring case' {
        $snap = New-Snapshot -Rows $script:Rows -Applications @('audacity 3.7.9', 'CMTrace Open')
        $p = New-Packager 'package-audacity.ps1' 'Audacity 3.7.9' 'Audacity*'
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript 'package-audacity.ps1').Count | Should -Be 0
    }

    It 'keeps a packager whose CMName only partly matches an Application name' {
        $snap = New-Snapshot -Rows $script:Rows -Applications @('7-Zip 26.03 (x64 edition)')
        $p = New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*'
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript 'package-7zip.ps1').Count | Should -Be 1
    }

    It 'returns nothing for zero inventory rows' {
        $snap = New-Snapshot -Rows @() -ClientTotal 0 -InventoryNewest $null
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '*') -HiddenScript 'package-7zip.ps1').Count | Should -Be 0
    }

    It 'skips a hidden packager with no matching inventory row' {
        $snap = New-Snapshot -Rows $script:Rows
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-vlc.ps1' 'VLC' 'VLC media player*') -HiddenScript 'package-vlc.ps1').Count | Should -Be 0
    }

    It 'does not throw on a malformed wildcard and matches it as literal text' {
        $snap = New-Snapshot -Rows @(, @('Odd [Name', 'P', '1', 'x86', 2))
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-odd.ps1' 'Odd' 'Odd [Name') -HiddenScript 'package-odd.ps1')
        $c.Count | Should -Be 1
    }

    It 'returns only the contract properties, with no machine names' {
        $snap = New-Snapshot -Rows $script:Rows
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*' -Vendor 'Igor Pavlov') -HiddenScript 'package-7zip.ps1')
        @($c[0].PSObject.Properties.Name) | Should -Be @('Script', 'CMName', 'Vendor', 'ClientCount', 'MatchedName')
        $c[0].Vendor | Should -Be 'Igor Pavlov'
    }

    It 'sorts candidates by CMName' {
        $snap = New-Snapshot -Rows $script:Rows
        $pk = @(
            (New-Packager 'package-npp.ps1' 'Notepad++' 'Notepad++*')
            (New-Packager 'package-7zip.ps1' '7-Zip' '7-Zip*')
            (New-Packager 'package-ff.ps1' 'Firefox' 'Mozilla Firefox*')
        )
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager $pk -HiddenScript @('package-npp.ps1', 'package-7zip.ps1', 'package-ff.ps1'))
        @($c.CMName) | Should -Be @('7-Zip', 'Firefox', 'Notepad++')
    }
}

Describe 'Get-DiscoveryStatusText' {
    It 'reports no inventory data for zero rows' {
        $snap = New-Snapshot -Rows @() -ClientTotal 0 -InventoryNewest $null
        $t = Get-DiscoveryStatusText -Snapshot $snap -CandidateCount 0
        $t | Should -Match '^No inventory data'
        $t | Should -Not -Match 'Discovered'
    }

    It 'states the client total, inventory age and candidate count' {
        $snap = New-Snapshot -Rows @(, @('A', 'P', '1', 'x86', 1)) -ClientTotal 4 -InventoryNewest (Get-Date '2026-10-15T02:55:30')
        $t = Get-DiscoveryStatusText -Snapshot $snap -CandidateCount 3 -Now (Get-Date '2026-10-17T12:00:00')
        $t | Should -Be 'Inventory covers 4 clients; newest data is 2 days old. Discovered 3 apps.'
    }

    It 'uses singular words and today for fresh data' {
        $snap = New-Snapshot -Rows @(, @('A', 'P', '1', 'x86', 1)) -ClientTotal 1 -InventoryNewest (Get-Date '2026-10-17T08:00:00')
        Get-DiscoveryStatusText -Snapshot $snap -CandidateCount 1 -Now (Get-Date '2026-10-17T12:00:00') |
            Should -Be 'Inventory covers 1 client; newest data is from today. Discovered 1 app.'
        $snap.InventoryNewest = Get-Date '2026-10-16T08:00:00'
        Get-DiscoveryStatusText -Snapshot $snap -CandidateCount 0 -Now (Get-Date '2026-10-17T12:00:00') |
            Should -Be 'Inventory covers 1 client; newest data is 1 day old. Discovered 0 apps.'
    }

    It 'states an unknown age when the newest time is missing' {
        $snap = New-Snapshot -Rows @(, @('A', 'P', '1', 'x86', 1)) -ClientTotal 2 -InventoryNewest $null
        Get-DiscoveryStatusText -Snapshot $snap -CandidateCount 0 | Should -Be 'Inventory covers 2 clients; age of newest data is unknown. Discovered 0 apps.'
    }
}
