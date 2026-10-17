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

    # One fixture row per (ResourceID, DisplayName, Publisher, Version, Architecture),
    # the shape of the first result set.
    function New-ArpTable {
        param([object[]]$Rows)
        $t = New-Object System.Data.DataTable
        [void]$t.Columns.Add('ResourceID', [int])
        foreach ($c in 'DisplayName', 'Publisher', 'Version', 'Architecture') { [void]$t.Columns.Add($c, [string]) }
        foreach ($r in $Rows) {
            $row = $t.NewRow()
            $row.ResourceID = $r[0]
            $row.DisplayName = $r[1]
            $row.Publisher = if ($null -eq $r[2]) { [DBNull]::Value } else { $r[2] }
            $row.Version = if ($null -eq $r[3]) { [DBNull]::Value } else { $r[3] }
            $row.Architecture = $r[4]
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
        # Builds the snapshot through Get-CMInventorySnapshot from fixture SQL
        # rows. The fifth element is either explicit ResourceIDs (int[]) or a
        # count that receives ResourceIDs no other row shares.
        $nextId = 1000
        $script:FixtureArpRows = @(foreach ($r in @($Rows)) {
                $ids = if ($r[4] -is [array]) { [int[]]$r[4] } else { [int[]]@(for ($i = 0; $i -lt [int]$r[4]; $i++) { ($nextId++) }) }
                foreach ($id in $ids) { , @($id, $r[0], $r[1], $r[2], $r[3]) }
            })
        $script:FixtureClientTotal = $ClientTotal
        $script:FixtureNewest = if ($null -eq $InventoryNewest) { [DBNull]::Value } else { $InventoryNewest }
        $script:FixtureApplications = $Applications
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows $script:FixtureArpRows -ClientTotal $script:FixtureClientTotal -Newest $script:FixtureNewest -Applications $script:FixtureApplications
        }
        Get-CMInventorySnapshot -Server 's' -Database 'd'
    }

    function Get-AllValueText {
        # Every property value of an object graph, as text, depth first.
        param($Value, [int]$Depth = 0)
        if ($null -eq $Value -or $Depth -gt 6) { return }
        if ($Value -is [string] -or $Value -is [ValueType]) { return [string]$Value }
        if ($Value -is [System.Collections.IEnumerable]) {
            foreach ($item in $Value) { Get-AllValueText -Value $item -Depth ($Depth + 1) }
            return
        }
        foreach ($p in $Value.PSObject.Properties) {
            "@name:$($p.Name)"
            Get-AllValueText -Value $p.Value -Depth ($Depth + 1)
        }
    }

    function New-Packager {
        param([string]$Script, [string]$CMName, [string]$ArpName, [string]$Vendor = 'Vendor', [string]$App = 'App', [string]$ArpVersion = '', [string]$ArpArch = '')
        [pscustomobject]@{ Script = $Script; Vendor = $Vendor; App = $App; CMName = $CMName; ArpName = $ArpName; ArpVersion = $ArpVersion; ArpArch = $ArpArch }
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

    It 'throws a plain message without the raw provider error when no server is known' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'RPC server unavailable (HRESULT 0x800706BA)' }
        $message = $null
        try { Resolve-CMDatabaseTarget -SiteCode 'MCM' -ProviderMachineName 'cm01' } catch { $message = $_.Exception.Message }
        $message | Should -BeLike '*SQL Server for site MCM*SMS Provider cm01*'
        $message | Should -Not -BeLike '*RPC server*'
        $message | Should -Not -BeLike '*HRESULT*'
    }

    It 'uses both overrides without a site code' {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'must not be called' }
        InModuleScope AppPackagerDiscovery { Mock Get-DiscoveryProviderMachineName { throw 'must not be called' } }
        $t = Resolve-CMDatabaseTarget -SiteCode '' -ServerOverride 'sql01' -DatabaseOverride 'CM_ALT'
        $t.Server | Should -Be 'sql01'
        $t.Database | Should -Be 'CM_ALT'
        $t = Resolve-CMDatabaseTarget -ServerOverride 'sql01' -DatabaseOverride 'CM_ALT'
        $t.Database | Should -Be 'CM_ALT'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 0
    }

    It 'throws a plain message with no site code and only one override' -ForEach @(
        @{ Server = 'sql01'; Database = '' }
        @{ Server = ''; Database = 'CM_ALT' }
        @{ Server = ''; Database = '' }
    ) {
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition { throw 'must not be called' }
        { Resolve-CMDatabaseTarget -SiteCode '' -ServerOverride $Server -DatabaseOverride $Database } |
            Should -Throw -ExpectedMessage '*site code or SQL overrides are required*'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 0
    }

    It 'finds the SMS Provider of the session when no provider is passed' {
        InModuleScope AppPackagerDiscovery { Mock Get-DiscoveryProviderMachineName { 'cm01.contoso.com' } }
        Mock -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition {
            [pscustomobject]@{ SQLServerName = 'SQL02.contoso.com'; SQLDatabaseName = 'CM_MCM' }
        }
        $t = Resolve-CMDatabaseTarget -SiteCode 'MCM'
        $t.Server | Should -Be 'SQL02.contoso.com'
        Should -Invoke -ModuleName AppPackagerDiscovery Get-DiscoverySiteDefinition -Times 1 -ParameterFilter { $ProviderMachineName -eq 'cm01.contoso.com' }
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
        $script:Captured | Should -Not -Match '(?i)\bName0\b|Netbios|ResourceName|MachineName|User_Name'
    }

    It 'returns ResourceIDs to nobody: rows carry counts only' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows @(
                , @(16777301, 'Git', 'Git', '2.51', 'x64')
                , @(16777305, 'Git', 'Git', '2.51', 'x64')
            ) -ClientTotal 2
        }
        $s = Get-CMInventorySnapshot -Server 's' -Database 'd'
        $text = ($s | ConvertTo-Json -Depth 6)
        $text | Should -Not -Match '1677730[15]'
        @($s.Rows).Count | Should -Be 1
        $s.Rows[0].ClientCount | Should -Be 2
    }

    It 'returns no per-client key in any property of the snapshot or the candidates' {
        # Client 1 has two entries that the packager matches; client 2 has one.
        $snap = New-Snapshot -Rows @(
            , @('Java 8 Update 471', 'Oracle', '8.0.4710', 'x86', @(1, 2))
            , @('Java 8 Update 471 (64-bit)', 'Oracle', '8.0.4710', 'x64', @(1))
            , @('Notepad++ (64-bit x64)', 'Notepad++ Team', '8.8', 'x64', @(2))
        ) -ClientTotal 2
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-java8.ps1' 'Java 8' 'Java 8 Update *') -HiddenScript 'package-java8.ps1')
        $c.Count | Should -Be 1
        $c[0].ClientCount | Should -Be 2

        foreach ($object in @($snap) + @($snap.Rows) + $c) {
            $values = @(Get-AllValueText -Value $object)
            $values | Should -Not -Contain '@name:ClientKeys'
            $values | Should -Not -Contain '@name:ResourceID'
            # Every number in the output is a count or a version, never a list of ordinals per row.
            foreach ($p in $object.PSObject.Properties) {
                if ($p.Value -is [System.Collections.IEnumerable] -and $p.Value -isnot [string]) {
                    @($p.Value | Where-Object { $_ -is [int] -or $_ -is [long] }).Count | Should -Be 0 -Because "property $($p.Name) must not list client keys"
                }
            }
        }
        @($snap.Rows[0].PSObject.Properties.Name) | Should -Be @('DisplayName', 'Publisher', 'Version', 'Architecture', 'ClientCount')
    }

    It 'counts a client once per product row when the query returns the same client twice' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows @(
                , @(1, 'Git', 'Git', '2.51', 'x64')
                , @(1, 'Git', 'Git', '2.51', 'x64')
                , @(2, 'Git', 'Git', '2.51', 'x64')
            )
        }
        $s = Get-CMInventorySnapshot -Server 's' -Database 'd'
        $s.Rows[0].ClientCount | Should -Be 2
    }

    It 'excludes obsolete and decommissioned clients from rows, client total and newest time' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            $script:Captured = $CommandText
            New-SnapshotDataSet
        }
        $null = Get-CMInventorySnapshot -Server 's' -Database 'd'
        $statements = $script:Captured -split ';' | Where-Object { $_ -match '(?i)v_GS_ADD_REMOVE_PROGRAMS' }
        @($statements).Count | Should -Be 2
        foreach ($statement in $statements) {
            ([regex]::Matches($statement, '(?i)JOIN dbo\.v_R_System')).Count | Should -BeGreaterOrEqual 1
            ([regex]::Matches($statement, '(?i)ISNULL\(\w+\.Obsolete0, 0\) = 0')).Count | Should -Be ([regex]::Matches($statement, '(?i)JOIN dbo\.v_R_System')).Count
            ([regex]::Matches($statement, '(?i)ISNULL\(\w+\.Decommissioned0, 0\) = 0')).Count | Should -Be ([regex]::Matches($statement, '(?i)JOIN dbo\.v_R_System')).Count
        }
    }

    It 'merges 32-bit and 64-bit rows and maps the summary' {
        Mock -ModuleName AppPackagerDiscovery Invoke-DiscoverySqlQuery {
            New-SnapshotDataSet -ArpRows @(
                , @(1, '7-Zip 26.03', 'Igor Pavlov', '26.03', 'x86')
                , @(2, '7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64')
                , @(3, '7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64')
                , @(4, '7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64')
                , @(2, 'Notepad++ (64-bit x64)', $null, $null, 'x64')
                , @(3, 'Notepad++ (64-bit x64)', $null, $null, 'x64')
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
            New-SnapshotDataSet -ArpRows @(, @(1, 'A', 'P', '1', 'x86')) -ClientTotal 1 -Newest (Get-Date)
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

    It 'skips <Script> when ConfigMgr has the packaged title <Title>' -ForEach @(
        @{ Script = 'package-m365apps-x64.ps1'; CMName = 'M365 Apps for Enterprise'; Arp = 'Microsoft 365 Apps for enterprise - *'; Title = 'M365 Apps for Enterprise (x64) (MEC)' }
        @{ Script = 'package-m365apps-x64.ps1'; CMName = 'M365 Apps for Enterprise'; Arp = 'Microsoft 365 Apps for enterprise - *'; Title = 'M365 Apps for Enterprise - 16.0.19127.20264 (x64) (MEC)' }
        @{ Script = 'package-audacity.ps1'; CMName = 'Audacity'; Arp = 'Audacity [0-9]*'; Title = 'Audacity 3.7.9' }
        @{ Script = 'package-7zip.ps1'; CMName = '7-Zip'; Arp = '7-Zip [0-9]*'; Title = '7-Zip' }
        @{ Script = 'package-7zip.ps1'; CMName = '7-Zip'; Arp = '7-Zip [0-9]*'; Title = '7-zip 26.03 (x64)' }
        @{ Script = 'package-notepadplusplus.ps1'; CMName = 'Notepad++'; Arp = 'Notepad++ (64-bit x64)'; Title = 'Notepad++ (64-bit x64)' }
        @{ Script = 'package-firefox.ps1'; CMName = 'Mozilla Firefox'; Arp = 'Mozilla Firefox (x64 *)'; Title = 'Mozilla Firefox (en-US)' }
    ) {
        $snap = New-Snapshot -Rows @(
            , @('Microsoft 365 Apps for enterprise - en-us', 'Microsoft Corporation', '16.0.19127.20264', 'x64', 2)
            , @('Audacity 3.7.9', 'Audacity Team', '3.7.9', 'x64', 1)
            , @('7-Zip 26.03 (x64)', 'Igor Pavlov', '26.03', 'x64', 3)
            , @('Notepad++ (64-bit x64)', 'Notepad++ Team', '8.8', 'x64', 2)
            , @('Mozilla Firefox (x64 en-US)', 'Mozilla', '144.0', 'x64', 4)
        ) -Applications @('CMTrace Open', $Title)
        $p = New-Packager $Script $CMName $Arp
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript $Script).Count | Should -Be 0
    }

    It 'keeps <Script> when ConfigMgr only has another product that starts with the same words: <Title>' -ForEach @(
        @{ Script = 'package-git.ps1'; CMName = 'Git'; Arp = 'Git'; Row = 'Git'; Title = 'Git Extensions 4.3' }
        @{ Script = 'package-git.ps1'; CMName = 'Git'; Arp = 'Git'; Row = 'Git'; Title = 'GitHub Desktop' }
        @{ Script = 'package-firefox.ps1'; CMName = 'Mozilla Firefox'; Arp = 'Mozilla Firefox (x64 *)'; Row = 'Mozilla Firefox (x64 en-US)'; Title = 'Mozilla Firefox ESR (x64 en-US)' }
        @{ Script = 'package-m365apps-x64.ps1'; CMName = 'M365 Apps for Enterprise'; Arp = 'Microsoft 365 Apps for enterprise - *'; Row = 'Microsoft 365 Apps for enterprise - en-us'; Title = 'M365 Apps for Enterprise Project (x64)' }
        @{ Script = 'package-ms-openjdk17-exe.ps1'; CMName = 'Microsoft OpenJDK 17 (EXE)'; Arp = 'Microsoft Build of OpenJDK with Hotspot 17*'; Row = 'Microsoft Build of OpenJDK with Hotspot 17.0.16+8 (x64)'; Title = 'Microsoft OpenJDK 17 (MSI) (EXE)' }
        @{ Script = 'package-corretto-jdk11-x64.ps1'; CMName = 'Amazon Corretto 11 (x64)'; Arp = 'Amazon Corretto 11*'; Row = 'Amazon Corretto 11.0.28 (x64)'; Title = 'Amazon Corretto 11 (x86) (x64)' }
    ) {
        $snap = New-Snapshot -Rows @(, @($Row, 'P', '1.0', 'x64', 2)) -Applications @($Title)
        $p = New-Packager $Script $CMName $Arp
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript $Script)
        $c.Count | Should -Be 1
        $c[0].ClientCount | Should -Be 2
    }

    It 'returns nothing for zero inventory rows' {
        $snap = New-Snapshot -Rows @() -ClientTotal 0 -InventoryNewest $null
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-7zip.ps1' '7-Zip' '*') -HiddenScript 'package-7zip.ps1').Count | Should -Be 0
    }

    It 'skips a hidden packager with no matching inventory row' {
        $snap = New-Snapshot -Rows $script:Rows
        @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-vlc.ps1' 'VLC' 'VLC media player*') -HiddenScript 'package-vlc.ps1').Count | Should -Be 0
    }

    It 'throws a plain message that names the packager for a malformed <Tag> pattern <Pattern>' -ForEach @(
        @{ Tag = 'ArpName'; Pattern = 'Odd [Name' }
        @{ Tag = 'ArpName'; Pattern = '[' }
        @{ Tag = 'ArpName'; Pattern = 'Git | Odd[' }
        @{ Tag = 'ArpVersion'; Pattern = '17.[0' }
    ) {
        $snap = New-Snapshot -Rows @(, @('Odd [Name', 'P', '17.0', 'x86', 2))
        $p = if ($Tag -eq 'ArpName') { New-Packager 'package-odd.ps1' 'Odd' $Pattern } else { New-Packager 'package-odd.ps1' 'Odd' 'Odd*' -ArpVersion $Pattern }
        { Find-DiscoveryCandidate -Snapshot $snap -Packager @($p) -HiddenScript 'package-odd.ps1' } |
            Should -Throw -ExpectedMessage "Packager package-odd.ps1 has a $Tag pattern that is not a valid wildcard*"
    }

    It 'emits nothing, not one null, when no packager matches' {
        $snap = New-Snapshot -Rows $script:Rows
        $seen = New-Object System.Collections.ArrayList
        Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-vlc.ps1' 'VLC' 'VLC media player*') -HiddenScript 'package-vlc.ps1' |
            ForEach-Object { [void]$seen.Add($_) }
        $seen.Count | Should -Be 0
        $direct = Find-DiscoveryCandidate -Snapshot $snap -Packager @() -HiddenScript @()
        $direct.Count | Should -Be 0
        $null -eq $direct | Should -BeTrue
    }

    It 'counts a client once when it has two matching entries' {
        $snap = New-Snapshot -Rows @(
            , @('Java 8 Update 471', 'Oracle', '8.0.4710', 'x86', @(1, 2))
            , @('Java 8 Update 471 (64-bit)', 'Oracle', '8.0.4710', 'x64', @(2, 3))
            , @('Java 8 Update 461 (64-bit)', 'Oracle', '8.0.4610', 'x64', @(3))
        )
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-java8.ps1' 'Java 8' 'Java 8 Update *') -HiddenScript 'package-java8.ps1')
        $c[0].ClientCount | Should -Be 3
        $c[0].MatchedName | Should -Be 'Java 8 Update 471'
    }

    It 'never overstates the count for rows that Get-CMInventorySnapshot did not build' {
        $snap = [pscustomobject]@{
            Rows         = @(
                [pscustomobject]@{ DisplayName = 'Java 8 Update 471'; Publisher = 'Oracle'; Version = '8.0.4710'; Architecture = 'x86'; ClientCount = 2 }
                [pscustomobject]@{ DisplayName = 'Java 8 Update 471 (64-bit)'; Publisher = 'Oracle'; Version = '8.0.4710'; Architecture = 'x64'; ClientCount = 3 }
            )
            Applications = @()
        }
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-java8.ps1' 'Java 8' 'Java 8 Update *') -HiddenScript 'package-java8.ps1')
        $c[0].ClientCount | Should -Be 3
    }

    It 'matches with invariant culture rules' {
        $before = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('tr-TR')
            $snap = New-Snapshot -Rows @(, @('INKSCAPE 1.4', 'Inkscape', '1.4', 'x64', 1))
            @(Find-DiscoveryCandidate -Snapshot $snap -Packager @(New-Packager 'package-inkscape.ps1' 'Inkscape' 'inkscape*') -HiddenScript 'package-inkscape.ps1').Count | Should -Be 1
        }
        finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $before }
    }

    It 'limits matches to ArpVersion and ArpArch when the packager sets them' {
        $snap = New-Snapshot -Rows @(
            , @('Amazon Corretto (x64)', 'Amazon', '17.0.16.8', 'x64', @(1))
            , @('Amazon Corretto (x64)', 'Amazon', '21.0.8.9', 'x64', @(2, 3))
            , @('Amazon Corretto (x86)', 'Amazon', '17.0.16.8', 'x86', @(4))
        )
        $pk = @(
            (New-Packager 'package-c17.ps1' 'Corretto 17' 'Amazon Corretto (x??)' -ArpVersion '17.*' -ArpArch 'x64')
            (New-Packager 'package-c21.ps1' 'Corretto 21' 'Amazon Corretto (x64)' -ArpVersion '21.*')
            (New-Packager 'package-c25.ps1' 'Corretto 25' 'Amazon Corretto (x64)' -ArpVersion '25.*')
            (New-Packager 'package-c17x86.ps1' 'Corretto 17 x86' 'Amazon Corretto*' -ArpArch 'X86')
            (New-Packager 'package-carm.ps1' 'Corretto arm64' 'Amazon Corretto*' -ArpArch 'arm64')
        )
        $c = @(Find-DiscoveryCandidate -Snapshot $snap -Packager $pk -HiddenScript @($pk.Script))
        ($c | ForEach-Object { '{0}={1}' -f $_.Script, $_.ClientCount }) -join ',' |
            Should -Be 'package-c17.ps1=1,package-c17x86.ps1=1,package-c21.ps1=2'
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
