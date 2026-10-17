#Requires -Version 5.1
<#
    The site database views are outside ConfigMgr role-based administration.
    No function returns a client machine name or a ResourceID: client
    identity leaves Get-CMInventorySnapshot only as per-snapshot ordinals.
#>

Set-StrictMode -Version 2.0

$script:DiscoveryInventoryQuery = @'
SET NOCOUNT ON;

SELECT DISTINCT arp.ResourceID, arp.DisplayName0 AS DisplayName, arp.Publisher0 AS Publisher,
       arp.Version0 AS Version, @Arch32 AS Architecture
FROM dbo.v_GS_ADD_REMOVE_PROGRAMS AS arp
JOIN dbo.v_R_System AS sys ON sys.ResourceID = arp.ResourceID
WHERE arp.DisplayName0 IS NOT NULL AND arp.DisplayName0 <> N''
  AND ISNULL(sys.Obsolete0, 0) = 0 AND ISNULL(sys.Decommissioned0, 0) = 0
UNION ALL
SELECT DISTINCT arp64.ResourceID, arp64.DisplayName0, arp64.Publisher0,
       arp64.Version0, @Arch64
FROM dbo.v_GS_ADD_REMOVE_PROGRAMS_64 AS arp64
JOIN dbo.v_R_System AS sys64 ON sys64.ResourceID = arp64.ResourceID
WHERE arp64.DisplayName0 IS NOT NULL AND arp64.DisplayName0 <> N''
  AND ISNULL(sys64.Obsolete0, 0) = 0 AND ISNULL(sys64.Decommissioned0, 0) = 0;

SELECT COUNT(DISTINCT inv.ResourceID) AS ClientTotal, MAX(inv.TimeStamp) AS InventoryNewest
FROM (
    SELECT ResourceID, TimeStamp FROM dbo.v_GS_ADD_REMOVE_PROGRAMS
    UNION ALL
    SELECT ResourceID, TimeStamp FROM dbo.v_GS_ADD_REMOVE_PROGRAMS_64
) AS inv
JOIN dbo.v_R_System AS invsys ON invsys.ResourceID = inv.ResourceID
WHERE ISNULL(invsys.Obsolete0, 0) = 0 AND ISNULL(invsys.Decommissioned0, 0) = 0;

SELECT DISTINCT app.DisplayName
FROM dbo.v_Applications AS app
JOIN dbo.v_ConfigurationItems AS ci ON ci.ModelId = app.ModelId AND ci.IsLatest = 1
WHERE app.ObjectTypeID = @AppObjectType AND ci.IsExpired = 0 AND ci.IsTombstoned = 0;
'@

$script:SqlLoginErrors   = @(18452, 18456, 18470, 18486, 18487, 18488)
$script:SqlDatabaseErrors = @(911, 4060)
$script:SqlPermissionErrors = @(229, 230)
$script:SqlNetworkErrors = @(-1, 2, 26, 40, 53, 64, 121, 1225, 10053, 10054, 10060, 10061, 11001, 11004)
$script:SqlTimeoutErrors = @(-2)

# ---------------------------------------------------------------------------
# Target resolution
# ---------------------------------------------------------------------------

function Get-DiscoveryProviderMachineName {
    <#
    .SYNOPSIS
        The SMS Provider of the current session: the SuiteCommon connection
        first, then the AppPackagerCommon provider resolution.
    #>
    $info = Get-Command -Name 'Get-CMConnectionInfo' -ErrorAction SilentlyContinue
    if ($info) {
        $connection = & $info
        if ($connection -and -not [string]::IsNullOrWhiteSpace([string]$connection.SMSProvider)) {
            return ([string]$connection.SMSProvider).Trim()
        }
    }
    $resolver = Get-Command -Name 'Resolve-CMProviderMachineName' -ErrorAction SilentlyContinue
    if ($resolver) {
        $provider = & $resolver
        if (-not [string]::IsNullOrWhiteSpace([string]$provider)) { return ([string]$provider).Trim() }
    }
    return $null
}

function Get-DiscoverySiteDefinition {
    <#
    .SYNOPSIS
        Reads SMS_SCI_SiteDefinition (SQLServerName, SQLDatabaseName) from
        the SMS Provider.
    #>
    param(
        [Parameter(Mandatory)][string]$ProviderMachineName,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9]{3}$')][string]$SiteCode
    )
    $splat = @{
        ComputerName        = $ProviderMachineName
        Namespace           = "root\sms\site_$SiteCode"
        ClassName           = 'SMS_SCI_SiteDefinition'
        Filter              = "SiteCode='$SiteCode'"
        OperationTimeoutSec = 30
        ErrorAction         = 'Stop'
    }
    return (Get-CimInstance @splat | Select-Object -First 1)
}

function Resolve-CMDatabaseTarget {
    <#
    .SYNOPSIS
        Returns the SQL Server and database of a site: overrides first, then
        the site definition on the SMS Provider, with CM_<SiteCode> as the
        database default. Throws a plain message when no server is known.
    #>
    param(
        [AllowEmptyString()][AllowNull()][ValidatePattern('^(?:[A-Za-z0-9]{3})?$')][string]$SiteCode = '',
        [string]$ServerOverride = '',
        [string]$DatabaseOverride = '',
        [string]$ProviderMachineName = ''
    )

    $code = ([string]$SiteCode).ToUpperInvariant()
    $server = ([string]$ServerOverride).Trim()
    $database = ([string]$DatabaseOverride).Trim()
    $lookupError = ''

    if ($server -and $database) {
        return [pscustomobject]@{ Server = $server; Database = $database }
    }
    if (-not $code) {
        throw 'A site code or SQL overrides are required. Set the site code, or the SQL server and database, in Options, ConfigMgr Preferences.'
    }

    if (-not $server -or -not $database) {
        $provider = ([string]$ProviderMachineName).Trim()
        if (-not $provider) { $provider = [string](Get-DiscoveryProviderMachineName) }

        $definition = $null
        if ($provider) {
            try { $definition = Get-DiscoverySiteDefinition -ProviderMachineName $provider -SiteCode $code }
            catch { $lookupError = "the SMS Provider $provider did not return the site definition" }
        }
        else {
            $lookupError = 'no SMS Provider is configured'
        }

        if ($definition) {
            $siteServer = ([string]$definition.SQLServerName).Trim()
            $siteDatabase = ([string]$definition.SQLDatabaseName).Trim()
            # A named instance is stored as INSTANCE\CM_<code> in SQLDatabaseName.
            if ($siteDatabase -match '^(?<instance>[^\\]+)\\(?<db>[^\\]+)$') {
                $siteDatabase = $Matches['db']
                if ($siteServer -and $Matches['instance'] -ne 'MSSQLSERVER') { $siteServer = "$siteServer\$($Matches['instance'])" }
            }
            if (-not $siteServer) { $lookupError = 'the site definition has no SQL Server name' }
            if (-not $server) { $server = $siteServer }
            if (-not $database) { $database = $siteDatabase }
        }
        elseif ($provider -and -not $lookupError) {
            $lookupError = "the SMS Provider $provider has no site definition for $code"
        }
    }

    if (-not $database) { $database = "CM_$code" }
    if (-not $server) {
        throw "The SQL Server for site $code is not known ($lookupError). Enter the SQL Server name in Preferences."
    }

    return [pscustomobject]@{ Server = $server; Database = $database }
}

# ---------------------------------------------------------------------------
# SQL access
# ---------------------------------------------------------------------------

function New-DiscoveryConnectionString {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Database,
        [int]$TimeoutSeconds = 30
    )
    # The builder is an IDictionary: PowerShell resolves $builder.DataSource as a
    # dictionary key, which the builder rejects. Keywords go through the indexer.
    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $Server
    $builder['Initial Catalog'] = $Database
    $builder['Integrated Security'] = $true
    $builder['ApplicationIntent'] = 'ReadOnly'
    $builder['Application Name'] = 'AppPackager Discovery'
    $builder['Connect Timeout'] = [Math]::Max(1, $TimeoutSeconds)
    return $builder.ConnectionString
}

function New-DiscoverySqlCommand {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$CommandText,
        [System.Collections.IDictionary]$Parameters = @{},
        [int]$TimeoutSeconds = 30
    )
    $command = $Connection.CreateCommand()
    $command.CommandText = $CommandText
    $command.CommandType = [System.Data.CommandType]::Text
    $command.CommandTimeout = [Math]::Max(1, $TimeoutSeconds)
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        $name = '@' + ([string]$key).TrimStart('@')
        if ($value -is [int]) {
            $p = $command.Parameters.Add($name, [System.Data.SqlDbType]::Int)
        }
        elseif ($value -is [string]) {
            $p = $command.Parameters.Add($name, [System.Data.SqlDbType]::NVarChar, [Math]::Max(1, $value.Length))
        }
        else {
            $command.Dispose()
            throw "Parameter $name has an unsupported type."
        }
        $p.Value = $value
    }
    return $command
}

function Invoke-DiscoverySqlQuery {
    <#
    .SYNOPSIS
        Runs one command and returns every result set as a DataSet. This is
        the only function that opens a SQL connection.
    #>
    param(
        [Parameter(Mandatory)][string]$ConnectionString,
        [Parameter(Mandatory)][string]$CommandText,
        [System.Collections.IDictionary]$Parameters = @{},
        [int]$TimeoutSeconds = 30
    )
    $connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    $command = $null
    try {
        $connection.Open()
        $command = New-DiscoverySqlCommand -Connection $connection -CommandText $CommandText -Parameters $Parameters -TimeoutSeconds $TimeoutSeconds
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        $dataSet = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataSet)
        $adapter.Dispose()
        return , $dataSet
    }
    finally {
        if ($command) { $command.Dispose() }
        $connection.Dispose()
    }
}

function Get-DiscoverySqlErrorInfo {
    param([System.Exception]$Exception)
    $numbers = New-Object System.Collections.Generic.List[int]
    $messages = New-Object System.Collections.Generic.List[string]
    $current = $Exception
    while ($current) {
        $messages.Add([string]$current.Message)
        $errorsProperty = $current.PSObject.Properties['Errors']
        $numberProperty = $current.PSObject.Properties['Number']
        if ($errorsProperty -and $errorsProperty.Value -is [System.Collections.IEnumerable]) {
            foreach ($e in $errorsProperty.Value) {
                $n = $e.PSObject.Properties['Number']
                if ($n) { $numbers.Add([int]$n.Value) }
            }
        }
        if ($numberProperty -and $numbers.Count -eq 0) { $numbers.Add([int]$numberProperty.Value) }
        $current = $current.InnerException
    }
    return [pscustomobject]@{ Numbers = @($numbers); Message = ($messages -join ' ') }
}

function ConvertTo-DiscoverySqlMessage {
    param(
        [Parameter(Mandatory)][System.Exception]$Exception,
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Database,
        [int]$TimeoutSeconds = 30
    )
    $info = Get-DiscoverySqlErrorInfo -Exception $Exception
    $has = { param($set) @($info.Numbers | Where-Object { $set -contains $_ }).Count -gt 0 }

    if (& $has $script:SqlDatabaseErrors) {
        return "Database $Database does not exist on SQL Server $Server, or the Windows login cannot open it. Check the database name."
    }
    if (& $has $script:SqlLoginErrors) {
        if ($info.Message -match 'ANONYMOUS LOGON') {
            return "SQL Server on $Server received an anonymous login. Windows credentials do not pass through a second remoting hop without Kerberos delegation. Run the query from a session that holds your own logon."
        }
        return "SQL Server on $Server denied the Windows login. The account needs the db_datareader role on $Database."
    }
    if ($info.Message -match 'SSPI') {
        return "Windows authentication to SQL Server $Server failed. Run the app as a domain account on a domain-joined computer, and check the SQL Server service principal name."
    }
    if (& $has $script:SqlPermissionErrors) {
        return "The Windows login cannot read the inventory views in $Database on $Server. The account needs the db_datareader role on $Database."
    }
    if (& $has $script:SqlTimeoutErrors) {
        return "SQL Server $Server did not answer within $TimeoutSeconds seconds."
    }
    if (& $has $script:SqlNetworkErrors) {
        return "Cannot reach SQL Server $Server. Check the server name, the network path and the SQL Server firewall port."
    }
    $first = ([string]$Exception.Message -split "`r?`n")[0]
    return "The site database query on $Server failed: $first"
}

function Get-DiscoveryCellValue {
    param($Row, [string]$Column)
    $value = $Row[$Column]
    if ($value -is [System.DBNull]) { return $null }
    return $value
}

function Get-CMInventorySnapshot {
    <#
    .SYNOPSIS
        Reads Add or Remove Programs inventory (32-bit and 64-bit views),
        the client total, the newest inventory time and the current
        ConfigMgr Application names. Throws a plain message on a denied
        login, an unreachable server or a wrong database.
    #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Database,
        [ValidateRange(1, 600)][int]$TimeoutSeconds = 30
    )

    $connectionString = New-DiscoveryConnectionString -Server $Server -Database $Database -TimeoutSeconds $TimeoutSeconds
    $parameters = [ordered]@{ Arch32 = 'x86'; Arch64 = 'x64'; AppObjectType = 31 }

    try {
        $dataSet = Invoke-DiscoverySqlQuery -ConnectionString $connectionString -CommandText $script:DiscoveryInventoryQuery -Parameters $parameters -TimeoutSeconds $TimeoutSeconds
    }
    catch {
        $message = ConvertTo-DiscoverySqlMessage -Exception $_.Exception -Server $Server -Database $Database -TimeoutSeconds $TimeoutSeconds
        throw (New-Object System.InvalidOperationException -ArgumentList $message, $_.Exception)
    }

    if (-not $dataSet -or $dataSet.Tables.Count -lt 3) {
        throw "The site database query on $Server returned an unexpected result."
    }

    $ordinalByResource = New-Object 'System.Collections.Generic.Dictionary[long,int]'
    $groups = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([System.StringComparer]::Ordinal)
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($r in $dataSet.Tables[0].Rows) {
        $resource = Get-DiscoveryCellValue $r 'ResourceID'
        if ($null -eq $resource) { continue }
        $resourceId = [long]$resource
        $ordinal = 0
        if (-not $ordinalByResource.TryGetValue($resourceId, [ref]$ordinal)) {
            $ordinal = $ordinalByResource.Count
            $ordinalByResource[$resourceId] = $ordinal
        }
        $name = [string](Get-DiscoveryCellValue $r 'DisplayName')
        $publisher = [string](Get-DiscoveryCellValue $r 'Publisher')
        $version = [string](Get-DiscoveryCellValue $r 'Version')
        $architecture = [string](Get-DiscoveryCellValue $r 'Architecture')
        $key = $name + [char]0 + $publisher + [char]0 + $version + [char]0 + $architecture
        $group = $null
        if (-not $groups.TryGetValue($key, [ref]$group)) {
            $group = [pscustomobject]@{
                DisplayName  = $name
                Publisher    = $publisher
                Version      = $version
                Architecture = $architecture
                Clients      = New-Object 'System.Collections.Generic.HashSet[int]'
            }
            $groups[$key] = $group
            $order.Add($key)
        }
        [void]$group.Clients.Add($ordinal)
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($key in $order) {
        $group = $groups[$key]
        $keys = New-Object int[] $group.Clients.Count
        $group.Clients.CopyTo($keys)
        $rows.Add([pscustomobject]@{
                DisplayName  = $group.DisplayName
                Publisher    = $group.Publisher
                Version      = $group.Version
                Architecture = $group.Architecture
                ClientCount  = $keys.Length
                ClientKeys   = $keys
            })
    }

    $clientTotal = 0
    $newest = $null
    if ($dataSet.Tables[1].Rows.Count -gt 0) {
        $summary = $dataSet.Tables[1].Rows[0]
        $clientTotal = [int](Get-DiscoveryCellValue $summary 'ClientTotal')
        $n = Get-DiscoveryCellValue $summary 'InventoryNewest'
        if ($null -ne $n) { $newest = [datetime]$n }
    }

    $applications = @(foreach ($r in $dataSet.Tables[2].Rows) {
            $name = [string](Get-DiscoveryCellValue $r 'DisplayName')
            if ($name) { $name }
        })

    return [pscustomobject]@{
        Rows            = $rows.ToArray()
        ClientTotal     = $clientTotal
        InventoryNewest = $newest
        Applications    = $applications
    }
}

# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

function Get-DiscoveryArpPatterns {
    param($ArpName, [string]$Script = '', [string]$Tag = 'ArpName')
    $options = [System.Management.Automation.WildcardOptions]::IgnoreCase -bor [System.Management.Automation.WildcardOptions]::CultureInvariant
    $patterns = New-Object System.Collections.Generic.List[System.Management.Automation.WildcardPattern]
    foreach ($value in @($ArpName)) {
        foreach ($part in ([string]$value -split '\|')) {
            $text = $part.Trim()
            if (-not $text -or $text -eq 'none') { continue }
            $pattern = New-Object System.Management.Automation.WildcardPattern -ArgumentList $text, $options
            # WildcardPattern parses on first use; an unclosed [ throws there.
            try { [void]$pattern.IsMatch('') }
            catch { throw "Packager $Script has a $Tag pattern that is not a valid wildcard: $text" }
            $patterns.Add($pattern)
        }
    }
    return , $patterns.ToArray()
}

function Get-DiscoveryPackagerValue {
    param($Packager, [string]$Name)
    $property = $Packager.PSObject.Properties[$Name]
    if (-not $property -or $null -eq $property.Value) { return '' }
    return ([string]$property.Value).Trim()
}

function Find-DiscoveryCandidate {
    <#
    .SYNOPSIS
        Returns the hidden, not dismissed packagers whose ArpName (and
        ArpVersion and ArpArch, when set) matches at least one inventory row
        and whose CMName is not a current ConfigMgr Application. ClientCount
        is the number of distinct clients over the matched rows. Throws when
        a pattern is not a valid wildcard.
    #>
    param(
        [Parameter(Mandatory)]$Snapshot,
        [AllowEmptyCollection()][object[]]$Packager = @(),
        [AllowEmptyCollection()][AllowNull()][string[]]$HiddenScript = @(),
        [AllowEmptyCollection()][AllowNull()][string[]]$DismissedScript = @()
    )

    $rows = @($Snapshot.Rows | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.DisplayName) })
    if ($rows.Count -eq 0) { return }

    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $hidden = New-Object System.Collections.Generic.HashSet[string] -ArgumentList $comparer
    foreach ($s in @($HiddenScript)) { if ($s) { [void]$hidden.Add($s.Trim()) } }
    $dismissed = New-Object System.Collections.Generic.HashSet[string] -ArgumentList $comparer
    foreach ($s in @($DismissedScript)) { if ($s) { [void]$dismissed.Add($s.Trim()) } }
    $applications = New-Object System.Collections.Generic.HashSet[string] -ArgumentList $comparer
    foreach ($a in @($Snapshot.Applications)) { if ($a) { [void]$applications.Add(([string]$a).Trim()) } }

    $candidates = foreach ($p in @($Packager)) {
        if (-not $p) { continue }
        $scriptName = ([string]$p.Script).Trim()
        if (-not $scriptName -or -not $hidden.Contains($scriptName) -or $dismissed.Contains($scriptName)) { continue }

        $cmName = ([string]$p.CMName).Trim()
        if ($cmName -and $applications.Contains($cmName)) { continue }

        $patterns = Get-DiscoveryArpPatterns -ArpName (Get-DiscoveryPackagerValue $p 'ArpName') -Script $scriptName
        if ($patterns.Count -eq 0) { continue }
        $versionPatterns = Get-DiscoveryArpPatterns -ArpName (Get-DiscoveryPackagerValue $p 'ArpVersion') -Script $scriptName -Tag 'ArpVersion'
        $arch = Get-DiscoveryPackagerValue $p 'ArpArch'

        $clients = New-Object 'System.Collections.Generic.HashSet[int]'
        $best = $null
        foreach ($row in $rows) {
            if ($arch -and -not [string]::Equals($arch, [string]$row.Architecture, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($versionPatterns.Count -gt 0) {
                $version = [string]$row.Version
                $versionHit = $false
                foreach ($pattern in $versionPatterns) { if ($pattern.IsMatch($version)) { $versionHit = $true; break } }
                if (-not $versionHit) { continue }
            }
            $name = [string]$row.DisplayName
            $hit = $false
            foreach ($pattern in $patterns) { if ($pattern.IsMatch($name)) { $hit = $true; break } }
            if (-not $hit) { continue }
            foreach ($k in @($row.ClientKeys)) { [void]$clients.Add([int]$k) }
            if (-not $best -or [int]$row.ClientCount -gt [int]$best.ClientCount) { $best = $row }
        }
        if (-not $best) { continue }

        [pscustomobject]@{
            Script      = $scriptName
            CMName      = $cmName
            Vendor      = [string]$p.Vendor
            ClientCount = $clients.Count
            MatchedName = [string]$best.DisplayName
        }
    }

    if ($candidates) { @($candidates) | Sort-Object -Property CMName }
}

# ---------------------------------------------------------------------------
# Status text
# ---------------------------------------------------------------------------

function Get-DiscoveryStatusText {
    <#
    .SYNOPSIS
        One line for the log and the prompt: client total, inventory age and
        candidate count, or the no-inventory-data text for zero rows.
    #>
    param(
        [Parameter(Mandatory)]$Snapshot,
        [Parameter(Mandatory)][int]$CandidateCount,
        [datetime]$Now = (Get-Date)
    )

    if (@($Snapshot.Rows).Count -eq 0) {
        return 'No inventory data: the site database has no Add or Remove Programs rows. Check that hardware inventory collects Add or Remove Programs.'
    }

    $clients = [int]$Snapshot.ClientTotal
    $clientWord = if ($clients -eq 1) { 'client' } else { 'clients' }
    $age = 'age of newest data is unknown'
    if ($null -ne $Snapshot.InventoryNewest) {
        $days = [int](($Now.Date - ([datetime]$Snapshot.InventoryNewest).Date).TotalDays)
        $age = if ($days -le 0) { 'newest data is from today' }
        elseif ($days -eq 1) { 'newest data is 1 day old' }
        else { "newest data is $days days old" }
    }
    $appWord = if ($CandidateCount -eq 1) { 'app' } else { 'apps' }
    return "Inventory covers $clients $clientWord; $age. Discovered $CandidateCount $appWord."
}

Export-ModuleMember -Function @(
    'Resolve-CMDatabaseTarget', 'Get-CMInventorySnapshot', 'Find-DiscoveryCandidate', 'Get-DiscoveryStatusText'
)
