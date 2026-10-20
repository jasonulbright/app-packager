#Requires -Version 5.1
<#
    The site database views are outside ConfigMgr role-based administration.
    Discovery reads distinct software titles, never per-client inventory.
#>

Set-StrictMode -Version 2.0

$script:DiscoveryInventoryQuery = @'
SET NOCOUNT ON;

SELECT DISTINCT DisplayName0 AS DisplayName, Version0 AS Version, @Arch32 AS Architecture
FROM dbo.v_GS_ADD_REMOVE_PROGRAMS
WHERE DisplayName0 IS NOT NULL AND DisplayName0 <> N''
UNION ALL
SELECT DISTINCT DisplayName0, Version0, @Arch64
FROM dbo.v_GS_ADD_REMOVE_PROGRAMS_64
WHERE DisplayName0 IS NOT NULL AND DisplayName0 <> N'';

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

function New-DiscoveryCimSession {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][ValidateSet('Dcom', 'Wsman')][string]$Protocol
    )
    $splat = @{
        ComputerName        = $ComputerName
        SessionOption       = (New-CimSessionOption -Protocol $Protocol)
        OperationTimeoutSec = 30
        ErrorAction         = 'Stop'
    }
    return (New-CimSession @splat)
}

function Get-DiscoveryCimSiteDefinition {
    param(
        [Parameter(Mandatory)]$CimSession,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9]{3}$')][string]$SiteCode
    )
    $splat = @{
        CimSession          = $CimSession
        Namespace           = "root\sms\site_$SiteCode"
        ClassName           = 'SMS_SCI_SiteDefinition'
        Filter              = "SiteCode='$SiteCode'"
        OperationTimeoutSec = 30
        ErrorAction         = 'Stop'
    }
    return (Get-CimInstance @splat | Select-Object -First 1)
}

function Get-DiscoverySiteDefinition {
    <#
    .SYNOPSIS
        Reads SMS_SCI_SiteDefinition (SQLServerName, SQLDatabaseName) from
        the SMS Provider, over DCOM first and WinRM second.
    #>
    param(
        [Parameter(Mandatory)][string]$ProviderMachineName,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9]{3}$')][string]$SiteCode
    )
    # WinRM to the provider requires local Administrators or Remote Management
    # Users membership; a ConfigMgr administrator without either gets "Access
    # is denied" there and reaches the SMS Provider only through DCOM.
    foreach ($protocol in 'Dcom', 'Wsman') {
        $session = $null
        try {
            $session = New-DiscoveryCimSession -ComputerName $ProviderMachineName -Protocol $protocol
            return (Get-DiscoveryCimSiteDefinition -CimSession $session -SiteCode $SiteCode)
        }
        catch {
            Write-Verbose "Site definition lookup over $protocol failed."
        }
        finally {
            if ($null -ne $session) { $session.Dispose() }
        }
    }
    throw "The site definition could not be read from the SMS Provider $ProviderMachineName. Enter the SQL server and database in Options, ConfigMgr Preferences."
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
        [int]$TimeoutSeconds = 30,
        [hashtable]$OperationState
    )
    $connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    $command = $null
    $adapter = $null
    try {
        $command = New-DiscoverySqlCommand -Connection $connection -CommandText $CommandText -Parameters $Parameters -TimeoutSeconds $TimeoutSeconds
        if ($OperationState) {
            $OperationState.Command = $command
            $token = $OperationState.Cancellation.Token
            $token.ThrowIfCancellationRequested()
            $connection.OpenAsync($token).GetAwaiter().GetResult()
            $token.ThrowIfCancellationRequested()
        }
        else { $connection.Open() }
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        $dataSet = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataSet)
        return , $dataSet
    }
    finally {
        if ($OperationState) { $OperationState.Command = $null }
        if ($adapter) { $adapter.Dispose() }
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
        Reads distinct Add or Remove Programs titles (32-bit and 64-bit
        views) and current ConfigMgr Application names. Throws a plain message
        on a denied login, an unreachable server or a wrong database.
    #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Database,
        [ValidateRange(1, 600)][int]$TimeoutSeconds = 30,
        [hashtable]$OperationState
    )

    $connectionString = New-DiscoveryConnectionString -Server $Server -Database $Database -TimeoutSeconds $TimeoutSeconds
    $parameters = [ordered]@{ Arch32 = 'x86'; Arch64 = 'x64'; AppObjectType = 31 }

    try {
        $dataSet = Invoke-DiscoverySqlQuery -ConnectionString $connectionString -CommandText $script:DiscoveryInventoryQuery -Parameters $parameters -TimeoutSeconds $TimeoutSeconds -OperationState $OperationState
    }
    catch {
        $message = ConvertTo-DiscoverySqlMessage -Exception $_.Exception -Server $Server -Database $Database -TimeoutSeconds $TimeoutSeconds
        throw (New-Object System.InvalidOperationException -ArgumentList $message, $_.Exception)
    }

    if (-not $dataSet -or $dataSet.Tables.Count -lt 2) {
        throw "The site database query on $Server returned an unexpected result."
    }

    $rows = @(foreach ($r in $dataSet.Tables[0].Rows) {
        [pscustomobject]@{
            DisplayName  = [string](Get-DiscoveryCellValue $r 'DisplayName')
            Version      = [string](Get-DiscoveryCellValue $r 'Version')
            Architecture = [string](Get-DiscoveryCellValue $r 'Architecture')
        }
    })

    $applications = @(foreach ($r in $dataSet.Tables[1].Rows) {
            $name = [string](Get-DiscoveryCellValue $r 'DisplayName')
            if ($name) { $name }
        })

    return [pscustomobject]@{
        Rows         = $rows
        Applications = $applications
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

$script:MecmQualifierPattern = '(?i)(?<![a-z0-9])(exe|msi|x64|x86|arm64|amd64|32-bit|64-bit)(?![a-z0-9])'
$script:MecmTitleTokenPattern = '\([^)]*\)|\S+'
# A "." suffix group must not start with a digit, or "2.3D" passes as "2" plus ".3D".
$script:MecmVersionPattern = 'v?\d+(\.\d+)*([-+_]\w+|\.[a-z_]\w*)*'
# PowerShell -match, -like and -replace fold case with the current culture;
# under tr-TR "MSI" does not match "msi".
$script:MecmRegexOptions = [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'

function Test-MecmMatch {
    param([string]$Text, [string]$Pattern)
    return [regex]::IsMatch($Text, $Pattern, $script:MecmRegexOptions)
}

function Get-MecmTitleQualifier {
    # Class name -> sorted distinct values. amd64 and 64-bit are x64;
    # 32-bit is x86.
    param([string]$Text)
    $classes = @{}
    foreach ($m in [regex]::Matches([string]$Text, $script:MecmQualifierPattern, $script:MecmRegexOptions)) {
        $value = $m.Value.ToLowerInvariant()
        $class = 'arch'
        switch ($value) {
            'exe' { $class = 'method' }
            'msi' { $class = 'method' }
            'amd64' { $value = 'x64' }
            '64-bit' { $value = 'x64' }
            '32-bit' { $value = 'x86' }
        }
        if (-not $classes.ContainsKey($class)) { $classes[$class] = New-Object System.Collections.Generic.SortedSet[string] }
        [void]$classes[$class].Add($value)
    }
    return $classes
}

function Test-MecmQualifierWord {
    # True when the word holds only installer-method and architecture
    # qualifiers: "(x64)", "(EXE)", "(x86+x64)", "ARM64".
    param([string]$Word)
    if (-not (Test-MecmMatch $Word $script:MecmQualifierPattern)) { return $false }
    return ([regex]::Replace([regex]::Replace($Word, $script:MecmQualifierPattern, '', $script:MecmRegexOptions), '[\s()+,/]', '')).Length -eq 0
}

function Test-MecmQualifierEqual {
    # True when the token carries every qualifier class of the word with the
    # same normalized values, so "(x64)" equals "(64-bit)".
    param([string]$Word, [string]$Token)
    $wordQualifier = Get-MecmTitleQualifier $Word
    $tokenQualifier = Get-MecmTitleQualifier $Token
    foreach ($class in $wordQualifier.Keys) {
        if (-not $tokenQualifier.ContainsKey($class) -or -not $tokenQualifier[$class].SetEquals($wordQualifier[$class])) { return $false }
    }
    return $true
}

function Get-MecmApplicationNamePattern {
    <#
    .SYNOPSIS
        Get-CMApplication -Name wildcard that returns every title
        Test-MecmApplicationTitle can accept for CMName.
    #>
    param([Parameter(Mandatory)][string]$CMName)
    $words = @([regex]::Matches($CMName.Trim(), $script:MecmTitleTokenPattern) | ForEach-Object { $_.Value })
    $kept = @($words | Where-Object { -not (Test-MecmQualifierWord $_) })
    if ($kept.Count -eq 0) { $kept = $words }
    return (($kept -join '*') + '*')
}

function Test-MecmApplicationTitle {
    <#
    .SYNOPSIS
        True when a ConfigMgr Application title is a packaged title of
        CMName. Check ConfigMgr and Discover share this rule.
    #>
    # A packaged title holds the words of CMName in order, with only release
    # details around them: a version, qualifiers in parentheses, an
    # architecture, and after the last word optionally " - <detail>", where
    # the first token of <detail> is a whole version ("26.03", "v2",
    # "1.2.3-beta", or "Latest"), a parenthetical or a qualifier word, and
    # any text may follow it. "2FA" or "3D" is a word, not a version. When
    # the last word continues as a dotted version right before the dash,
    # <detail> may also be "Windows Server Hosting" followed only by release
    # details: the vendor title of the ASP.NET hosting bundle. A
    # word of CMName that ends in a number may continue as a version
    # ("18" as "18.5.1.1") or carry an architecture ("7" as "7-x64"); vendor
    # product names put these in the middle ("Microsoft ODBC Driver
    # 18.5.1.1 for SQL Server") or glue them on ("PowerShell 7-x64"). Any
    # other word names a different product: "Git" must not find
    # "Git Extensions", "Git - Extensions" or "GitHub Desktop", and "Mozilla Firefox" must not
    # find "Mozilla Firefox ESR".
    # Installer method (EXE, MSI) and architecture are variant classes. When
    # CMName and the title both carry a class, the values must be equal, or
    # "OpenJDK 17 (EXE)" counts as managed by "OpenJDK 17 (MSI) (EXE)". A
    # CMName word that holds only qualifiers of classes the title does not
    # carry is optional ("Foo (x64)" accepts "Foo 1.2").
    param(
        [Parameter(Mandatory)][string]$CMName,
        [AllowEmptyString()][string]$Title
    )
    if ([string]::IsNullOrWhiteSpace($Title) -or [string]::IsNullOrWhiteSpace($CMName)) { return $false }
    $nameQualifier = Get-MecmTitleQualifier $CMName
    $titleQualifier = Get-MecmTitleQualifier $Title
    foreach ($class in $nameQualifier.Keys) {
        if ($titleQualifier.ContainsKey($class) -and -not $titleQualifier[$class].SetEquals($nameQualifier[$class])) { return $false }
    }
    $arch = '(-(x64|x86|arm64|amd64))?'
    $detail = '^(\([^)]*\)|' + $script:MecmVersionPattern + '|x64|x86|arm64|amd64)$'
    $dashDetail = '^(' + $script:MecmVersionPattern + '|\([^)]*\)|latest)$'
    $allWords = @([regex]::Matches($CMName.Trim(), $script:MecmTitleTokenPattern) | ForEach-Object { $_.Value })
    $words = @($allWords | Where-Object {
            if (-not (Test-MecmQualifierWord $_)) { return $true }
            foreach ($class in (Get-MecmTitleQualifier $_).Keys) { if ($titleQualifier.ContainsKey($class)) { return $true } }
            return $false
        })
    if ($words.Count -eq 0) { $words = $allWords }
    $next = 0
    $lastContinued = $false
    $tokens = @([regex]::Matches($Title, $script:MecmTitleTokenPattern) | ForEach-Object { $_.Value })
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $token = $tokens[$i]
        if ($next -eq $words.Count -and $token -eq '-') {
            $rest = @($tokens | Select-Object -Skip ($i + 1))
            if ($rest.Count -eq 0) { return $false }
            if ($lastContinued -and $rest.Count -ge 3 -and [string]::Equals(($rest[0..2] -join ' '), 'Windows Server Hosting', [System.StringComparison]::OrdinalIgnoreCase)) {
                foreach ($tail in @($rest | Select-Object -Skip 3)) { if (-not (Test-MecmMatch $tail $detail)) { return $false } }
                return $true
            }
            if (Test-MecmMatch $rest[0] $dashDetail) { return $true }
            # Test-MecmQualifierWord ignores parentheses, so an unclosed "(x64" passes it.
            return ((Test-MecmQualifierWord $rest[0]) -and $rest[0].IndexOfAny([char[]]'()') -lt 0)
        }
        $lastContinued = $false
        if ($next -lt $words.Count) {
            if ((Test-MecmQualifierWord $words[$next]) -and (Test-MecmQualifierWord $token)) {
                if (Test-MecmQualifierEqual -Word $words[$next] -Token $token) { $next++; continue }
            }
            $continuation = if ([regex]::IsMatch($words[$next], '\d$')) { '(?<version>(\.\d+)+([-+_]\w+|\.[a-z_]\w*)*)?' } else { '' }
            $wordMatch = [regex]::Match($token, ('^' + [regex]::Escape($words[$next]) + $continuation + $arch + '$'), $script:MecmRegexOptions)
            if ($wordMatch.Success) {
                $next++
                $lastContinued = ($next -eq $words.Count -and $wordMatch.Groups['version'].Success)
                continue
            }
        }
        if (Test-MecmMatch $token $detail) { continue }
        return $false
    }
    return ($next -eq $words.Count)
}

function Test-DiscoveryApplicationManaged {
    param([string]$CMName, [string[]]$Application)
    # Every title that the rule accepts contains the first word of CMName;
    # the substring test skips the per-title rule for all other titles.
    $first = @($CMName.Trim() -split '\s+')[0]
    foreach ($title in $Application) {
        if ([string]::Equals($CMName, $title, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
        if ($title.IndexOf($first, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if (Test-MecmApplicationTitle -CMName $CMName -Title $title) { return $true }
    }
    return $false
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
        and whose CMName has no current ConfigMgr Application (the Check
        ConfigMgr title rule of Test-MecmApplicationTitle). Throws when
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
    $applications = @(foreach ($a in @($Snapshot.Applications)) { if ($a) { ([string]$a).Trim() } })

    $candidates = foreach ($p in @($Packager)) {
        if (-not $p) { continue }
        $scriptName = ([string]$p.Script).Trim()
        if (-not $scriptName -or -not $hidden.Contains($scriptName) -or $dismissed.Contains($scriptName)) { continue }

        $cmName = ([string]$p.CMName).Trim()
        if ($cmName -and (Test-DiscoveryApplicationManaged -CMName $cmName -Application $applications)) { continue }

        $patterns = Get-DiscoveryArpPatterns -ArpName (Get-DiscoveryPackagerValue $p 'ArpName') -Script $scriptName
        if ($patterns.Count -eq 0) { continue }
        $versionPatterns = Get-DiscoveryArpPatterns -ArpName (Get-DiscoveryPackagerValue $p 'ArpVersion') -Script $scriptName -Tag 'ArpVersion'
        $arch = Get-DiscoveryPackagerValue $p 'ArpArch'

        $match = $null
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
            $match = $row
            break
        }
        if (-not $match) { continue }

        [pscustomobject]@{
            Script      = $scriptName
            CMName      = $cmName
            Vendor      = [string]$p.Vendor
            MatchedName = [string]$match.DisplayName
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
        One line for the log and the prompt: candidate count, or the
        no-inventory-data text for zero rows.
    #>
    param(
        [Parameter(Mandatory)]$Snapshot,
        [Parameter(Mandatory)][int]$CandidateCount
    )

    if (@($Snapshot.Rows).Count -eq 0) {
        return 'No inventory data: the site database has no Add or Remove Programs rows. Check that hardware inventory collects Add or Remove Programs.'
    }

    $appWord = if ($CandidateCount -eq 1) { 'app' } else { 'apps' }
    return "Discovered $CandidateCount $appWord."
}

Export-ModuleMember -Function @(
    'Resolve-CMDatabaseTarget', 'Get-CMInventorySnapshot', 'Find-DiscoveryCandidate', 'Get-DiscoveryStatusText',
    'Test-MecmApplicationTitle', 'Get-MecmApplicationNamePattern'
)
