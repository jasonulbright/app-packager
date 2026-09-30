<#
.SYNOPSIS
    Publishes staged AppPackager installers to WSUS as locally published
    updates.

.DESCRIPTION
    The WSUS destination beside the ConfigMgr and Intune adapters. A staged
    manifest becomes one Software Distribution Package: the vendor installer
    runs with the manifest's silent arguments, and the manifest's detection
    becomes the update's applicability rules. No packager script changes.

    Every publish is an update of an installed product. WSUS offers it only
    where the detection finds the product with an older version, and reports
    it installed at this version or newer. A computer without the product
    never gets it; first installations stay with ConfigMgr and Intune.

    Requirements on the publishing host: Windows PowerShell 5.1 and the WSUS
    administration API (the WSUS console, or RSAT WSUS tools on a client).
    The WSUS server needs a signing certificate; clients need that
    certificate in Trusted Publishers (and Trusted Root Certification
    Authorities when self-signed) plus the "Allow signed updates from an
    intranet Microsoft update service location" policy.

    Every call into Microsoft.UpdateServices.Administration goes through the
    adapter functions under "WSUS API adapter", so the rule and orchestration
    logic runs under test without the API installed.
#>

Set-StrictMode -Version 2.0

$script:WsusAdministrationAssemblyName = 'Microsoft.UpdateServices.Administration'


$script:WsusIdentityMarker = 'AppPackager identity:'

# Identity lines carry this type. Updates published with the retired type
# Application keep their own identity lines, so a new publish neither reuses
# nor supersedes them.
$script:WsusPackageType = 'Update'

# One vendor and one product category for every application: a server
# refuses local publishing past a small number of distinct categories, and
# vendor names such as Microsoft Corporation are reserved for Microsoft.
$script:WsusVendorName = 'AppPackager'
$script:WsusProductName = 'AppPackager Applications'

# A local-publishing cab cannot exceed this size, whatever the server's
# LocalPublishingMaxCabSize says.
$script:WsusCabHardLimitMegabytes = 2047

# AppPackager writes these beside every staged installer; none of them
# belongs in a WSUS payload. The scripts folder holds the generated
# detection and requirement scripts.
$script:WsusGeneratedContentNames = @(
    'install.bat', 'install.ps1', 'uninstall.bat', 'uninstall.ps1',
    'install-generated.ps1', 'uninstall-generated.ps1',
    'install-before.ps1', 'install-after.ps1', 'uninstall-before.ps1', 'uninstall-after.ps1',
    'stage-manifest.json', 'staged-version.txt', 'app-icon.ico', 'app-icon.png'
)
$script:WsusGeneratedContentFolder = 'scripts'


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

function Write-WsusLog {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log $Message -Level $Level
    }
}

function Get-WsusMemberValue {
    param($InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function ConvertTo-WsusXmlAttribute {
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-WsusVersionString {
    # WSUS version comparisons take four numeric parts; a vendor version with
    # fewer parts is padded, a non-numeric one is refused.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Version)
    $trimmed = ([string]$Version).Trim()
    if ($trimmed -notmatch '^\d+(\.\d+){0,3}$') { return $null }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in $trimmed.Split('.')) {
        $number = 0L
        if (-not [long]::TryParse($part, [ref]$number) -or $number -gt 65535) { return $null }
        $parts.Add([string]$number)
    }
    while ($parts.Count -lt 4) { $parts.Add('0') }
    return ($parts.ToArray() -join '.')
}


# ---------------------------------------------------------------------------
# Settings and identity
# ---------------------------------------------------------------------------

function Get-WsusClassificationNames {
    <#
    .SYNOPSIS
        The update classifications a published update can carry.
    #>
    return @('Updates', 'SecurityUpdates', 'CriticalUpdates', 'FeaturePacks', 'ServicePacks', 'Tools', 'UpdateRollups')
}

function ConvertTo-WsusPublishSettings {
    <#
    .SYNOPSIS
        Validates and normalizes WSUS publish settings from preferences or
        parameters.

    .DESCRIPTION
        Throws on a value that would publish somewhere other than intended (a
        malformed server name, an unknown classification, or a package type
        other than Update); fills defaults for absent values. The port
        defaults to 8531 with SSL and 8530 without.

    .OUTPUTS
        [hashtable] ServerName, PortNumber, UseSsl, Classification,
        ApprovalGroup, DeclineSuperseded.
    #>
    param([AllowNull()]$InputObject)

    $server = ([string](Get-WsusMemberValue -InputObject $InputObject -Name 'ServerName')).Trim()
    if ($server -and $server -notmatch '^[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)*$') {
        throw "WSUS server name '$server' is not a host name."
    }

    $useSsl = [bool](Get-WsusMemberValue -InputObject $InputObject -Name 'UseSsl')
    $port = 0
    [void][int]::TryParse([string](Get-WsusMemberValue -InputObject $InputObject -Name 'PortNumber'), [ref]$port)
    if ($port -lt 1 -or $port -gt 65535) { $port = $(if ($useSsl) { 8531 } else { 8530 }) }

    $packageType = ([string](Get-WsusMemberValue -InputObject $InputObject -Name 'PackageType')).Trim()
    if ($packageType -and $packageType -ne $script:WsusPackageType) {
        throw ("WSUS package type '{0}' is not available. AppPackager publishes updates for installed products only: an update never installs where the product is missing. Remove the PackageType setting." -f $packageType)
    }

    $classification = [string](Get-WsusMemberValue -InputObject $InputObject -Name 'Classification')
    if ([string]::IsNullOrWhiteSpace($classification)) { $classification = 'Updates' }
    if ($classification -notin (Get-WsusClassificationNames)) {
        throw ("WSUS classification '{0}' is not one of: {1}." -f $classification, ((Get-WsusClassificationNames) -join ', '))
    }

    return @{
        ServerName        = $server
        PortNumber        = $port
        UseSsl            = $useSsl
        Classification    = $classification
        ApprovalGroup     = ([string](Get-WsusMemberValue -InputObject $InputObject -Name 'ApprovalGroup')).Trim()
        DeclineSuperseded = [bool](Get-WsusMemberValue -InputObject $InputObject -Name 'DeclineSuperseded')
    }
}

function Get-WsusIdentityTag {
    <#
    .SYNOPSIS
        The identity an AppPackager update carries on the WSUS server.

    .DESCRIPTION
        'AppPackager:<ApplicationId>/<ProfileId>', the same form the Intune
        publisher writes, so one application maps to one identity on every
        destination. A manifest staged before the workbench falls back to a
        key derived from the publisher and the title.
    #>
    param([Parameter(Mandatory)]$Manifest)

    $applicationId = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'ApplicationId')
    if ([string]::IsNullOrWhiteSpace($applicationId)) {
        # Titles often carry the version; the key keeps only its major, so
        # every release of one major shares the identity that supersedence
        # follows. Side-by-side lines such as .NET 8 and .NET 10 differ only
        # in the major, and a shared key lets one line supersede the other.
        # The publisher keeps two vendors' products of one title apart.
        $title = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'AppName')).Trim()
        $version = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'SoftwareVersion')).Trim()
        if ($version) {
            $major = [regex]::Match($version, '^\d+').Value
            $title = [regex]::Replace($title, ('(?<![\w.])v?' + [regex]::Escape($version) + '(?![\w.])'), $major, 'IgnoreCase')
        }
        $key = ($title -replace '[^\w.]+', '-').Trim('-')
        $publisher = (([string](Get-WsusMemberValue -InputObject $Manifest -Name 'Publisher')) -replace '[^\w.]+', '-').Trim('-')
        if ($publisher) { $key = $publisher + ':' + $key }
        $applicationId = 'legacy:' + $key
    }
    $profileId = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'ProfileId')
    if ([string]::IsNullOrWhiteSpace($profileId)) { $profileId = 'default' }
    return ('AppPackager:{0}/{1}' -f $applicationId, $profileId)
}

function Get-WsusIdentityLine {
    <#
    .SYNOPSIS
        The description line that marks an update as this identity's, with
        the version that supersedence compares.
    #>
    param([Parameter(Mandatory)][string]$IdentityTag, [Parameter(Mandatory)][string]$PackageType, [string]$Version = '')
    $line = '{0} {1}; type {2}' -f $script:WsusIdentityMarker, $IdentityTag, $PackageType
    if (-not [string]::IsNullOrWhiteSpace($Version)) { $line += ('; version {0}' -f $Version.Trim()) }
    return $line
}

function Get-WsusUpdateIdentity {
    <#
    .SYNOPSIS
        Reads the identity line back out of an update description.

    .OUTPUTS
        [pscustomobject] IdentityTag, PackageType, Version ('' when the line
        carries none); $null when the update does not carry an AppPackager
        identity line.
    #>
    param([AllowEmptyString()][AllowNull()][string]$Description)
    if ([string]::IsNullOrWhiteSpace($Description)) { return $null }
    $match = [regex]::Match($Description, ('(?m)^' + [regex]::Escape($script:WsusIdentityMarker) + ' (?<tag>AppPackager:\S+); type (?<type>Update|Application)(; version (?<version>[^\r\n]+?))?[ \t]*\r?$'))
    if (-not $match.Success) { return $null }
    return [pscustomobject]@{ IdentityTag = $match.Groups['tag'].Value; PackageType = $match.Groups['type'].Value; Version = $match.Groups['version'].Value }
}

function Compare-WsusVersion {
    # Orders two vendor versions by their numeric runs, so 5.10 follows 5.9
    # and 2026.08.2-200 follows 2026.08.2-199; a missing run counts as zero.
    # $null when either version carries no digits.
    param([AllowEmptyString()][string]$Left, [AllowEmptyString()][string]$Right)
    $a = @([regex]::Matches([string]$Left, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })
    $b = @([regex]::Matches([string]$Right, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })
    if ($a.Count -eq 0 -or $b.Count -eq 0) { return $null }
    for ($i = 0; $i -lt [Math]::Max($a.Count, $b.Count); $i++) {
        $x = if ($i -lt $a.Count) { [string]$a[$i] } else { '' }
        $y = if ($i -lt $b.Count) { [string]$b[$i] } else { '' }
        if ($x.Length -ne $y.Length) { return $(if ($x.Length -lt $y.Length) { -1 } else { 1 }) }
        $order = [string]::CompareOrdinal($x, $y)
        if ($order -ne 0) { return [Math]::Sign($order) }
    }
    return 0
}

function ConvertFrom-WsusCatalogInput {
    <#
    .SYNOPSIS
        Extracts Microsoft Update Catalog update IDs from pasted text.

    .DESCRIPTION
        Accepts bare update IDs and catalog links that carry one
        (ScopedViewInline.aspx?updateid=...), any number per line, and
        returns each distinct ID once, in input order.

    .OUTPUTS
        [guid[]]
    #>
    param([AllowEmptyString()][AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $seen = New-Object 'System.Collections.Generic.HashSet[guid]'
    $ids = New-Object 'System.Collections.Generic.List[guid]'
    foreach ($match in [regex]::Matches($Text, '(?<![0-9A-Fa-f])[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}(?![0-9A-Fa-f])')) {
        $id = [guid]$match.Value
        if ($seen.Add($id)) { $ids.Add($id) }
    }
    return @($ids.ToArray())
}

# ---------------------------------------------------------------------------
# Installer payload
# ---------------------------------------------------------------------------

function Split-WsusCommandLine {
    # Splits an argument string the way the Windows command-line parser
    # groups it: whitespace separates, double quotes group, and a quoted
    # segment stays attached to the text before it (PROPERTY="a b").
    param([AllowEmptyString()][string]$CommandLine)

    $tokens = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($CommandLine)) { return @() }
    $current = New-Object System.Text.StringBuilder
    $quoted = $false
    foreach ($character in $CommandLine.ToCharArray()) {
        if ($character -eq '"') { $quoted = -not $quoted; [void]$current.Append($character); continue }
        if ([char]::IsWhiteSpace($character) -and -not $quoted) {
            if ($current.Length -gt 0) { $tokens.Add($current.ToString()); [void]$current.Clear() }
            continue
        }
        [void]$current.Append($character)
    }
    if ($current.Length -gt 0) { $tokens.Add($current.ToString()) }
    return @($tokens.ToArray())
}

function ConvertTo-WsusMsiCommandLine {
    <#
    .SYNOPSIS
        Reduces an msiexec argument string to the property assignments WSUS
        passes to Windows Installer.

    .DESCRIPTION
        WSUS runs the package itself, quietly; msiexec switches (/qn,
        /norestart, /l*v) have no meaning there and are dropped. /norestart
        becomes REBOOT=ReallySuppress so the package still defers its restart
        to the client's restart policy.
    #>
    param([AllowEmptyString()][string]$InstallArgs)

    $properties = New-Object System.Collections.Generic.List[string]
    $suppressReboot = $false
    $skipNext = $false
    foreach ($token in (Split-WsusCommandLine -CommandLine $InstallArgs)) {
        if ($skipNext) { $skipNext = $false; continue }
        if ($token -match '^[/-]') {
            if ($token -match '^[/-](norestart|promptrestart)$') { $suppressReboot = $true }
            # /l*v <file> and /log <file> carry their file as the next token.
            if ($token -match '^[/-](l[a-z\*\+!]*|log)$') { $skipNext = $true }
            continue
        }
        if ($token -match '^[A-Za-z_][A-Za-z0-9_.]*=') { $properties.Add($token) }
    }
    if ($suppressReboot -and -not ($properties | Where-Object { $_ -match '^REBOOT=' })) {
        $properties.Add('REBOOT=ReallySuppress')
    }
    return ($properties.ToArray() -join ' ')
}

function Get-WsusPayloadFiles {
    <#
    .SYNOPSIS
        The staged files a WSUS package carries: every file the stage
        recorded except the files AppPackager generates for its other
        destinations, with the installer first.

    .DESCRIPTION
        An installer reads companion files that its arguments do not name
        (setup.ini, an external MSI cab, a patch), so the recorded stage
        content is the payload, not a list derived from the arguments.

    .OUTPUTS
        [string[]] Paths relative to the content folder.
    #>
    param([Parameter(Mandatory)]$Manifest)

    $installer = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerFile')
    $icon = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'Icon')
    $files = New-Object System.Collections.Generic.List[string]
    $files.Add($installer)
    foreach ($entry in @(Get-WsusMemberValue -InputObject $Manifest -Name 'FileHashes')) {
        if ($null -eq $entry) { continue }
        $relative = ([string](Get-WsusMemberValue -InputObject $entry -Name 'RelativePath')).Replace('/', '\').TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($relative)) { continue }
        if ($relative -in $script:WsusGeneratedContentNames -or ($icon -and $relative -eq $icon)) { continue }
        if ($relative -like ($script:WsusGeneratedContentFolder + '\*')) { continue }
        if (-not ($files -contains $relative)) { $files.Add($relative) }
    }
    return @($files.ToArray())
}

function Resolve-WsusPayloadPath {
    # A recorded payload path resolves inside the content folder or not at
    # all: a rooted, drive-relative, stream, or parent-relative path would
    # copy a file from elsewhere on this computer into a signed update.
    param([Parameter(Mandatory)][string]$ContentFolder, [Parameter(Mandatory)][AllowEmptyString()][string]$RelativePath)

    if ([string]::IsNullOrWhiteSpace($RelativePath) -or $RelativePath.Contains(':') -or [System.IO.Path]::IsPathRooted($RelativePath)) {
        throw ("The staged file path '{0}' is not a path inside the content folder." -f $RelativePath)
    }
    $root = [System.IO.Path]::GetFullPath($ContentFolder).TrimEnd('\')
    $full = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($root, $RelativePath))
    if (-not $full.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw ("The staged file path '{0}' points outside the content folder." -f $RelativePath)
    }
    return $full
}

function Get-WsusInstallScriptExtras {
    # The steps an install.ps1 takes besides reads, messages, closing a
    # running copy, one installer process, and its exit code. WSUS runs the
    # installer itself, so each step listed here never reaches the client.
    param([Parameter(Mandatory)][string]$Path)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) { return @('install.ps1 does not parse') }

    $neutral = @('Join-Path', 'Split-Path', 'Test-Path', 'Resolve-Path', 'Get-ChildItem', 'Get-Item', 'Get-ItemProperty', 'Get-Date',
        'Where-Object', 'Select-Object', 'Sort-Object', 'ForEach-Object', 'Set-Location', 'Push-Location', 'Pop-Location',
        'Out-Null', 'Write-Output', 'Write-Host', 'Write-Verbose', 'Write-Warning', 'Write-Error', 'Write-Debug', 'Write-Information',
        'Get-Process', 'Stop-Process', 'Wait-Process', 'Start-Sleep')
    $extras = New-Object System.Collections.Generic.List[string]
    $launches = 0
    foreach ($command in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true))) {
        $name = [string]$command.GetCommandName()
        $alias = if ($name) { Get-Alias -Name ([System.Management.Automation.WildcardPattern]::Escape($name)) -ErrorAction SilentlyContinue | Select-Object -First 1 } else { $null }
        if ($alias) { $name = [string]$alias.ResolvedCommandName }
        if ($name -in $neutral) { continue }
        if ($name -eq 'Out-File' -and $command.Extent.Text -match 'exitcode\.txt') { continue }
        $native = ($name -match '\.(exe|com|cmd|bat)$' -or ($name -and (Get-Command -Name $name -CommandType Application -ErrorAction SilentlyContinue)))
        if ($name -eq 'Start-Process' -or $command.InvocationOperator -eq 'Ampersand' -or $native) {
            $launches++
            continue
        }
        $extras.Add($(if ($name) { $name } else { $command.Extent.Text }))
    }
    if ($launches -gt 1) { $extras.Add(('{0} process launches' -f $launches)) }

    $pathTypes = @('IO.Path', 'System.IO.Path', 'string', 'System.String')
    foreach ($member in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.MemberExpressionAst] -and $node.Static }, $true))) {
        $typeName = [string]$member.Expression.TypeName.FullName
        if ($typeName -in $pathTypes) { continue }
        $extras.Add(('[{0}]::{1}' -f $typeName, $member.Member.Extent.Text))
    }

    # An object from an allowed read, such as a file from Get-Item, can still
    # change the computer through its own methods and properties.
    $harmlessMethods = @('WaitForExit', 'Refresh', 'Kill', 'CloseMainWindow', 'ToString', 'Trim', 'TrimStart', 'TrimEnd', 'Replace', 'Split',
        'Substring', 'ToLower', 'ToUpper', 'ToLowerInvariant', 'ToUpperInvariant', 'StartsWith', 'EndsWith', 'Contains', 'IndexOf', 'Equals',
        'GetType', 'Where', 'ForEach')
    foreach ($call in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and -not $node.Static }, $true))) {
        $method = if ($call.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) { [string]$call.Member.Value } else { $call.Member.Extent.Text }
        if ($method -in $harmlessMethods) { continue }
        $extras.Add(('.{0}()' -f $method))
    }
    foreach ($assignment in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))) {
        $target = $assignment.Left
        while ($target -is [System.Management.Automation.Language.ConvertExpressionAst]) { $target = $target.Child }
        $isMember = $target -is [System.Management.Automation.Language.MemberExpressionAst] -and -not $target.Static
        $isEnvironment = $target -is [System.Management.Automation.Language.VariableExpressionAst] -and $target.VariablePath.DriveName -eq 'env'
        if ($isMember -or $isEnvironment) { $extras.Add(('{0} assignment' -f $target.Extent.Text)) }
    }
    return @($extras.ToArray() | Select-Object -Unique)
}

# ---------------------------------------------------------------------------
# Applicability rules
# ---------------------------------------------------------------------------

# An equality detection becomes "this version or newer" so a client that
# already runs a newer copy reports the update installed instead of being
# offered a downgrade.
$script:WsusVersionComparison = @{
    IsEquals      = 'GreaterThanOrEqualTo'
    GreaterEquals = 'GreaterThanOrEqualTo'
    GreaterThan   = 'GreaterThan'
}
# The complement of each installed comparison: together they cover every
# version a present copy can carry, and neither holds for a missing copy.
$script:WsusOlderComparison = @{
    GreaterThanOrEqualTo = 'LessThan'
    GreaterThan          = 'LessThanOrEqualTo'
}
$script:WsusStringComparison = @{
    IsEquals   = 'EqualTo'
    BeginsWith = 'BeginsWith'
    EndsWith   = 'EndsWith'
    Contains   = 'Contains'
}

function Get-WsusRegistryLocation {
    param([Parameter(Mandatory)]$Clause)

    $hive = [string](Get-WsusMemberValue -InputObject $Clause -Name 'Hive')
    if ($hive -match '^(CurrentUser|HKCU)$') {
        throw 'The detection reads HKEY_CURRENT_USER. WSUS evaluates applicability as SYSTEM, which has no signed-in user hive.'
    }
    $subkey = ([string](Get-WsusMemberValue -InputObject $Clause -Name 'RegistryKeyRelative')).Trim().Trim('\')
    if ([string]::IsNullOrWhiteSpace($subkey)) { throw 'The detection names no registry key.' }
    # Same default as the ConfigMgr clause builder: without Is64Bit the key
    # is read from the 32-bit view on a 64-bit client.
    $attributes = 'Key="HKEY_LOCAL_MACHINE" Subkey="{0}"' -f (ConvertTo-WsusXmlAttribute $subkey)
    if (-not [bool](Get-WsusMemberValue -InputObject $Clause -Name 'Is64Bit')) { $attributes += ' RegType32="true"' }
    return [pscustomobject]@{
        Attributes      = $attributes
        Subkey          = $subkey
        # A Windows Installer product key changes with every version, so its
        # presence proves this exact version, never an older one.
        VersionSpecific = ($subkey -match '\\\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$')
    }
}

function Get-WsusFileLocation {
    param([Parameter(Mandatory)]$Clause)

    $folder = [string](Get-WsusMemberValue -InputObject $Clause -Name 'FilePath')
    $name = [string](Get-WsusMemberValue -InputObject $Clause -Name 'FileName')
    if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($name)) { throw 'The file detection names no folder or file.' }
    if ($folder -match '%(LOCALAPPDATA|APPDATA|USERPROFILE|HOMEPATH|HOMEDRIVE|USERNAME|TEMP|TMP|ONEDRIVE)%') {
        throw ("The file detection reads the per-user path '{0}'. WSUS evaluates applicability as SYSTEM, which resolves no signed-in user profile." -f $folder)
    }
    # Same expansion as the ConfigMgr and Intune clauses: a clause not
    # marked 64-bit expands %ProgramFiles% and %CommonProgramFiles% the way
    # a 32-bit process does, whatever the bitness of this process.
    $is64Bit = [bool](Get-WsusMemberValue -InputObject $Clause -Name 'Is64Bit')
    $expandProgramFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $expandCommonFiles = if ($env:CommonProgramW6432) { $env:CommonProgramW6432 } else { $env:CommonProgramFiles }
    if (-not $is64Bit -and ${env:ProgramFiles(x86)}) { $expandProgramFiles = ${env:ProgramFiles(x86)} }
    if (-not $is64Bit -and ${env:CommonProgramFiles(x86)}) { $expandCommonFiles = ${env:CommonProgramFiles(x86)} }
    $folder = ($folder -ireplace '%ProgramFiles%', $expandProgramFiles.Replace('$', '$$')) -ireplace '%CommonProgramFiles%', $expandCommonFiles.Replace('$', '$$')
    $expanded = [Environment]::ExpandEnvironmentVariables($folder).TrimEnd('\')
    if ($expanded -match '%[^%]+%') { throw ("The file detection path '{0}' carries an unresolved variable." -f $folder) }
    if ($expanded -notmatch '^[A-Za-z]:\\') { throw ("The file detection path '{0}' is not an absolute local path." -f $folder) }

    # Known folders resolve on the client through their CSIDL, so a client
    # whose Windows lives on another drive still matches.
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $knownFolders = @(
        @{ Path = ${env:ProgramFiles(x86)}; Csidl = 42 }
        @{ Path = $programFiles; Csidl = 38 }
        @{ Path = $env:ProgramData; Csidl = 35 }
        @{ Path = ([System.IO.Path]::Combine([string]$env:SystemRoot, 'System32')); Csidl = 37 }
        @{ Path = $env:SystemRoot; Csidl = 36 }
    )
    foreach ($known in $knownFolders) {
        if ([string]::IsNullOrWhiteSpace([string]$known.Path)) { continue }
        $root = ([string]$known.Path).TrimEnd('\')
        if ($expanded -ieq $root -or $expanded.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            $relative = $expanded.Substring($root.Length).TrimStart('\')
            $path = if ($relative) { [System.IO.Path]::Combine($relative, $name) } else { $name }
            return [pscustomobject]@{ Attributes = ('Path="{0}" Csidl="{1}"' -f (ConvertTo-WsusXmlAttribute $path), $known.Csidl) }
        }
    }
    return [pscustomobject]@{ Attributes = ('Path="{0}"' -f (ConvertTo-WsusXmlAttribute ([System.IO.Path]::Combine($expanded, $name)))) }
}

function Assert-WsusStableRegistryLocation {
    param([Parameter(Mandatory)]$Location)
    if ($Location.VersionSpecific) {
        throw ("The detection reads the Windows Installer product key '{0}'. An MSI registers a new product key for each version, so WSUS cannot find an older installed version under that key. Change the detection to a file version, or to a version value under a registry key that every version uses, or add a WsusDetection block that names the product's Add/Remove Programs entry." -f $Location.Subkey)
    }
}

function ConvertTo-WsusClauseRules {
    <#
        One detection clause as WSUS rules. Kind Version: Installed is "this
        version or newer" and Older is "the file or value exists and carries
        an older version"; a missing file or value fails both. Kind Presence:
        Installed is the existence check and Older is empty, because
        existence says nothing about the version.
    #>
    param([Parameter(Mandatory)]$Clause)

    $type = [string](Get-WsusMemberValue -InputObject $Clause -Name 'Type')
    if ([string]::IsNullOrWhiteSpace($type)) { $type = 'RegistryKeyValue' }
    switch ($type) {
        'RegistryKeyValue' {
            $location = Get-WsusRegistryLocation -Clause $Clause
            Assert-WsusStableRegistryLocation -Location $location
            $valueName = [string](Get-WsusMemberValue -InputObject $Clause -Name 'ValueName')
            if ([string]::IsNullOrWhiteSpace($valueName)) { $valueName = 'DisplayVersion' }
            $expected = [string](Get-WsusMemberValue -InputObject $Clause -Name 'ExpectedValue')
            if ([string]::IsNullOrWhiteSpace($expected)) { $expected = [string](Get-WsusMemberValue -InputObject $Clause -Name 'DisplayVersion') }
            $operator = [string](Get-WsusMemberValue -InputObject $Clause -Name 'Operator')
            if ([string]::IsNullOrWhiteSpace($operator)) { $operator = 'IsEquals' }
            $propertyType = [string](Get-WsusMemberValue -InputObject $Clause -Name 'PropertyType')
            if ($propertyType -and $propertyType -notin @('String', 'Version')) {
                throw ("Registry value '{0}' is compared as {1}; the WSUS publisher maps version comparisons only." -f $valueName, $propertyType)
            }
            if (-not $script:WsusVersionComparison.ContainsKey($operator)) {
                if ($script:WsusStringComparison.ContainsKey($operator)) {
                    throw ("The detection compares registry value '{0}' as text ({1} '{2}'). WSUS can only tell an older installed version from a newer one with a version comparison. Change the detection to compare the value as a version." -f $valueName, $operator, $expected)
                }
                throw ("Detection operator '{0}' on registry value '{1}' has no WSUS applicability mapping." -f $operator, $valueName)
            }
            $version = ConvertTo-WsusVersionString -Version $expected
            if (-not $version) {
                throw ("The detection compares registry value '{0}' with '{1}', which is not a version of up to four numbers. WSUS cannot compare it with the installed version. Change the detection to a numeric version." -f $valueName, $expected)
            }
            $comparison = $script:WsusVersionComparison[$operator]
            $format = '<bar:RegSzToVersion {0} Value="{1}" Comparison="{2}" Data="{3}" />'
            # RegSzToVersion compares a missing value as version 0, so an older
            # check without the existence guard holds on a computer that has
            # no copy of the product.
            $exists = '<bar:RegValueExists {0} Value="{1}" Type="REG_SZ" />' -f $location.Attributes, (ConvertTo-WsusXmlAttribute $valueName)
            return [pscustomobject]@{
                Kind      = 'Version'
                Installed = ($format -f $location.Attributes, (ConvertTo-WsusXmlAttribute $valueName), $comparison, $version)
                Older     = (Join-WsusRules -Connector And -Rules @($exists, ($format -f $location.Attributes, (ConvertTo-WsusXmlAttribute $valueName), $script:WsusOlderComparison[$comparison], $version)))
            }
        }
        'RegistryKey' {
            $location = Get-WsusRegistryLocation -Clause $Clause
            Assert-WsusStableRegistryLocation -Location $location
            return [pscustomobject]@{ Kind = 'Presence'; Installed = ('<bar:RegKeyExists {0} />' -f $location.Attributes); Older = $null }
        }
        'File' {
            $location = Get-WsusFileLocation -Clause $Clause
            $propertyType = [string](Get-WsusMemberValue -InputObject $Clause -Name 'PropertyType')
            if ($propertyType -eq 'Version') {
                $operator = [string](Get-WsusMemberValue -InputObject $Clause -Name 'Operator')
                if ([string]::IsNullOrWhiteSpace($operator)) { $operator = 'GreaterEquals' }
                if (-not $script:WsusVersionComparison.ContainsKey($operator)) {
                    throw ("Detection operator '{0}' on a file version has no WSUS applicability mapping." -f $operator)
                }
                $version = ConvertTo-WsusVersionString -Version ([string](Get-WsusMemberValue -InputObject $Clause -Name 'ExpectedValue'))
                if (-not $version) { throw 'The file detection compares a version that is not numeric.' }
                $comparison = $script:WsusVersionComparison[$operator]
                $format = '<bar:FileVersion {0} Comparison="{1}" Version="{2}" />'
                $exists = '<bar:FileExists {0} />' -f $location.Attributes
                return [pscustomobject]@{
                    Kind      = 'Version'
                    Installed = ($format -f $location.Attributes, $comparison, $version)
                    Older     = (Join-WsusRules -Connector And -Rules @($exists, ($format -f $location.Attributes, $script:WsusOlderComparison[$comparison], $version)))
                }
            }
            if ($propertyType -in @('', 'Existence')) {
                return [pscustomobject]@{ Kind = 'Presence'; Installed = ('<bar:FileExists {0} />' -f $location.Attributes); Older = $null }
            }
            throw ("The file detection compares {0}; the WSUS publisher maps file versions and existence only." -f $propertyType)
        }
        'Script' { throw 'Script detection has no WSUS equivalent: applicability rules cannot run a script.' }
        default { throw ("Detection type '{0}' has no WSUS applicability mapping." -f $type) }
    }
}

function Join-WsusRules {
    param([Parameter(Mandatory)][ValidateSet('And', 'Or')][string]$Connector, [Parameter(Mandatory)][string[]]$Rules)
    if ($Rules.Count -eq 1) { return $Rules[0] }
    return ('<lar:{0}>{1}</lar:{0}>' -f $Connector, ($Rules -join ''))
}

function Get-WsusDetectionClauses {
    # Flat clause list plus connector; two-group detections have no single
    # connector and are refused.
    param([Parameter(Mandatory)]$Detection)

    $type = [string](Get-WsusMemberValue -InputObject $Detection -Name 'Type')
    if ($type -ne 'Compound') { return [pscustomobject]@{ Connector = 'And'; Clauses = @($Detection) } }
    $groups = Get-WsusMemberValue -InputObject $Detection -Name 'GroupSizes'
    if ($null -ne $groups -and @($groups).Count -gt 0) {
        throw 'Grouped compound detections have no single AND or OR connector the WSUS publisher can map.'
    }
    $connector = [string](Get-WsusMemberValue -InputObject $Detection -Name 'Connector')
    if ($connector -notin @('And', 'Or')) { $connector = 'And' }
    $clauses = @(Get-WsusMemberValue -InputObject $Detection -Name 'Clauses' | Where-Object { $null -ne $_ })
    if ($clauses.Count -eq 0) { throw 'The compound detection carries no clauses.' }
    return [pscustomobject]@{ Connector = $connector; Clauses = $clauses }
}

function ConvertTo-WsusArpEntryRules {
    <#
        WSUS-only detection from the product's own Add/Remove Programs entry:
        a registry loop over the Uninstall subkeys finds the entry by its
        DisplayName (exact, or prefix with an optional suffix), an optional
        Publisher and, with WindowsInstaller, the DWORD that Windows
        Installer writes on its own entries, then compares DisplayVersion. The Windows Update Agent reads
        the looped subkey through Key="HKEY_LOOP_TARGET" Subkey="\" ("." never
        matches), and each child rule must repeat RegType32 or it reads the
        64-bit view.
    #>
    param([Parameter(Mandatory)]$Detection)

    $type = [string](Get-WsusMemberValue -InputObject $Detection -Name 'Type')
    if ($type -ne 'ArpEntry') { throw ("WsusDetection type '{0}' is not supported; the WSUS publisher reads type ArpEntry." -f $type) }
    $exactName = [string](Get-WsusMemberValue -InputObject $Detection -Name 'DisplayName')
    $prefix = [string](Get-WsusMemberValue -InputObject $Detection -Name 'DisplayNamePrefix')
    if ([string]::IsNullOrEmpty($prefix) -and [string]::IsNullOrEmpty($exactName)) { throw 'WsusDetection names no DisplayName or DisplayNamePrefix, so it cannot find the Add/Remove Programs entry.' }
    $suffix = [string](Get-WsusMemberValue -InputObject $Detection -Name 'DisplayNameSuffix')
    $publisher = [string](Get-WsusMemberValue -InputObject $Detection -Name 'Publisher')
    $windowsInstaller = [bool](Get-WsusMemberValue -InputObject $Detection -Name 'WindowsInstaller')
    $expected = [string](Get-WsusMemberValue -InputObject $Detection -Name 'Version')
    $version = ConvertTo-WsusVersionString -Version $expected
    if (-not $version) { throw ("WsusDetection compares DisplayVersion with '{0}', which is not a version of up to four numbers." -f $expected) }
    $viewText = [string](Get-WsusMemberValue -InputObject $Detection -Name 'View')
    $views = switch -Regex ($viewText) {
        '^32$' { @($true) }
        '^64$' { @($false) }
        '^Both$' { @($true, $false) }
        default { throw ("WsusDetection view '{0}' is not 32, 64 or Both." -f $viewText) }
    }

    $installed = New-Object System.Collections.Generic.List[string]
    $older = New-Object System.Collections.Generic.List[string]
    foreach ($is32 in $views) {
        $flag = if ($is32) { ' RegType32="true"' } else { '' }
        $target = 'Key="HKEY_LOOP_TARGET" Subkey="\"' + $flag
        $match = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrEmpty($exactName)) {
            $match.Add(('<bar:RegSz {0} Value="DisplayName" Comparison="EqualTo" Data="{1}" />' -f $target, (ConvertTo-WsusXmlAttribute $exactName)))
        }
        else {
            $match.Add(('<bar:RegSz {0} Value="DisplayName" Comparison="BeginsWith" Data="{1}" />' -f $target, (ConvertTo-WsusXmlAttribute $prefix)))
            if (-not [string]::IsNullOrEmpty($suffix)) { $match.Add(('<bar:RegSz {0} Value="DisplayName" Comparison="EndsWith" Data="{1}" />' -f $target, (ConvertTo-WsusXmlAttribute $suffix))) }
        }
        if (-not [string]::IsNullOrEmpty($publisher)) { $match.Add(('<bar:RegSz {0} Value="Publisher" Comparison="EqualTo" Data="{1}" />' -f $target, (ConvertTo-WsusXmlAttribute $publisher))) }
        if ($windowsInstaller) { $match.Add(('<bar:RegDword {0} Value="WindowsInstaller" Comparison="EqualTo" Data="1" />' -f $target)) }
        $loop = '<bar:RegKeyLoop Key="HKEY_LOCAL_MACHINE" Subkey="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"' + $flag + ' TrueIf="Any">{0}</bar:RegKeyLoop>'
        $installed.Add(($loop -f (Join-WsusRules -Connector And -Rules (@($match) + ('<bar:RegSzToVersion {0} Value="DisplayVersion" Comparison="GreaterThanOrEqualTo" Data="{1}" />' -f $target, $version)))))
        # A missing DisplayVersion compares as version 0, so the older check
        # needs the existence guard.
        $older.Add(($loop -f (Join-WsusRules -Connector And -Rules (@($match) +
                        ('<bar:RegValueExists {0} Value="DisplayVersion" Type="REG_SZ" />' -f $target) +
                        ('<bar:RegSzToVersion {0} Value="DisplayVersion" Comparison="LessThan" Data="{1}" />' -f $target, $version)))))
    }
    return [pscustomobject]@{
        IsInstalled   = (Join-WsusRules -Connector Or -Rules $installed.ToArray())
        IsInstallable = (Join-WsusRules -Connector Or -Rules $older.ToArray())
    }
}

function Get-WsusMsiInfo {
    # ProductName, ProductVersion, Manufacturer and the Template platform of
    # an MSI, read-only.
    param([Parameter(Mandatory)][string]$Path)

    $installer = $null; $db = $null; $summary = $null
    $views = New-Object System.Collections.Generic.List[object]
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $values = @{}
        foreach ($name in 'ProductName', 'ProductVersion', 'Manufacturer') {
            $view = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db, @("SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$name'"))
            $views.Add($view)
            [void]$view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
            $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
            $values[$name] = if ($null -ne $record) { [string]$record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1) } else { '' }
        }
        $summary = $db.GetType().InvokeMember('SummaryInformation', 'GetProperty', $null, $db, 0)
        $template = [string]$summary.GetType().InvokeMember('Property', 'GetProperty', $null, $summary, 7)
        return [pscustomobject]@{
            ProductName    = $values.ProductName
            ProductVersion = $values.ProductVersion
            Manufacturer   = $values.Manufacturer
            Platform       = ($template -split ';')[0]
        }
    }
    finally {
        # An unreleased database handle keeps the MSI locked for later copies.
        foreach ($o in @($views.ToArray()) + @($summary, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) }
        }
    }
}

function Get-WsusMsiArpDetection {
    <#
        The Add/Remove Programs entry that the staged MSI registers, as an
        ArpEntry: Windows Installer writes DisplayName = ProductName,
        Publisher = Manufacturer, DisplayVersion = ProductVersion and
        WindowsInstaller = 1, in the registry view of the package platform. A
        version inside ProductName becomes the gap between DisplayNamePrefix
        and DisplayNameSuffix, so every release of the product matches; the
        WindowsInstaller flag keeps an EXE-installed copy with a similar name
        from matching. $null when the manifest does
        not install an MSI or the MSI lacks one of the three properties.
    #>
    param([Parameter(Mandatory)]$Manifest, [Parameter(Mandatory)][string]$ContentFolder)

    if ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerType') -ne 'MSI') { return $null }
    $file = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerFile')
    if ([string]::IsNullOrWhiteSpace($file)) { return $null }
    try { $path = Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $file } catch { return $null }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try { $msi = Get-WsusMsiInfo -Path $path } catch { return $null }

    $name = ([string]$msi.ProductName).Trim()
    $publisher = ([string]$msi.Manufacturer).Trim()
    $productVersion = ([string]$msi.ProductVersion).Trim()
    if (-not $name -or -not $publisher -or -not (ConvertTo-WsusVersionString -Version $productVersion)) { return $null }

    # The name may carry the version in a shorter form ("24.08" for
    # ProductVersion 24.08.00.0); the longest form found wins.
    $parts = $productVersion.Split('.')
    $forms = @(for ($n = $parts.Count; $n -ge 2; $n--) { $parts[0..($n - 1)] -join '.' })
    $softwareVersion = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'SoftwareVersion')).Trim()
    if ($softwareVersion) { $forms += $softwareVersion }
    $token = $null
    foreach ($form in @($forms | Select-Object -Unique | Sort-Object Length -Descending)) {
        $m = [regex]::Match($name, ('(?<![\w.])v?' + [regex]::Escape($form) + '(?![\w.])'), 'IgnoreCase')
        if ($m.Success) { $token = $m; break }
    }

    $entry = [ordered]@{
        Type                  = 'ArpEntry'
        View                  = $(if ($msi.Platform -match '^(x64|Arm64|Intel64|AMD64)$') { '64' } else { '32' })
        Publisher             = $publisher
        WindowsInstaller      = $true
        Version               = $productVersion
    }
    if ($null -eq $token) { $entry.DisplayName = $name }
    else {
        $prefix = $name.Substring(0, $token.Index)
        if (-not $prefix.Trim()) { return $null }
        $entry.DisplayNamePrefix = $prefix
        $suffix = $name.Substring($token.Index + $token.Length)
        if ($suffix) { $entry.DisplayNameSuffix = $suffix }
    }
    return [pscustomobject]$entry
}

function Format-WsusArpEntry {
    param([Parameter(Mandatory)]$Entry)
    $exact = [string](Get-WsusMemberValue -InputObject $Entry -Name 'DisplayName')
    $name = if ($exact) { $exact } else { [string](Get-WsusMemberValue -InputObject $Entry -Name 'DisplayNamePrefix') + '*' + [string](Get-WsusMemberValue -InputObject $Entry -Name 'DisplayNameSuffix') }
    $publisher = [string](Get-WsusMemberValue -InputObject $Entry -Name 'Publisher')
    $tail = $(if ([bool](Get-WsusMemberValue -InputObject $Entry -Name 'WindowsInstaller')) { ', Windows Installer entry' } else { '' })
    if ($publisher) { return ("'{0}' by {1}, {2}-bit view{3}" -f $name, $publisher, [string](Get-WsusMemberValue -InputObject $Entry -Name 'View'), $tail) }
    return ("'{0}', {1}-bit view{2}" -f $name, [string](Get-WsusMemberValue -InputObject $Entry -Name 'View'), $tail)
}

function ConvertTo-WsusApplicabilityRules {
    <#
    .SYNOPSIS
        Maps a stage manifest onto WSUS applicability rules for an update of
        an installed product.

    .DESCRIPTION
        IsInstalled is the manifest detection, widened from "this version" to
        "this version or newer". IsInstallable holds only where the detection
        finds the product with an older version: a computer without the
        product never gets the update. A detection that cannot compare the
        installed version (text, existence only, a per-version MSI product
        key, a script) throws with the change that makes it publishable.
        There is no fallback that installs where the product is missing.

        A manifest WsusDetection block (type ArpEntry) replaces the manifest
        detection for WSUS only: the rules read the product's Add/Remove
        Programs entry. ConfigMgr and Intune keep the manifest detection.
        When the manifest detection cannot compare the installed version and
        ContentFolder holds the staged MSI, the rules read the Add/Remove
        Programs entry that MSI registers.

    .OUTPUTS
        [pscustomobject] IsInstalled, IsInstallable, Source (WsusDetection,
        Detection or MsiEntry), Entry (the ArpEntry for MsiEntry).
    #>
    param([Parameter(Mandatory)]$Manifest, [string]$ContentFolder = '')

    $wsusDetection = Get-WsusMemberValue -InputObject $Manifest -Name 'WsusDetection'
    if ($null -ne $wsusDetection) {
        $rules = ConvertTo-WsusArpEntryRules -Detection $wsusDetection
        return [pscustomobject]@{ IsInstalled = $rules.IsInstalled; IsInstallable = $rules.IsInstallable; Source = 'WsusDetection'; Entry = $wsusDetection }
    }
    try {
        $rules = ConvertTo-WsusManifestDetectionRules -Manifest $Manifest
        return [pscustomobject]@{ IsInstalled = $rules.IsInstalled; IsInstallable = $rules.IsInstallable; Source = 'Detection'; Entry = $null }
    }
    catch {
        $unmappable = $_
        $derived = if ($ContentFolder) { Get-WsusMsiArpDetection -Manifest $Manifest -ContentFolder $ContentFolder } else { $null }
        if ($null -eq $derived) { throw $unmappable }
        $rules = ConvertTo-WsusArpEntryRules -Detection $derived
        return [pscustomobject]@{ IsInstalled = $rules.IsInstalled; IsInstallable = $rules.IsInstallable; Source = 'MsiEntry'; Entry = $derived }
    }
}

function ConvertTo-WsusManifestDetectionRules {
    param([Parameter(Mandatory)]$Manifest)

    $detection = Get-WsusMemberValue -InputObject $Manifest -Name 'Detection'
    if ($null -eq $detection) { throw 'The manifest carries no detection, so WSUS could neither find an older version nor report the update installed.' }
    $shape = Get-WsusDetectionClauses -Detection $detection
    $clauses = @($shape.Clauses | ForEach-Object { ConvertTo-WsusClauseRules -Clause $_ })
    $versionClauses = @($clauses | Where-Object { $_.Kind -eq 'Version' })
    $presenceClauses = @($clauses | Where-Object { $_.Kind -eq 'Presence' })
    $existenceOnly = 'only checks that a file or registry key exists, which says nothing about the installed version. Add a file version or a registry version comparison to the detection, or a WsusDetection block that names the product''s Add/Remove Programs entry. A new installation belongs to a Configuration Manager or Intune deployment.'

    $installed = Join-WsusRules -Connector $shape.Connector -Rules @($clauses | ForEach-Object { $_.Installed })
    if ($shape.Connector -eq 'Or') {
        if ($presenceClauses.Count -gt 0) {
            throw ('The detection joins its clauses with OR, and at least one clause {0}' -f $existenceOnly)
        }
        $installable = Join-WsusRules -Connector Or -Rules @($versionClauses | ForEach-Object { $_.Older })
    }
    else {
        if ($versionClauses.Count -eq 0) { throw ('The detection {0}' -f $existenceOnly) }
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($clause in $presenceClauses) { $parts.Add($clause.Installed) }
        if ($versionClauses.Count -eq 1) {
            $parts.Add($versionClauses[0].Older)
        }
        else {
            # Every version clause must find its value, and at least one must
            # find an older version. For any target above version 0, Older or
            # Installed holds only where the value exists.
            foreach ($clause in $versionClauses) { $parts.Add((Join-WsusRules -Connector Or -Rules @($clause.Older, $clause.Installed))) }
            $parts.Add((Join-WsusRules -Connector Or -Rules @($versionClauses | ForEach-Object { $_.Older })))
        }
        $installable = Join-WsusRules -Connector And -Rules $parts.ToArray()
    }
    return [pscustomobject]@{ IsInstalled = $installed; IsInstallable = $installable }
}

# ---------------------------------------------------------------------------
# Compatibility findings
# ---------------------------------------------------------------------------

function Get-WsusCompatibilityFindings {
    <#
    .SYNOPSIS
        Reports what a WSUS publish of this manifest would and would not do.

    .DESCRIPTION
        The publisher reads this list before it connects, so a blocking gap
        never leaves a half-published update. Severity Blocking refuses the
        publish; Review needs an operator decision; Info records a capability
        the destination does not carry.

    .OUTPUTS
        [pscustomobject[]] Severity, Code, Message.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [string]$ContentFolder = ''
    )

    $findings = New-Object System.Collections.Generic.List[object]
    $add = { param($Severity, $Code, $Message) $findings.Add([pscustomobject]@{ Severity = $Severity; Code = $Code; Message = $Message }) }

    $deploymentTypes = @(Get-WsusMemberValue -InputObject $Manifest -Name 'DeploymentTypes' | Where-Object { $null -ne $_ })
    if ($deploymentTypes.Count -gt 0) {
        & $add 'Blocking' 'MultipleDeploymentTypes' 'This application has deployment-type variants. A WSUS update carries one installer, so the variant set cannot be published as one update.'
    }
    $installerType = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerType')
    if ($installerType -notin @('MSI', 'EXE')) {
        & $add 'Blocking' 'InstallerTypeUnsupported' ("Installer type '{0}' cannot be published: WSUS installs Windows Installer packages and command-line executables only." -f $installerType)
    }
    if ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallationBehaviorType') -eq 'InstallForUser') {
        & $add 'Blocking' 'PerUserInstall' 'This application installs for the signed-in user. WSUS installs as SYSTEM, so a per-user installer would land in the wrong profile.'
    }
    if ((Get-WsusMemberValue -InputObject $Manifest -Name 'RequireUserInteraction') -eq $true) {
        & $add 'Blocking' 'InteractiveInstall' 'This deployment type lets the user interact with the installation. WSUS installs silently.'
    }

    $installCommand = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallCommandLine')).Trim().Trim('"')
    $customInstall = @(@(Get-WsusMemberValue -InputObject $Manifest -Name 'CustomAssets') | Where-Object {
            $null -ne $_ -and [string](Get-WsusMemberValue -InputObject $_ -Name 'Category') -like 'Install*'
        })
    $overrideInstall = [string](Get-WsusMemberValue -InputObject (Get-WsusMemberValue -InputObject $Manifest -Name 'CommandOverrides') -Name 'Install')
    if (($installCommand -and $installCommand -ne 'install.bat') -or $customInstall.Count -gt 0 -or -not [string]::IsNullOrWhiteSpace($overrideInstall)) {
        & $add 'Blocking' 'CustomInstall' 'This build installs through a custom command or script. WSUS runs the vendor installer with its silent arguments, so the customization would not reach the client.'
    }
    elseif ($ContentFolder -and (Test-Path -LiteralPath ([System.IO.Path]::Combine($ContentFolder, 'install.ps1')) -PathType Leaf)) {
        $extras = @(Get-WsusInstallScriptExtras -Path ([System.IO.Path]::Combine($ContentFolder, 'install.ps1')))
        if ($extras.Count -gt 0) {
            & $add 'Blocking' 'CustomInstall' ("install.ps1 does more than run the installer: {0}. WSUS runs only the installer with its arguments, so these steps would not reach the client." -f ($extras -join ', '))
        }
    }

    $installerFile = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerFile')
    $outside = @()
    if ($ContentFolder) {
        foreach ($file in @(Get-WsusPayloadFiles -Manifest $Manifest | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            try { [void](Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $file) }
            catch { $outside += $file }
        }
    }
    if ($outside.Count -gt 0) {
        & $add 'Blocking' 'PayloadOutsideStage' ("The stage manifest names {0}, which is not a file inside the content folder." -f (($outside | ForEach-Object { "'" + $_ + "'" }) -join ', '))
    }
    if ([string]::IsNullOrWhiteSpace($installerFile)) {
        & $add 'Blocking' 'InstallerMissing' 'The manifest names no installer file.'
    }
    elseif ($outside -contains $installerFile) { }
    elseif ($ContentFolder -and -not (Test-Path -LiteralPath ([System.IO.Path]::Combine($ContentFolder, $installerFile)) -PathType Leaf)) {
        & $add 'Blocking' 'InstallerMissing' ("The installer '{0}' is not in the staged content." -f $installerFile)
    }
    elseif ($installerFile -match '[\\/]') {
        & $add 'Blocking' 'InstallerInSubfolder' ("The installer '{0}' sits in a subfolder; a WSUS update carries its files in one folder." -f $installerFile)
    }

    if ($ContentFolder -and @($findings | Where-Object { $_.Severity -eq 'Blocking' }).Count -eq 0) {
        try {
            $rules = ConvertTo-WsusApplicabilityRules -Manifest $Manifest -ContentFolder $ContentFolder
            if ($rules.Source -eq 'MsiEntry') {
                & $add 'Info' 'WsusRulesFromMsi' ("The detection cannot tell an older installed version apart, so WSUS reads the Add/Remove Programs entry that the MSI registers: {0}." -f (Format-WsusArpEntry -Entry $rules.Entry))
            }
        }
        catch { & $add 'Blocking' 'DetectionNotMappable' $_.Exception.Message }

        $payload = @(Get-WsusPayloadFiles -Manifest $Manifest)
        $nested = @($payload | Select-Object -Skip 1 | Where-Object { $_.Contains('\') })
        if ($nested.Count -gt 0) {
            & $add 'Blocking' 'PayloadInSubfolder' ("The staged content carries {0} in a subfolder; a WSUS update carries its files in one folder." -f ($nested -join ', '))
        }
        $missing = @($payload | Select-Object -Skip 1 | Where-Object { -not (Test-Path -LiteralPath (Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $_) -PathType Leaf) })
        if ($missing.Count -gt 0) {
            & $add 'Blocking' 'PayloadMissing' ("The stage recorded {0}, which is no longer in the staged content; re-run the Stage phase." -f ($missing -join ', '))
        }
        $megabytes = 0.0
        foreach ($file in @($payload | ForEach-Object { Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $_ } | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })) {
            $megabytes += (Get-Item -LiteralPath $file).Length / 1MB
        }
        if ($megabytes -gt $script:WsusCabHardLimitMegabytes) {
            & $add 'Blocking' 'PayloadTooLarge' ("The installer payload is {0:N0} MB. A locally published update is one signed cab, and a cab cannot exceed {1} MB." -f $megabytes, $script:WsusCabHardLimitMegabytes)
        }
    }

    $requirements = @(Get-WsusMemberValue -InputObject $Manifest -Name 'Requirements' | Where-Object { $null -ne $_ })
    if ($requirements.Count -gt 0) {
        & $add 'Review' 'RequirementsNotTranslated' ("This application carries {0} ConfigMgr requirement rule(s). The WSUS publisher does not translate them; approve the update only for computer groups that meet them." -f $requirements.Count)
    }
    $running = @(Get-WsusMemberValue -InputObject $Manifest -Name 'RunningProcess' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($running.Count -gt 0) {
        & $add 'Review' 'RunningProcessNotClosed' ("WSUS cannot close {0} before the install; the installer's own handling of a running copy applies." -f ($running -join ', '))
    }
    if ($null -ne (Get-WsusMemberValue -InputObject $Manifest -Name 'Timing')) {
        & $add 'Info' 'TimingNotApplied' 'Estimated and maximum runtime minutes are not sent to WSUS; the client applies its own installation timeout.'
    }
    & $add 'Info' 'UninstallNotPublished' 'WSUS publishes the install only; removal stays with ConfigMgr, Intune, or the vendor uninstaller.'

    return @($findings.ToArray())
}

# ---------------------------------------------------------------------------
# WSUS API adapter
# ---------------------------------------------------------------------------
# Every call into Microsoft.UpdateServices.Administration lives in this
# section. Type names resolve at run time, after the assembly loads, so the
# module imports on a host without the API.

function Get-WsusErrorMessage {
    # API failures arrive wrapped in MethodInvocationException; the inner
    # exception carries the server's own text.
    param([Parameter(Mandatory)]$ErrorRecord)
    $exception = $ErrorRecord
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $exception = $ErrorRecord.Exception }
    while ($exception -is [System.Management.Automation.MethodInvocationException] -and $exception.InnerException) {
        $exception = $exception.InnerException
    }
    return [string]$exception.Message
}

function Import-WsusAdministrationAssembly {
    $loaded = @([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq $script:WsusAdministrationAssemblyName })
    if ($loaded.Count -gt 0) { return $loaded[0] }
    if ($PSVersionTable.PSEdition -eq 'Core') {
        throw 'The WSUS administration API is a .NET Framework assembly and does not load in PowerShell 7. Run AppPackager under Windows PowerShell 5.1 (powershell.exe).'
    }
    $assembly = $null
    try { $assembly = [System.Reflection.Assembly]::LoadWithPartialName($script:WsusAdministrationAssemblyName) } catch { $assembly = $null }
    if ($assembly) { return $assembly }
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $path = [System.IO.Path]::Combine([string]$programFiles, 'Update Services\Api\Microsoft.UpdateServices.Administration.dll')
    if (Test-Path -LiteralPath $path -PathType Leaf) { return [System.Reflection.Assembly]::LoadFrom($path) }
    throw ('The WSUS administration API is not installed on this computer. Install the WSUS console: on Windows 10 or 11, ' +
        'Settings > System > Optional features > RSAT: Windows Server Update Services Tools (Add-WindowsCapability -Online -Name Rsat.WSUS.Tools~~~~0.0.1.0); ' +
        'on Windows Server, Install-WindowsFeature UpdateServices-UI.')
}

function Test-WsusAdministrationApi {
    <#
    .SYNOPSIS
        Reports whether this host can load the WSUS administration API.

    .OUTPUTS
        [pscustomobject] Available, Version, Location, Reason.
    #>
    try {
        $assembly = Import-WsusAdministrationAssembly
        return [pscustomobject]@{ Available = $true; Version = [string]$assembly.GetName().Version; Location = [string]$assembly.Location; Reason = '' }
    }
    catch {
        return [pscustomobject]@{ Available = $false; Version = ''; Location = ''; Reason = (Get-WsusErrorMessage -ErrorRecord $_) }
    }
}

function Test-WsusServerReachable {
    # A short TCP probe first: the API's own connect waits out a full HTTP
    # timeout against an address that never answers.
    param([Parameter(Mandatory)][string]$ServerName, [Parameter(Mandatory)][int]$PortNumber, [int]$TimeoutMilliseconds = 5000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $pending = $client.BeginConnect($ServerName, $PortNumber, $null, $null)
        if (-not $pending.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
            throw ('WSUS server {0} did not answer on port {1} within {2} seconds.' -f $ServerName, $PortNumber, [int]($TimeoutMilliseconds / 1000))
        }
        try { $client.EndConnect($pending) }
        catch { throw ('WSUS server {0} refused port {1}: {2}' -f $ServerName, $PortNumber, (Get-WsusErrorMessage -ErrorRecord $_)) }
    }
    finally { $client.Close() }
}

function Test-WsusServerIsLocal {
    # The API reports only its parameterless local connection as secure; a
    # connection to this computer by name reports insecure and would refuse
    # a PFX import on the WSUS server itself.
    param([Parameter(Mandatory)][string]$ServerName)
    $name = $ServerName.Trim().TrimEnd('.')
    if ($name -in @('localhost', '127.0.0.1')) { return $true }
    if ($name -ieq $env:COMPUTERNAME) { return $true }
    $fqdn = ''
    try { $fqdn = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName } catch { $fqdn = '' }
    return ([bool]$fqdn -and $name -ieq $fqdn)
}

function New-WsusServerObject {
    param([Parameter(Mandatory)][string]$ServerName, [Parameter(Mandatory)][bool]$UseSsl, [Parameter(Mandatory)][int]$PortNumber)
    if (Test-WsusServerIsLocal -ServerName $ServerName) {
        $server = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
    }
    else {
        $server = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ServerName, $UseSsl, $PortNumber)
    }
    $server.PreferredCulture = 'en'
    return $server
}

function Get-WsusServerConnection {
    param([Parameter(Mandatory)]$Settings)
    $normalized = ConvertTo-WsusPublishSettings -InputObject $Settings
    if ([string]::IsNullOrWhiteSpace($normalized.ServerName)) { throw 'No WSUS server is configured.' }
    [void](Import-WsusAdministrationAssembly)
    Test-WsusServerReachable -ServerName $normalized.ServerName -PortNumber $normalized.PortNumber
    try { return (New-WsusServerObject -ServerName $normalized.ServerName -UseSsl $normalized.UseSsl -PortNumber $normalized.PortNumber) }
    catch {
        $message = Get-WsusErrorMessage -ErrorRecord $_
        $hint = 'The WSUS console on this computer must match the server version.'
        if ($normalized.UseSsl -and $message -match 'SSL|TLS|secure channel|trust relationship') {
            $hint = 'For SSL, this computer must trust the server certificate and use TLS 1.2 (SchUseStrongCrypto = 1 under HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319).'
        }
        throw ('Connecting to WSUS {0}:{1} failed: {2}. {3}' -f $normalized.ServerName, $normalized.PortNumber, $message, $hint)
    }
}

function Get-WsusUserRoleName {
    param([Parameter(Mandatory)]$Server)
    return [string]$Server.GetCurrentUserRole()
}

function Assert-WsusAdministrator {
    # Publishing, approval and certificate changes need WSUS Administrators
    # membership; the role check names that before a call fails halfway.
    param([Parameter(Mandatory)]$Server)
    $role = Get-WsusUserRoleName -Server $Server
    if ($role -ne 'Administrator') {
        throw ('The account {0}\{1} has the role {2} on WSUS {3}. Add it to the WSUS Administrators group on the server.' -f $env:USERDOMAIN, $env:USERNAME, $role, [string]$Server.Name)
    }
}

function Get-WsusSigningCertificateObject {
    # The API saves the public certificate to a file on this computer rather
    # than returning it. $null when the server hands back no certificate.
    param([Parameter(Mandatory)]$Server)
    $path = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('apwsus-cert-' + [guid]::NewGuid().ToString('N') + '.cer'))
    try {
        try { $Server.GetConfiguration().GetSigningCertificate($path) }
        catch {
            # A server without a registered certificate answers with the
            # Win32 file-not-found error rather than an empty result.
            $inner = $_.Exception
            while ($inner.InnerException) { $inner = $inner.InnerException }
            $nativeCode = $null
            if ($inner.PSObject.Properties['NativeErrorCode']) { $nativeCode = $inner.NativeErrorCode }
            if ($nativeCode -eq 2 -or $inner.HResult -eq -2147024894 -or $inner.Message -match 'cannot find the file') { return $null }
            throw
        }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Item -LiteralPath $path).Length -eq 0) { return $null }
        return (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, [System.IO.File]::ReadAllBytes($path)))
    }
    finally { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
}

function Set-WsusSigningCertificateCore {
    param([Parameter(Mandatory)]$Server, [string]$PfxPath = '', [System.Security.SecureString]$Password)
    $configuration = $Server.GetConfiguration()
    if ([string]::IsNullOrWhiteSpace($PfxPath)) {
        try { $configuration.SetSigningCertificate() }
        catch {
            throw ('WSUS did not create a self-signed certificate: {0}. WSUS on Windows Server 2012 R2 and later creates one only when HKLM\SOFTWARE\Microsoft\Update Services\Server\Setup\EnableSelfSignedCertificates = 1 (DWORD); Microsoft treats this path as deprecated. A code-signing certificate from your PKI, imported with Import PFX, is the supported path.' -f (Get-WsusErrorMessage -ErrorRecord $_))
        }
    }
    else {
        # The PFX bytes and the password travel inside the API request.
        if (-not [bool]$Server.IsConnectionSecureForApiRemoting) {
            throw 'Importing a PFX sends its private key and password to the server inside the API call, and this connection is not secure. Connect with SSL, or run AppPackager on the WSUS server.'
        }
        $configuration.SetSigningCertificate($PfxPath, $Password)
    }
    $configuration.Save()
}

function Get-WsusLocalPublishingCabLimitMegabytes {
    param([Parameter(Mandatory)]$Server)
    return [int]$Server.GetConfiguration().LocalPublishingMaxCabSize
}

function Get-WsusUpdateObject {
    # $null when no update carries this id.
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][guid]$PackageId)
    try { return $Server.GetUpdate((New-Object Microsoft.UpdateServices.Administration.UpdateRevisionId($PackageId))) }
    catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        if ($inner.GetType().Name -eq 'WsusObjectNotFoundException' -or (Get-WsusErrorMessage -ErrorRecord $_) -match 'not found|could not be found|does not exist') { return $null }
        throw
    }
}

function Get-WsusLocallyPublishedUpdateObjects {
    param([Parameter(Mandatory)]$Server)
    $scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $scope.UpdateSources = [Microsoft.UpdateServices.Administration.UpdateSources]::Other
    return @($Server.GetUpdates($scope))
}

function Get-WsusComputerGroupObjects {
    param([Parameter(Mandatory)]$Server)
    return @($Server.GetComputerTargetGroups())
}

function New-WsusSoftwareDistributionPackageFile {
    # Builds the Software Distribution Package through the API class, which
    # computes the file digests and writes schema-valid XML.
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$Path)

    $sdp = New-Object Microsoft.UpdateServices.Administration.SoftwareDistributionPackage
    $installerPath = [System.IO.Path]::Combine($Plan.SourceFolder, $Plan.InstallerFile)
    if ($Plan.InstallerType -eq 'MSI') { $sdp.PopulatePackageFromWindowsInstaller($installerPath) }
    else { $sdp.PopulatePackageFromExe($installerPath) }

    $sdp.PackageId = $Plan.PackageId
    # Every package is a WSUS Update: WSUS files an Application package under
    # the Applications classification, which a Configuration Manager
    # software update point cannot subscribe to. Classification is an
    # Update-only property: the type is set first.
    $sdp.PackageType = [Microsoft.UpdateServices.Administration.PackageType]::Update
    $sdp.PackageUpdateClassification = [Microsoft.UpdateServices.Administration.PackageUpdateClassification]$Plan.Classification
    $sdp.Title = $Plan.Title
    $sdp.Description = $Plan.Description
    $sdp.VendorName = $Plan.VendorName
    $sdp.ProductNames.Clear()
    [void]$sdp.ProductNames.Add($Plan.ProductName)
    $supportUri = $null
    if ($Plan.SupportUrl -and [Uri]::TryCreate([string]$Plan.SupportUrl, [UriKind]::Absolute, [ref]$supportUri)) { $sdp.SupportUrl = $supportUri }
    $sdp.IsInstallable = $Plan.IsInstallable
    foreach ($superseded in @($Plan.SupersededPackageIds)) { $sdp.SupersededPackages.Add([guid]$superseded) }

    # Populate writes item-level rules (InstalledOnce for an EXE,
    # MsiApplicationInstalled for an MSI) that are ANDed with any package
    # rule; the item rules are replaced so detection means what the manifest
    # says.
    $item = $sdp.InstallableItems[0]
    $item.IsInstallableApplicabilityRule = $Plan.IsInstallable
    $item.IsInstalledApplicabilityRule = $Plan.IsInstalled
    if ($Plan.InstallerType -eq 'MSI') {
        $item.InstallCommandLine = $Plan.CommandLine
    }
    else {
        $item.Arguments = $Plan.CommandLine
        $item.RebootByDefault = $false
        $item.DefaultResult = [Microsoft.UpdateServices.Administration.InstallationResult]::Failed
        $item.ReturnCodes.Clear()
        foreach ($code in @($Plan.ReturnCodes)) {
            $returnCode = New-Object Microsoft.UpdateServices.Administration.ReturnCode
            $returnCode.ReturnCodeValue = [int]$code.Code
            $returnCode.InstallationResult = [Microsoft.UpdateServices.Administration.InstallationResult]$code.Result
            $returnCode.IsRebootRequired = [bool]$code.Reboot
            $item.ReturnCodes.Add($returnCode)
        }
        # An item that can request user input is excluded from automatic
        # installation.
        if ($null -eq $item.InstallBehavior) { $item.InstallBehavior = New-Object Microsoft.UpdateServices.Administration.InstallBehavior }
        $item.InstallBehavior.CanRequestUserInput = $false
        $item.InstallBehavior.Impact = [Microsoft.UpdateServices.Administration.InstallationImpact]::Normal
        $item.InstallBehavior.RebootBehavior = [Microsoft.UpdateServices.Administration.RebootBehavior]::CanRequestReboot
    }
    $sdp.Save($Path)
}

function Invoke-WsusPublisher {
    # The package folder on the server takes the package id as its name, so
    # a later removal finds it whatever the server's default naming is.
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][string]$SdpPath, [Parameter(Mandatory)][string]$SourceFolder, [Parameter(Mandatory)][guid]$PackageId)
    $publisher = $Server.GetPublisher($SdpPath)
    $publisher.PublishPackage($SourceFolder, $PackageId.ToString())
}

function Invoke-WsusUpdateApproval {
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)]$Update, [Parameter(Mandatory)][string]$GroupName)
    $group = @(Get-WsusComputerGroupObjects -Server $Server | Where-Object { [string]$_.Name -eq $GroupName }) | Select-Object -First 1
    if (-not $group) { throw ("WSUS has no computer group named '{0}'." -f $GroupName) }
    [void]$Update.Approve([Microsoft.UpdateServices.Administration.UpdateApprovalAction]::Install, $group)
}

function Invoke-WsusUpdateDecline {
    param([Parameter(Mandatory)]$Update)
    $Update.Decline()
}

function Invoke-WsusUpdateExpiry {
    param([Parameter(Mandatory)]$Update)
    $Update.ExpirePackage()
}

function Invoke-WsusUpdateDeletion {
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][guid]$UpdateId)
    $Server.DeleteUpdate($UpdateId)
}

function Remove-WsusPackageFolder {
    # Deleting an update leaves its signed cab in the server's
    # UpdateServicesPackages share. The folder is named by the package id
    # and belongs to that one update.
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][guid]$PackageId)
    $folder = '\\{0}\UpdateServicesPackages\{1}' -f [string]$Server.Name, $PackageId.ToString()
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return $false }
    Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
    return $true
}

function Invoke-WsusCatalogImport {
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][guid]$UpdateId)
    $Server.ImportUpdateFromCatalogSite($UpdateId, [string[]]@())
}

# ---------------------------------------------------------------------------
# Server operations
# ---------------------------------------------------------------------------

function ConvertTo-WsusCertificateInfo {
    param($Certificate)
    if ($null -eq $Certificate) { return $null }
    $keyLength = 0
    try { $keyLength = [int]$Certificate.PublicKey.Key.KeySize } catch { $keyLength = 0 }
    return [pscustomobject]@{
        Subject    = [string]$Certificate.Subject
        Issuer     = [string]$Certificate.Issuer
        Thumbprint = [string]$Certificate.Thumbprint
        NotBefore  = $Certificate.NotBefore
        NotAfter   = $Certificate.NotAfter
        KeyLength  = $keyLength
        SelfSigned = ([string]$Certificate.Subject -eq [string]$Certificate.Issuer)
        Expired    = ($Certificate.NotAfter -lt (Get-Date))
    }
}

function Get-WsusServerStatus {
    <#
    .SYNOPSIS
        Connects to the configured server and reads its version and signing
        certificate.

    .OUTPUTS
        [pscustomobject] Connected, ServerName, Version, SigningCertificate,
        Message. Never throws.
    #>
    param([Parameter(Mandatory)]$Settings)

    $name = [string](Get-WsusMemberValue -InputObject $Settings -Name 'ServerName')
    try {
        $server = Get-WsusServerConnection -Settings $Settings
        $notes = New-Object System.Collections.Generic.List[string]
        $role = ''
        try { $role = Get-WsusUserRoleName -Server $server } catch { $notes.Add('The account role could not be read: ' + (Get-WsusErrorMessage -ErrorRecord $_)) }
        if ($role -and $role -ne 'Administrator') { $notes.Add(('This account has the {0} role; publishing needs WSUS Administrators membership.' -f $role)) }
        $certificate = $null
        try { $certificate = ConvertTo-WsusCertificateInfo -Certificate (Get-WsusSigningCertificateObject -Server $server) }
        catch { $notes.Add('The signing certificate could not be read: ' + (Get-WsusErrorMessage -ErrorRecord $_)) }
        $secure = $false
        try { $secure = [bool]$server.IsConnectionSecureForApiRemoting } catch { $secure = $false }
        return [pscustomobject]@{
            Connected          = $true
            ServerName         = [string]$server.Name
            Version            = [string]$server.Version
            Role               = $role
            SecureConnection   = $secure
            SigningCertificate = $certificate
            Message            = ($notes.ToArray() -join ' ')
        }
    }
    catch {
        return [pscustomobject]@{ Connected = $false; ServerName = $name; Version = ''; Role = ''; SecureConnection = $false; SigningCertificate = $null; Message = (Get-WsusErrorMessage -ErrorRecord $_) }
    }
}

function New-WsusSelfSignedSigningCertificate {
    <#
    .SYNOPSIS
        Has the WSUS server create a self-signed signing certificate and use
        it for local publishing.
    #>
    param([Parameter(Mandatory)]$Settings)
    $server = Get-WsusServerConnection -Settings $Settings
    Assert-WsusAdministrator -Server $server
    Set-WsusSigningCertificateCore -Server $server
    Write-WsusLog ('WSUS signing certificate      : self-signed certificate created on {0}' -f [string]$server.Name)
    return (Get-WsusServerStatus -Settings $Settings)
}

function Set-WsusSigningCertificate {
    <#
    .SYNOPSIS
        Makes the certificate in a PFX file the server's signing certificate.
    #>
    param(
        [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)][string]$PfxPath,
        [Parameter(Mandatory)][System.Security.SecureString]$Password
    )
    if (-not (Test-Path -LiteralPath $PfxPath -PathType Leaf)) { throw "PFX file not found: $PfxPath" }
    # Opening the file here proves the password, the private key and the
    # code-signing usage before the server's configuration changes.
    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
    try { $probe = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $Password, $flags) }
    catch { throw ('The PFX file could not be opened: {0}' -f (Get-WsusErrorMessage -ErrorRecord $_)) }
    try {
        if (-not $probe.HasPrivateKey) { throw 'The PFX file carries no private key; WSUS signs with the private key.' }
        $usages = @($probe.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension] } |
            ForEach-Object { $_.EnhancedKeyUsages } | ForEach-Object { [string]$_.Value })
        if ($usages.Count -gt 0 -and $usages -notcontains '1.3.6.1.5.5.7.3.3') {
            throw 'The certificate is not valid for code signing (enhanced key usage 1.3.6.1.5.5.7.3.3).'
        }
        $keyUsage = @($probe.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension] }) | Select-Object -First 1
        if ($keyUsage -and -not ($keyUsage.KeyUsages -band [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature)) {
            throw 'The certificate key usage does not include Digital Signature.'
        }
        # Clients reject update signatures from keys under 1024 bits, and the
        # publishing guidance for this workload asks for 2048 or more.
        $keySize = 0
        try { $keySize = [int]$probe.PublicKey.Key.KeySize } catch { $keySize = 0 }
        if ($keySize -gt 0 -and $keySize -lt 2048) {
            throw ('The certificate key is {0} bits; WSUS update signing needs an RSA key of 2048 bits or more.' -f $keySize)
        }
        if ($probe.NotAfter -lt (Get-Date)) { throw ('The certificate expired on {0:yyyy-MM-dd}.' -f $probe.NotAfter) }
    }
    finally { $probe.Reset() }

    $server = Get-WsusServerConnection -Settings $Settings
    Assert-WsusAdministrator -Server $server
    Set-WsusSigningCertificateCore -Server $server -PfxPath $PfxPath -Password $Password
    Write-WsusLog ('WSUS signing certificate      : {0} set on {1}' -f [System.IO.Path]::GetFileName($PfxPath), [string]$server.Name)
    return (Get-WsusServerStatus -Settings $Settings)
}

function Export-WsusSigningCertificate {
    <#
    .SYNOPSIS
        Writes the server's public signing certificate (DER .cer) for client
        distribution.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$Path)
    $server = Get-WsusServerConnection -Settings $Settings
    $certificate = Get-WsusSigningCertificateObject -Server $server
    if (-not $certificate) { throw ('WSUS server {0} has no signing certificate to export.' -f [string]$server.Name) }
    [System.IO.File]::WriteAllBytes($Path, $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    return $Path
}

function Get-WsusComputerGroupNames {
    <#
    .SYNOPSIS
        Names of the server's computer target groups, sorted.
    #>
    param([Parameter(Mandatory)]$Settings)
    $server = Get-WsusServerConnection -Settings $Settings
    return @(Get-WsusComputerGroupObjects -Server $server | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
}

function ConvertTo-WsusUpdateRow {
    param([Parameter(Mandatory)]$Update, [hashtable]$GroupNames = @{})
    $identity = Get-WsusUpdateIdentity -Description ([string]$Update.Description)
    $approved = @()
    try {
        foreach ($approval in @($Update.GetUpdateApprovals())) {
            if ([string]$approval.Action -ne 'Install') { continue }
            $groupId = [string]$approval.ComputerTargetGroupId
            $approved += $(if ($GroupNames.ContainsKey($groupId)) { $GroupNames[$groupId] } else { $groupId })
        }
    }
    catch { $approved = @('unknown') }
    $created = $null
    try { $created = [datetime]$Update.CreationDate } catch { $created = $null }
    return [pscustomobject]@{
        PackageId      = [string]$Update.Id.UpdateId
        Title          = [string]$Update.Title
        PackageType    = $(if ($identity) { $identity.PackageType } else { '' })
        IdentityTag    = $(if ($identity) { $identity.IdentityTag } else { '' })
        Version        = $(if ($identity) { $identity.Version } else { '' })
        Created        = $created
        CreatedText    = $(if ($created) { $created.ToString('yyyy-MM-dd HH:mm') } else { '' })
        ApprovedGroups = ($approved -join ', ')
        Declined       = [bool]$Update.IsDeclined
        Superseded     = [bool]$Update.IsSuperseded
        Expired        = ([string](Get-WsusMemberValue -InputObject $Update -Name 'PublicationState') -eq 'Expired')
    }
}

function Get-WsusPublishedUpdates {
    <#
    .SYNOPSIS
        Lists locally published updates on the server, newest first.

    .DESCRIPTION
        By default only the updates that carry an AppPackager identity line;
        -IncludeOtherPublishers lists every locally published update.
    #>
    param([Parameter(Mandatory)]$Settings, [switch]$IncludeOtherPublishers)
    $server = Get-WsusServerConnection -Settings $Settings
    $groupNames = @{}
    try { foreach ($group in @(Get-WsusComputerGroupObjects -Server $server)) { $groupNames[[string]$group.Id] = [string]$group.Name } } catch { $groupNames = @{} }
    $rows = foreach ($update in @(Get-WsusLocallyPublishedUpdateObjects -Server $server)) {
        $row = ConvertTo-WsusUpdateRow -Update $update -GroupNames $groupNames
        if ($IncludeOtherPublishers -or $row.IdentityTag) { $row }
    }
    return @(@($rows) | Sort-Object -Property Created -Descending)
}

function Set-WsusPublishedUpdateState {
    <#
    .SYNOPSIS
        Approves for a group, declines, expires, or removes locally published
        updates.

    .DESCRIPTION
        One connection serves every id. Each id succeeds or fails on its own;
        only a connection or permission failure throws. Remove declines the
        update, deletes it from the server database, and then deletes its
        package folder on the server.

    .OUTPUTS
        [pscustomobject[]] PackageId, Title, Ok, Message.
    #>
    param(
        [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)][guid[]]$PackageId,
        [Parameter(Mandatory)][ValidateSet('Approve', 'Decline', 'Expire', 'Remove')][string]$Action,
        [string]$GroupName = ''
    )
    if ($Action -eq 'Approve' -and [string]::IsNullOrWhiteSpace($GroupName)) { throw 'Approve needs a computer group.' }
    $server = Get-WsusServerConnection -Settings $Settings
    Assert-WsusAdministrator -Server $server

    $results = New-Object System.Collections.Generic.List[object]
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($id in @($PackageId | Select-Object -Unique)) {
        $update = Get-WsusUpdateObject -Server $server -PackageId $id
        if ($update) { $targets.Add([pscustomobject]@{ Id = $id; Update = $update }) }
        else { $results.Add([pscustomobject]@{ PackageId = $id; Title = ''; Ok = $false; Message = ('WSUS has no update with id {0}.' -f $id) }) }
    }
    foreach ($target in $targets) {
        $update = $target.Update
        $title = [string]$update.Title
        $note = ''
        $declinedNow = $false
        try {
            switch ($Action) {
                'Approve' { Invoke-WsusUpdateApproval -Server $server -Update $update -GroupName $GroupName }
                'Decline' { Invoke-WsusUpdateDecline -Update $update }
                'Expire'  { Invoke-WsusUpdateExpiry -Update $update }
                'Remove'  {
                    if (-not [bool]$update.IsDeclined) {
                        Invoke-WsusUpdateDecline -Update $update
                        $declinedNow = $true
                    }
                    Invoke-WsusUpdateDeletion -Server $server -UpdateId $target.Id
                    try { [void](Remove-WsusPackageFolder -Server $server -PackageId $target.Id) }
                    catch { $note = 'The update is removed; its package folder on the server was not deleted: ' + (Get-WsusErrorMessage -ErrorRecord $_) }
                }
            }
            Write-WsusLog ('WSUS update {0,-17} : {1} ({2})' -f $Action.ToLowerInvariant(), $title, $target.Id)
            $results.Add([pscustomobject]@{ PackageId = $target.Id; Title = $title; Ok = $true; Message = $note })
        }
        catch {
            $message = Get-WsusErrorMessage -ErrorRecord $_
            if ($Action -eq 'Remove' -and $message -match 'referenced') {
                $message += ' Another update on the server still references this update. Expire this update instead, or remove the update that references it first.'
            }
            elseif ($Action -eq 'Expire' -and $message -match 'already (been )?expired|not imported|SDP') {
                $message += ' Only a locally published update that is not already expired can be expired on this server.'
            }
            if ($declinedNow) { $message += ' The update is now declined.' }
            $results.Add([pscustomobject]@{ PackageId = $target.Id; Title = $title; Ok = $false; Message = $message })
        }
    }
    return @($results.ToArray())
}

function Import-WsusCatalogUpdate {
    <#
    .SYNOPSIS
        Imports one Microsoft Update Catalog update into WSUS by update ID.

    .DESCRIPTION
        The server downloads the update metadata from the catalog itself; the
        update files follow the server's Update files setting. Never throws.

    .OUTPUTS
        [pscustomobject] UpdateId, Ok, Message.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][guid]$UpdateId)
    try {
        $server = Get-WsusServerConnection -Settings $Settings
        Assert-WsusAdministrator -Server $server
        Invoke-WsusCatalogImport -Server $server -UpdateId $UpdateId
        Write-WsusLog ('WSUS catalog import           : {0} on {1}' -f $UpdateId, [string]$server.Name)
        return [pscustomobject]@{ UpdateId = $UpdateId; Ok = $true; Message = '' }
    }
    catch {
        $message = Get-WsusErrorMessage -ErrorRecord $_
        # The import runs on the WSUS server, which downloads the metadata
        # itself; this computer's TLS settings play no part.
        if ($message -match 'SSL/TLS|secure channel|forcibly closed|underlying connection was closed') {
            $message += ' The WSUS server could not reach the Microsoft Update Catalog over TLS 1.2. On the WSUS server, set SchUseStrongCrypto = 1 (DWORD) under HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319, then restart the WSUS service and the World Wide Web Publishing service. SoftwareDistribution.log in the Update Services LogFiles folder on the server records the import.'
        }
        return [pscustomobject]@{ UpdateId = $UpdateId; Ok = $false; Message = $message }
    }
}

# ---------------------------------------------------------------------------
# Publishing
# ---------------------------------------------------------------------------

function Assert-WsusPayloadIntegrity {
    # The staged bytes must still be the bytes the stage recorded; WSUS signs
    # whatever it is handed.
    param([Parameter(Mandatory)]$Manifest, [Parameter(Mandatory)][string]$ContentFolder, [Parameter(Mandatory)][string[]]$Files)

    $recorded = @{}
    foreach ($entry in @(Get-WsusMemberValue -InputObject $Manifest -Name 'FileHashes')) {
        if ($null -eq $entry) { continue }
        $recorded[([string](Get-WsusMemberValue -InputObject $entry -Name 'RelativePath')).Replace('/', '\').TrimStart('\').ToLowerInvariant()] = [string](Get-WsusMemberValue -InputObject $entry -Name 'Sha256')
    }
    if ($recorded.Count -eq 0) { throw 'The stage manifest records no file hashes; re-run the Stage phase before publishing to WSUS.' }
    foreach ($file in $Files) {
        $expected = $recorded[$file.ToLowerInvariant()]
        if (-not $expected) { throw ("The stage manifest records no hash for '{0}'; re-run the Stage phase." -f $file) }
        $actual = (Get-FileHash -LiteralPath (Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $file) -Algorithm SHA256).Hash
        if ($actual -ne $expected.ToUpperInvariant()) { throw ("'{0}' changed after it was staged (SHA-256 mismatch); re-run the Stage phase." -f $file) }
    }
}

function Get-WsusUpdateTitle {
    # WSUS lists updates by title, so the version is always part of it, and
    # a title stays under the 80 characters the package schema expects.
    param([Parameter(Mandatory)]$Manifest)
    $title = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'AppName')).Trim()
    $version = ([string](Get-WsusMemberValue -InputObject $Manifest -Name 'SoftwareVersion')).Trim()
    if ($version -and $title.IndexOf($version, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { $title = ('{0} {1}' -f $title, $version) }
    if ($title.Length -gt 79) {
        $suffix = if ($version) { ' ' + $version } else { '' }
        $base = $title
        if ($version) { $base = (($base -replace [regex]::Escape($version), '') -replace '\s{2,}', ' ').Trim() }
        $title = $base.Substring(0, [Math]::Min($base.Length, 79 - $suffix.Length - 3)).TrimEnd() + '...' + $suffix
    }
    return $title
}

function Get-WsusPublishFailureHint {
    # The server's own error texts for the common local-publishing failures,
    # mapped to the fix.
    param([AllowEmptyString()][string]$Message)
    switch -Regex ($Message) {
        'Verification of file signature failed' { return 'This computer must trust the WSUS signing certificate: import it into the Local Computer Trusted Publishers store, and into Trusted Root Certification Authorities when it is self-signed. Options, WSUS Publishing, Export certificate saves it.' }
        'not a WSUS Administrator|Unauthorized' { return 'The account must be a member of the WSUS Administrators group on the server.' }
        'CreateDirectory failed' { return 'The UpdateServicesPackages or WSUSContent share on the server is missing, not shared, or not writable.' }
        'Failed to sign package' { return 'The server could not sign the package: check its signing certificate and that it reaches its time stamp server.' }
        'too many locally published categories' { return 'The server holds too many locally published vendor and product categories; remove unused ones on the server.' }
        'version' { return 'The WSUS console on this computer must match the server version.' }
        default { return '' }
    }
}

function Publish-WsusSoftwareUpdate {
    <#
    .SYNOPSIS
        Publishes a staged installer to WSUS as a locally published update.

    .DESCRIPTION
        Compatibility is decided before the first server call. The update
        applies only where the detection finds the product with an older
        version. Each new update gets a new package id. A repeat publish of
        the same version finds the update by its identity line, type and
        version, skips an expired one, and only reapplies the approval and
        the decline of earlier versions, unless the update is declined on the
        server. When the rules read the Add/Remove Programs entry of the
        staged MSI, the result message names that entry. AppPackager updates
        of the same identity and type with a lower
        version are listed as superseded, and declined when the settings ask
        for it. Updates of the retired type Application are left as they are
        and named in the result message. The payload is copied to a private
        folder so WSUS receives exactly the recorded stage content. After a
        successful publish, a failed approval or decline is returned as a
        warning, not thrown.

    .OUTPUTS
        [pscustomobject] PackageId, Title, Outcome (Published |
        AlreadyPublished), Approved, Superseded, Declined, Warnings, Message.
    #>
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)][string]$ContentFolder,
        [Parameter(Mandatory)]$Settings
    )

    $settings = ConvertTo-WsusPublishSettings -InputObject $Settings
    if ([string]::IsNullOrWhiteSpace($settings.ServerName)) { throw 'No WSUS server is configured.' }

    $findings = @(Get-WsusCompatibilityFindings -Manifest $Manifest -ContentFolder $ContentFolder)
    foreach ($finding in $findings) {
        Write-WsusLog ('WSUS compatibility            : [{0}] {1} - {2}' -f $finding.Severity, $finding.Code, $finding.Message) -Level $(if ($finding.Severity -eq 'Info') { 'INFO' } else { 'WARN' })
    }
    $blocking = @($findings | Where-Object { $_.Severity -eq 'Blocking' })
    if ($blocking.Count -gt 0) {
        # Data.WsusRefusal lets a caller tell an application WSUS cannot carry
        # apart from a failed server call.
        $refusal = New-Object System.InvalidOperationException(('This application cannot be published to WSUS: {0}' -f (($blocking | ForEach-Object { '{0}: {1}' -f $_.Code, $_.Message }) -join ' | ')))
        $refusal.Data['WsusRefusal'] = (@($blocking | ForEach-Object { [string]$_.Code }) -join ',')
        throw $refusal
    }

    $payload = @(Get-WsusPayloadFiles -Manifest $Manifest)
    Assert-WsusPayloadIntegrity -Manifest $Manifest -ContentFolder $ContentFolder -Files $payload
    $rules = ConvertTo-WsusApplicabilityRules -Manifest $Manifest -ContentFolder $ContentFolder

    $version = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'SoftwareVersion')
    $identity = Get-WsusIdentityTag -Manifest $Manifest
    $title = Get-WsusUpdateTitle -Manifest $Manifest
    $installerType = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallerType')
    $installArgs = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'InstallArgs')

    $server = Get-WsusServerConnection -Settings $settings
    Assert-WsusAdministrator -Server $server

    # A certificate that cannot be read (DCOM rights on the server's
    # certificate service) is not proof that none exists; the publish call
    # itself refuses when the server truly has none.
    $certificate = $null
    $certificateRead = $true
    try { $certificate = ConvertTo-WsusCertificateInfo -Certificate (Get-WsusSigningCertificateObject -Server $server) }
    catch {
        $certificateRead = $false
        Write-WsusLog ('WSUS signing certificate      : not readable ({0}); publishing continues' -f (Get-WsusErrorMessage -ErrorRecord $_)) -Level WARN
    }
    if ($certificateRead -and -not $certificate) { throw ('WSUS server {0} has no signing certificate. Create or import one in Options, WSUS Publishing.' -f [string]$server.Name) }
    if ($certificate -and $certificate.Expired) { throw ('The WSUS signing certificate {0} expired on {1:yyyy-MM-dd}. Import a new one, then clients must trust it.' -f $certificate.Thumbprint, $certificate.NotAfter) }

    $payloadMegabytes = 0.0
    foreach ($file in $payload) { $payloadMegabytes += (Get-Item -LiteralPath (Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $file)).Length / 1MB }
    $cabLimit = 0
    try { $cabLimit = Get-WsusLocalPublishingCabLimitMegabytes -Server $server } catch { $cabLimit = 0 }
    if ($cabLimit -gt 0 -and $payloadMegabytes -gt $cabLimit) {
        Write-WsusLog ('WSUS cab size                 : payload {0:N0} MB exceeds LocalPublishingMaxCabSize {1} MB on the server' -f $payloadMegabytes, $cabLimit) -Level WARN
    }

    # Only a lower version is superseded: a build of an older version,
    # published after a newer one, must not retire the newer update.
    $earlier = New-Object System.Collections.Generic.List[object]
    $retired = 0
    $existing = $null
    foreach ($candidate in @(Get-WsusLocallyPublishedUpdateObjects -Server $server)) {
        $parsed = Get-WsusUpdateIdentity -Description ([string]$candidate.Description)
        if (-not $parsed -or $parsed.IdentityTag -ne $identity) { continue }
        if ($parsed.PackageType -ne $script:WsusPackageType) {
            if ([string](Get-WsusMemberValue -InputObject $candidate -Name 'PublicationState') -ne 'Expired') { $retired++ }
            continue
        }
        if ([string](Get-WsusMemberValue -InputObject $candidate -Name 'PublicationState') -eq 'Expired') { continue }
        $order = Compare-WsusVersion -Left $parsed.Version -Right $version
        if ($order -eq 0) {
            if ($null -eq $existing -or [datetime]$candidate.CreationDate -gt [datetime]$existing.CreationDate) { $existing = $candidate }
            continue
        }
        if ($order -eq -1) { $earlier.Add($candidate); continue }
        $reason = if ($null -eq $order) { 'its version cannot be compared' } else { 'it carries a newer version' }
        Write-WsusLog ('WSUS supersedence             : {0} left as is; {1}' -f [string]$candidate.Title, $reason)
    }

    # A repeat publish of this version finds the update that is not expired.
    # Otherwise the update gets a new id: WSUS restarts a removed id at
    # revision 1, and Configuration Manager keeps the content it recorded for
    # that id and revision, so a reused id installs with a digest mismatch.
    $update = $existing
    $packageId = if ($existing) { [guid]$existing.Id.UpdateId } else { [guid]::NewGuid() }
    $result = [ordered]@{ PackageId = $packageId; Title = $title; Outcome = ''; Approved = ''; Superseded = @(); Declined = @(); Warnings = @(); Message = '' }
    if ($update) {
        $result.Outcome = 'AlreadyPublished'
        Write-WsusLog ('WSUS update exists            : {0} ({1})' -f $title, $packageId)
    }
    else {
        $work = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('apwsus-' + [guid]::NewGuid().ToString('N')))
        $source = [System.IO.Path]::Combine($work, 'content')
        [void](New-Item -ItemType Directory -Path $source -Force)
        try {
            foreach ($file in $payload) {
                Copy-Item -LiteralPath (Resolve-WsusPayloadPath -ContentFolder $ContentFolder -RelativePath $file) -Destination (Resolve-WsusPayloadPath -ContentFolder $source -RelativePath $file) -Force -ErrorAction Stop
            }
            $buildId = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'BuildId')
            $appName = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'AppName')
            $publisherName = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'Publisher')
            $firstLine = $appName
            if ($version -and $appName.IndexOf($version, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { $firstLine += (' ' + $version) }
            if ($publisherName) { $firstLine += (' from {0}' -f $publisherName) }
            $firstLine += ', published by AppPackager'
            if ($buildId) { $firstLine += (' from build {0}' -f $buildId) }
            $plan = [pscustomobject]@{
                PackageId            = $packageId
                Title                = $title
                Description          = (@(($firstLine + '.'), '', (Get-WsusIdentityLine -IdentityTag $identity -PackageType $script:WsusPackageType -Version $version)) -join "`r`n")
                VendorName           = $script:WsusVendorName
                ProductName          = $script:WsusProductName
                Classification       = $settings.Classification
                SupportUrl           = [string](Get-WsusMemberValue -InputObject $Manifest -Name 'VendorUrl')
                IsInstallable        = $rules.IsInstallable
                IsInstalled          = $rules.IsInstalled
                SupersededPackageIds = @($earlier | ForEach-Object { [guid]$_.Id.UpdateId })
                InstallerType        = $installerType
                InstallerFile        = $payload[0]
                SourceFolder         = $source
                CommandLine          = $(if ($installerType -eq 'MSI') { ConvertTo-WsusMsiCommandLine -InstallArgs $installArgs } else { $installArgs })
                ReturnCodes          = @(
                    @{ Code = 0; Result = 'Succeeded'; Reboot = $false }
                    @{ Code = 1707; Result = 'Succeeded'; Reboot = $false }
                    @{ Code = 3010; Result = 'Succeeded'; Reboot = $true }
                    @{ Code = 1641; Result = 'Succeeded'; Reboot = $true }
                )
            }
            $sdpPath = [System.IO.Path]::Combine($work, 'package.xml')
            New-WsusSoftwareDistributionPackageFile -Plan $plan -Path $sdpPath
            Write-WsusLog ('Publishing to WSUS            : {0} ({1}) on {2}' -f $title, $packageId, [string]$server.Name)
            try { Invoke-WsusPublisher -Server $server -SdpPath $sdpPath -SourceFolder $source -PackageId $packageId }
            catch {
                $message = Get-WsusErrorMessage -ErrorRecord $_
                $hint = Get-WsusPublishFailureHint -Message $message
                if (-not $hint -and $cabLimit -gt 0 -and $payloadMegabytes -gt $cabLimit) {
                    $hint = ('The payload is {0:N0} MB and the server allows {1} MB per cab: raise LocalPublishingMaxCabSize on the server (at most 2047).' -f $payloadMegabytes, $cabLimit)
                }
                throw ('WSUS rejected the package: {0}{1}' -f $message, $(if ($hint) { ' ' + $hint } else { '' }))
            }
            $result.Outcome = 'Published'
            $result.Superseded = @($plan.SupersededPackageIds)
        }
        finally {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        }
        $update = Get-WsusUpdateObject -Server $server -PackageId $packageId
    }

    $warnings = New-Object System.Collections.Generic.List[string]
    if ($result.Outcome -eq 'AlreadyPublished' -and $update -and [bool]$update.IsDeclined) {
        $warnings.Add('the update is declined on the server, so its approval and the decline of earlier versions were skipped')
    }
    else {
        if ($settings.ApprovalGroup) {
            try {
                if (-not $update) { throw ('the update is not visible on {0} after publishing' -f [string]$server.Name) }
                Invoke-WsusUpdateApproval -Server $server -Update $update -GroupName $settings.ApprovalGroup
                $result.Approved = $settings.ApprovalGroup
            }
            catch { $warnings.Add(('approval for {0} failed: {1}' -f $settings.ApprovalGroup, (Get-WsusErrorMessage -ErrorRecord $_))) }
        }
        if ($settings.DeclineSuperseded) {
            foreach ($old in $earlier) {
                if ([bool]$old.IsDeclined) { continue }
                try {
                    Invoke-WsusUpdateDecline -Update $old
                    $result.Declined += [string]$old.Id.UpdateId
                }
                catch { $warnings.Add(('declining {0} failed: {1}' -f [string]$old.Title, (Get-WsusErrorMessage -ErrorRecord $_))) }
            }
        }
    }
    $result.Warnings = @($warnings.ToArray())

    $parts = @(('{0} {1}' -f $(if ($result.Outcome -eq 'Published') { 'published' } else { 'already on the server' }), $packageId))
    if ($rules.Source -eq 'MsiEntry') { $parts += ('rules read the Add/Remove Programs entry that the MSI registers: {0}' -f (Format-WsusArpEntry -Entry $rules.Entry)) }
    if ($result.Approved) { $parts += ('approved for {0}' -f $result.Approved) }
    if (@($result.Superseded).Count) { $parts += ('supersedes {0} earlier version(s)' -f @($result.Superseded).Count) }
    if (@($result.Declined).Count) { $parts += ('declined {0} earlier version(s)' -f @($result.Declined).Count) }
    if ($retired -gt 0) { $parts += ('{0} update(s) of the retired type Application stay on the server; expire or remove them in Published updates' -f $retired) }
    $parts += @($result.Warnings)
    $result.Message = ($parts -join '; ')
    Write-WsusLog ('WSUS publish                  : {0} - {1}' -f $title, $result.Message) -Level $(if ($warnings.Count -gt 0 -or $retired -gt 0) { 'WARN' } else { 'INFO' })
    return [pscustomobject]$result
}

Export-ModuleMember -Function @(
    'Get-WsusClassificationNames'
    'ConvertTo-WsusPublishSettings'
    'Get-WsusIdentityTag'
    'Get-WsusIdentityLine'
    'Get-WsusUpdateIdentity'
    'ConvertFrom-WsusCatalogInput'
    'ConvertTo-WsusMsiCommandLine'
    'ConvertTo-WsusApplicabilityRules'
    'Get-WsusCompatibilityFindings'
    'Test-WsusAdministrationApi'
    'Get-WsusServerStatus'
    'New-WsusSelfSignedSigningCertificate'
    'Set-WsusSigningCertificate'
    'Export-WsusSigningCertificate'
    'Get-WsusComputerGroupNames'
    'Publish-WsusSoftwareUpdate'
    'Get-WsusPublishedUpdates'
    'Set-WsusPublishedUpdateState'
    'Import-WsusCatalogUpdate'
)
