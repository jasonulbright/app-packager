<#
.SYNOPSIS
    Application Workbench definition, profile, run-snapshot and build model.

.DESCRIPTION
    Storage, identity, profile revisions, effective-value resolution, legacy
    preference migration, the stage finalization hook, build records and
    portable bundles for AppPackager.

    Effective value precedence, highest last:
        Global -> Packager -> Profile -> Variant -> Target -> Run

    Data root resolution order: APP_PACKAGER_WORKBENCH_ROOT, the
    WorkbenchDataRoot preference, %LOCALAPPDATA%\AppPackagerData\Workbench.
#>

Set-StrictMode -Off

$script:WorkbenchProfileSchemaVersion = 1
$script:WorkbenchBundleSchemaVersion = 1
$script:WorkbenchManifestSchemaVersion = 4
$script:WorkbenchJsonDepth = 12
$script:WorkbenchTokenNames = @('InstallerFile', 'Version', 'ProductCode', 'ContentRoot')

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

function Write-WorkbenchLog {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'DEBUG')][string]$Level = 'INFO')
    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log $Message -Level $Level
        return
    }
    Write-Verbose $Message
}

function New-WorkbenchFolder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
    return $Path
}

function ConvertTo-WorkbenchJson {
    param([Parameter(Mandatory)][AllowNull()]$InputObject)
    # ConvertTo-Json defaults to Depth 2 and silently truncates deeper graphs
    # in Windows PowerShell 5.1, so every serialization here states a depth.
    return ($InputObject | ConvertTo-Json -Depth $script:WorkbenchJsonDepth)
}

function Write-WorkbenchJsonFile {
    <#
    .SYNOPSIS
        Serializes an object to JSON and replaces the target file atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowNull()]$InputObject
    )
    $folder = Split-Path -Path $Path -Parent
    if ($folder) { [void](New-WorkbenchFolder -Path $folder) }
    $json = ConvertTo-WorkbenchJson -InputObject $InputObject
    $temp = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    # UTF8 without a byte order mark keeps the files readable by the 5.1
    # ConvertFrom-Json path and by any external tool that reads them.
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($temp, $json, $encoding)
    $backup = "$Path.$([guid]::NewGuid().ToString('N')).bak"
    try {
        if (Test-Path -LiteralPath $Path) {
            # File.Replace keeps the swap atomic; it needs a real backup path
            # because a null backup argument binds to an empty string here.
            [System.IO.File]::Replace($temp, $Path, $backup)
        }
        else {
            [System.IO.File]::Move($temp, $Path)
        }
    }
    catch {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
        throw
    }
    finally {
        if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue }
    }
    return $Path
}

function Read-WorkbenchJsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $raw = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return ($raw | ConvertFrom-Json)
}

function Get-WorkbenchFileSha256 {
    param([Parameter(Mandatory)][string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try { return ([BitConverter]::ToString($sha.ComputeHash($stream)) -replace '-', '').ToLowerInvariant() }
        finally { $stream.Dispose() }
    }
    finally { $sha.Dispose() }
}

function Get-WorkbenchTextSha256 {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Test-WorkbenchReparsePoint {
    <#
    .SYNOPSIS
        True when the path itself carries the ReparsePoint file attribute.

    .DESCRIPTION
        FileAttributes.ReparsePoint is the documented marker for every
        reparse point kind (symbolic link, junction, mount point); the
        PowerShell LinkType property only names a subset.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $attributes = [System.IO.File]::GetAttributes($Path)
    }
    catch { return $false }
    return (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint)
}

function Test-WorkbenchPathChainForReparse {
    <#
    .SYNOPSIS
        Throws when any existing element between Root and Path is a reparse
        point, so a staged file can never be written through a link.
    #>
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Path
    )
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\')
    $current = [System.IO.Path]::GetFullPath($Path)
    while ($current -and $current.TrimEnd('\').Length -gt $rootFull.Length) {
        if (Test-WorkbenchReparsePoint -Path $current) {
            throw "Source file destination '$Path' resolves through a reparse point at '$current'."
        }
        $current = Split-Path -Path $current -Parent
    }
    if (Test-WorkbenchReparsePoint -Path $rootFull) {
        throw "Stage root '$rootFull' is a reparse point."
    }
}

function ConvertTo-WorkbenchHashtable {
    <#
    .SYNOPSIS
        Deep-copies a PSCustomObject graph into hashtables and arrays.
    #>
    param([Parameter(Mandatory)][AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string] -or $InputObject -is [bool] -or $InputObject -is [int] -or
        $InputObject -is [long] -or $InputObject -is [double] -or $InputObject -is [decimal]) {
        return $InputObject
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $copy = @{}
        foreach ($key in @($InputObject.Keys)) { $copy[[string]$key] = ConvertTo-WorkbenchHashtable -InputObject $InputObject[$key] }
        return $copy
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $items = @()
        foreach ($item in $InputObject) { $items += , (ConvertTo-WorkbenchHashtable -InputObject $item) }
        return , $items
    }
    if ($InputObject -is [psobject] -and $InputObject.PSObject.Properties) {
        $copy = @{}
        foreach ($property in $InputObject.PSObject.Properties) { $copy[$property.Name] = ConvertTo-WorkbenchHashtable -InputObject $property.Value }
        return $copy
    }
    return $InputObject
}

function Get-WorkbenchMember {
    <#
    .SYNOPSIS
        Reads a dotted path from a hashtable or PSCustomObject graph.

    .OUTPUTS
        A result object with Found (bool) and Value, so an explicit $null
        override stays distinguishable from an absent field.
    #>
    param(
        [Parameter(Mandatory)][AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$Path
    )
    $current = $InputObject
    foreach ($segment in ($Path -split '\.')) {
        if ($null -eq $current) { return [pscustomobject]@{ Found = $false; Value = $null } }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($segment)) { return [pscustomobject]@{ Found = $false; Value = $null } }
            $current = $current[$segment]
            continue
        }
        $property = $current.PSObject.Properties[$segment]
        if (-not $property) { return [pscustomobject]@{ Found = $false; Value = $null } }
        $current = $property.Value
    }
    return [pscustomobject]@{ Found = $true; Value = $current }
}

function Set-WorkbenchMember {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$InputObject,
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()]$Value
    )
    $segments = @($Path -split '\.')
    $current = $InputObject
    for ($i = 0; $i -lt $segments.Count - 1; $i++) {
        $segment = $segments[$i]
        if (-not $current.Contains($segment) -or -not ($current[$segment] -is [System.Collections.IDictionary])) {
            $current[$segment] = @{}
        }
        $current = $current[$segment]
    }
    $current[$segments[-1]] = $Value
    return $InputObject
}

function Remove-WorkbenchMember {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$InputObject,
        [Parameter(Mandatory)][string]$Path
    )
    $segments = @($Path -split '\.')
    $current = $InputObject
    for ($i = 0; $i -lt $segments.Count - 1; $i++) {
        $segment = $segments[$i]
        if (-not $current.Contains($segment) -or -not ($current[$segment] -is [System.Collections.IDictionary])) { return $InputObject }
        $current = $current[$segment]
    }
    if ($current.Contains($segments[-1])) { [void]$current.Remove($segments[-1]) }
    return $InputObject
}

# ---------------------------------------------------------------------------
# Storage roots and identity
# ---------------------------------------------------------------------------

function Get-WorkbenchDataRoot {
    <#
    .SYNOPSIS
        Resolves the workbench data root and creates it.

    .DESCRIPTION
        APP_PACKAGER_WORKBENCH_ROOT wins so a child packager process inherits
        the same root as the GUI or CLI that launched it. The
        WorkbenchDataRoot preference is next, then
        %LOCALAPPDATA%\AppPackagerData\Workbench.
    #>
    param(
        [AllowNull()]$Preferences,
        [switch]$NoCreate
    )

    $root = [string]$env:APP_PACKAGER_WORKBENCH_ROOT
    if ([string]::IsNullOrWhiteSpace($root) -and $Preferences) {
        $member = Get-WorkbenchMember -InputObject $Preferences -Path 'WorkbenchDataRoot'
        if ($member.Found) { $root = [string]$member.Value }
    }
    if ([string]::IsNullOrWhiteSpace($root)) {
        $localAppData = [string]$env:LOCALAPPDATA
        if ([string]::IsNullOrWhiteSpace($localAppData)) {
            $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
        }
        $root = Join-Path (Join-Path $localAppData 'AppPackagerData') 'Workbench'
    }
    if (-not $NoCreate) { [void](New-WorkbenchFolder -Path $root) }
    return $root
}

function Get-SharedDownloadCacheRoot {
    <#
    .SYNOPSIS
        Resolves the shared installer download cache for any download root.

    .DESCRIPTION
        Per-profile staging isolation passes <DownloadRoot>\profiles\<ProfileId>
        as the download root. The cache must not fork with it, so a profile
        subroot resolves back to <DownloadRoot>\_cache.
    #>
    param(
        [Parameter(Mandatory)][string]$DownloadRoot,
        [switch]$NoCreate
    )

    $full = [System.IO.Path]::GetFullPath($DownloadRoot).TrimEnd('\')
    $parent = Split-Path -Path $full -Parent
    if ($parent -and (Split-Path -Path $parent -Leaf) -ieq 'profiles') {
        $full = (Split-Path -Path $parent -Parent).TrimEnd('\')
    }
    $cache = Join-Path $full '_cache'
    if (-not $NoCreate) { [void](New-WorkbenchFolder -Path $cache) }
    return $cache
}

function New-ApplicationId {
    <#
    .SYNOPSIS
        Builds a stable ApplicationId for a catalog, custom or BYO source.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Catalog', 'Custom', 'Byo')][string]$Kind,
        [string]$ScriptPath,
        [string]$Name
    )

    switch ($Kind) {
        'Byo' { return ('byo:{0}' -f [guid]::NewGuid().ToString('N')) }
        default {
            $base = $Name
            if ([string]::IsNullOrWhiteSpace($base) -and -not [string]::IsNullOrWhiteSpace($ScriptPath)) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
            }
            if ([string]::IsNullOrWhiteSpace($base)) { throw 'An application id needs a script path or a name.' }
            $base = [System.IO.Path]::GetFileNameWithoutExtension($base)
            return ('{0}:{1}' -f $Kind.ToLowerInvariant(), $base)
        }
    }
}

function ConvertTo-ApplicationKey {
    <#
    .SYNOPSIS
        Converts an ApplicationId into its folder-safe storage key.
    #>
    param([Parameter(Mandatory)][string]$ApplicationId)

    if ($ApplicationId -notmatch '^(catalog|custom|byo):[A-Za-z0-9._-]+$') {
        throw "Invalid ApplicationId '$ApplicationId'; expected catalog:<name>, custom:<name> or byo:<id>."
    }
    return ($ApplicationId -replace ':', '_')
}

function ConvertFrom-ApplicationKey {
    param([Parameter(Mandatory)][string]$ApplicationKey)
    $index = $ApplicationKey.IndexOf('_')
    if ($index -lt 1) { throw "Invalid ApplicationKey '$ApplicationKey'." }
    return ($ApplicationKey.Substring(0, $index) + ':' + $ApplicationKey.Substring($index + 1))
}

function Get-WorkbenchApplicationFolder {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$DataRoot,
        [switch]$NoCreate
    )
    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    $folder = Join-Path (Join-Path $DataRoot 'applications') (ConvertTo-ApplicationKey -ApplicationId $ApplicationId)
    if (-not $NoCreate) { [void](New-WorkbenchFolder -Path $folder) }
    return $folder
}

# ---------------------------------------------------------------------------
# Application discovery and definitions
# ---------------------------------------------------------------------------

function Get-WorkbenchApplications {
    <#
    .SYNOPSIS
        Lists every application the workbench can edit.

    .DESCRIPTION
        Catalog packagers discovered under PackagersRoot, user-authored
        packagers discovered under CustomScriptRoot, and every persisted BYO
        application under the data root. A persisted definition supplies the
        display metadata; a discovered script without one is still listed.
    #>
    param(
        [string]$PackagersRoot,
        [string]$CustomScriptRoot,
        [string]$DataRoot
    )

    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    if ([string]::IsNullOrWhiteSpace($PackagersRoot)) { $PackagersRoot = $PSScriptRoot }
    if ([string]::IsNullOrWhiteSpace($CustomScriptRoot)) { $CustomScriptRoot = Join-Path $DataRoot 'scripts' }

    $results = New-Object System.Collections.ArrayList
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    $discover = {
        param([string]$Root, [string]$Kind)
        if ([string]::IsNullOrWhiteSpace($Root) -or -not (Test-Path -LiteralPath $Root)) { return }
        $files = @(Get-ChildItem -LiteralPath $Root -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^package-.*\.(?:ps1|notps1)$' } | Sort-Object Name)
        foreach ($file in $files) {
            $id = New-ApplicationId -Kind $Kind -ScriptPath $file.FullName
            if (-not $seen.Add($id)) { continue }
            $definition = Get-ApplicationDefinition -ApplicationId $id -DataRoot $DataRoot
            [void]$results.Add([pscustomobject]@{
                ApplicationId   = $id
                ApplicationKey  = ConvertTo-ApplicationKey -ApplicationId $id
                Kind            = $Kind
                DisplayName     = $(if ($definition.DisplayName) { [string]$definition.DisplayName } else { [System.IO.Path]::GetFileNameWithoutExtension($file.Name) })
                Publisher       = [string]$definition.Publisher
                ScriptPath      = $file.FullName
                ScriptName      = $file.Name
                Runnable        = ($file.Extension -ine '.notps1')
                ActiveProfileId = [string]$definition.ActiveProfileId
                UpdatePolicy    = [string]$definition.UpdatePolicy
                Persisted       = [bool]$definition.Persisted
            })
        }
    }

    & $discover $PackagersRoot 'Catalog'
    & $discover $CustomScriptRoot 'Custom'

    $applicationsRoot = Join-Path $DataRoot 'applications'
    if (Test-Path -LiteralPath $applicationsRoot) {
        foreach ($folder in @(Get-ChildItem -LiteralPath $applicationsRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
            if ($folder.Name -notmatch '^byo_') { continue }
            $id = ConvertFrom-ApplicationKey -ApplicationKey $folder.Name
            if (-not $seen.Add($id)) { continue }
            $definition = Get-ApplicationDefinition -ApplicationId $id -DataRoot $DataRoot
            [void]$results.Add([pscustomobject]@{
                ApplicationId   = $id
                ApplicationKey  = $folder.Name
                Kind            = 'Byo'
                DisplayName     = [string]$definition.DisplayName
                Publisher       = [string]$definition.Publisher
                ScriptPath      = $null
                ScriptName      = $null
                Runnable        = $true
                ActiveProfileId = [string]$definition.ActiveProfileId
                UpdatePolicy    = $(if ($definition.UpdatePolicy) { [string]$definition.UpdatePolicy } else { 'Manual' })
                Persisted       = $true
            })
        }
    }

    return @($results)
}

function Get-ApplicationDefinition {
    <#
    .SYNOPSIS
        Reads application.json, or returns the unsaved default definition.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$DataRoot
    )

    $folder = Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot -NoCreate
    $stored = Read-WorkbenchJsonFile -Path (Join-Path $folder 'application.json')
    if ($stored) {
        $definition = ConvertTo-WorkbenchHashtable -InputObject $stored
        $definition['Persisted'] = $true
        if (-not $definition['ActiveProfileId']) { $definition['ActiveProfileId'] = 'default' }
        # Callers run under StrictMode: every field of the unsaved default
        # shape must also exist on a stored definition.
        if (-not $definition.Contains('Sources')) { $definition['Sources'] = @() }
        foreach ($field in @('Origin', 'DisplayName', 'Publisher', 'Description', 'ProviderId', 'UpdatePolicy')) {
            if (-not $definition.Contains($field)) { $definition[$field] = $null }
        }
        return [pscustomobject]$definition
    }
    return [pscustomobject]@{
        SchemaVersion   = $script:WorkbenchProfileSchemaVersion
        ApplicationId   = $ApplicationId
        Origin          = ($ApplicationId -split ':')[0]
        DisplayName     = $null
        Publisher       = $null
        Description     = $null
        ProviderId      = $null
        UpdatePolicy    = $(if ($ApplicationId -like 'byo:*') { 'Manual' } else { 'Discovered' })
        ActiveProfileId = 'default'
        Sources         = @()
        Persisted       = $false
    }
}

function Save-ApplicationDefinition {
    <#
    .SYNOPSIS
        Writes application.json atomically and returns the stored definition.
    #>
    param(
        [Parameter(Mandatory)]$Definition,
        [string]$DataRoot
    )

    $data = ConvertTo-WorkbenchHashtable -InputObject $Definition
    $applicationId = [string]$data['ApplicationId']
    if ([string]::IsNullOrWhiteSpace($applicationId)) { throw 'An application definition requires an ApplicationId.' }
    [void](ConvertTo-ApplicationKey -ApplicationId $applicationId)
    if (-not $data['SchemaVersion']) { $data['SchemaVersion'] = $script:WorkbenchProfileSchemaVersion }
    if ([int]$data['SchemaVersion'] -gt $script:WorkbenchProfileSchemaVersion) {
        throw "Application definition schema $($data['SchemaVersion']) is newer than this build supports ($script:WorkbenchProfileSchemaVersion)."
    }
    if (-not $data['ActiveProfileId']) { $data['ActiveProfileId'] = 'default' }
    [void]$data.Remove('Persisted')

    $folder = Get-WorkbenchApplicationFolder -ApplicationId $applicationId -DataRoot $DataRoot
    [void](Write-WorkbenchJsonFile -Path (Join-Path $folder 'application.json') -InputObject $data)
    return (Get-ApplicationDefinition -ApplicationId $applicationId -DataRoot $DataRoot)
}

# ---------------------------------------------------------------------------
# Profiles
# ---------------------------------------------------------------------------

function New-WorkbenchProfileObject {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$ProfileId,
        [string]$Name = 'Custom'
    )
    if ([string]::IsNullOrWhiteSpace($ProfileId)) { $ProfileId = [guid]::NewGuid().ToString('N') }
    return @{
        SchemaVersion    = $script:WorkbenchProfileSchemaVersion
        ProfileId        = $ProfileId
        ApplicationId    = $ApplicationId
        Name             = $Name
        Revision         = 0
        BaseProviderHash = $null
        PinnedVersion    = $null
        Application      = @{}
        Install          = @{}
        Uninstall        = @{}
        Detection        = @{ Mode = 'Inherit' }
        Requirements     = @{ Operations = @() }
        Variants         = @{ Split = @(); Overrides = @{} }
        InstallMode      = $null
        Timing           = @{}
        Execution        = @{}
        SourceFiles      = @()
        Tokens           = @{}
        Assets           = @{}
    }
}

function Get-WorkbenchProfileFolder {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$DataRoot,
        [switch]$NoCreate
    )
    if ($ProfileId -eq 'default') { throw 'The default profile is virtual and has no storage folder.' }
    if ($ProfileId -notmatch '^[A-Za-z0-9._-]+$') { throw "Invalid ProfileId '$ProfileId'." }
    $folder = Join-Path (Join-Path (Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot -NoCreate:$NoCreate) 'profiles') $ProfileId
    if (-not $NoCreate) { [void](New-WorkbenchFolder -Path $folder) }
    return $folder
}

function Get-Profiles {
    <#
    .SYNOPSIS
        Lists the virtual packager-default profile plus every saved profile.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$DataRoot
    )

    $definition = Get-ApplicationDefinition -ApplicationId $ApplicationId -DataRoot $DataRoot
    $active = [string]$definition.ActiveProfileId
    $results = New-Object System.Collections.ArrayList
    [void]$results.Add([pscustomobject]@{
        ProfileId     = 'default'
        ApplicationId = $ApplicationId
        Name          = 'Packager default'
        Revision      = 0
        IsDefault     = $true
        IsActive      = ($active -eq 'default' -or [string]::IsNullOrWhiteSpace($active))
        Path          = $null
    })

    $profilesRoot = Join-Path (Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot -NoCreate) 'profiles'
    if (Test-Path -LiteralPath $profilesRoot) {
        foreach ($folder in @(Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $path = Join-Path $folder.FullName 'profile.json'
            $stored = Read-WorkbenchJsonFile -Path $path
            if (-not $stored) { continue }
            [void]$results.Add([pscustomobject]@{
                ProfileId     = [string]$stored.ProfileId
                ApplicationId = $ApplicationId
                Name          = [string]$stored.Name
                Revision      = [int]$stored.Revision
                IsDefault     = $false
                IsActive      = ($active -eq [string]$stored.ProfileId)
                Path          = $path
            })
        }
    }
    return @($results)
}

function Get-Profile {
    <#
    .SYNOPSIS
        Loads one saved profile, or the virtual default profile.

    .DESCRIPTION
        The default profile is never stored and never editable; it is
        returned as an empty override set with Revision 0.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$DataRoot
    )

    if ($ProfileId -eq 'default') {
        $default = New-WorkbenchProfileObject -ApplicationId $ApplicationId -ProfileId 'default' -Name 'Packager default'
        $default['IsDefault'] = $true
        return [pscustomobject]$default
    }

    $path = Join-Path (Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot -NoCreate) 'profile.json'
    $stored = Read-WorkbenchJsonFile -Path $path
    if (-not $stored) { throw "Profile '$ProfileId' was not found for application '$ApplicationId'." }
    $data = ConvertTo-WorkbenchHashtable -InputObject $stored
    if ([int]$data['SchemaVersion'] -gt $script:WorkbenchProfileSchemaVersion) {
        throw "Profile schema $($data['SchemaVersion']) is newer than this build supports ($script:WorkbenchProfileSchemaVersion): $path"
    }
    $data['IsDefault'] = $false
    return [pscustomobject]$data
}

function Save-Profile {
    <#
    .SYNOPSIS
        Saves a profile, raising its revision by one, with a conflict check.

    .DESCRIPTION
        The incoming Revision is the revision the editor loaded. A stored
        revision that no longer matches means another writer saved in
        between, and the save is refused rather than silently overwriting.
    #>
    param(
        [Parameter(Mandatory)]$Profile,
        [string]$DataRoot,
        [switch]$SetActive
    )

    $data = ConvertTo-WorkbenchHashtable -InputObject $Profile
    [void]$data.Remove('IsDefault')
    $applicationId = [string]$data['ApplicationId']
    $profileId = [string]$data['ProfileId']
    if ([string]::IsNullOrWhiteSpace($applicationId)) { throw 'A profile requires an ApplicationId.' }
    if ($profileId -eq 'default') { throw 'The packager default profile cannot be saved; use Copy-Profile to create an editable profile.' }
    if ([string]::IsNullOrWhiteSpace($profileId)) { $profileId = [guid]::NewGuid().ToString('N'); $data['ProfileId'] = $profileId }
    if (-not $data['SchemaVersion']) { $data['SchemaVersion'] = $script:WorkbenchProfileSchemaVersion }
    if ([int]$data['SchemaVersion'] -gt $script:WorkbenchProfileSchemaVersion) {
        throw "Profile schema $($data['SchemaVersion']) is newer than this build supports ($script:WorkbenchProfileSchemaVersion)."
    }
    if ([string]::IsNullOrWhiteSpace([string]$data['Name'])) { throw 'A profile requires a Name.' }

    $folder = Get-WorkbenchProfileFolder -ApplicationId $applicationId -ProfileId $profileId -DataRoot $DataRoot
    $path = Join-Path $folder 'profile.json'
    $existing = Read-WorkbenchJsonFile -Path $path
    $incoming = [int]$data['Revision']
    if ($existing) {
        $storedRevision = [int]$existing.Revision
        if ($storedRevision -ne $incoming) {
            throw "Profile '$profileId' changed on disk (stored revision $storedRevision, this editor loaded $incoming); reload before saving."
        }
    }
    elseif ($incoming -gt 0) {
        throw "Profile '$profileId' carries revision $incoming but has no stored profile; save it as a new profile instead."
    }

    Assert-WorkbenchProfileValid -Profile $data
    $data['Revision'] = $incoming + 1
    $data['SavedAt'] = (Get-Date -Format 'o')
    [void](Write-WorkbenchJsonFile -Path $path -InputObject $data)

    if ($SetActive) {
        $definition = ConvertTo-WorkbenchHashtable -InputObject (Get-ApplicationDefinition -ApplicationId $applicationId -DataRoot $DataRoot)
        $definition['ActiveProfileId'] = $profileId
        [void](Save-ApplicationDefinition -Definition $definition -DataRoot $DataRoot)
    }
    return (Get-Profile -ApplicationId $applicationId -ProfileId $profileId -DataRoot $DataRoot)
}

function Assert-WorkbenchProfileValid {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Profile)

    $timing = $Profile['Timing']
    if ($timing -is [System.Collections.IDictionary]) {
        $estimated = $timing['EstimatedMinutes']
        $maximum = $timing['MaximumMinutes']
        foreach ($pair in @(@{ N = 'EstimatedMinutes'; V = $estimated }, @{ N = 'MaximumMinutes'; V = $maximum })) {
            if ($null -eq $pair.V) { continue }
            $value = 0
            if (-not [int]::TryParse([string]$pair.V, [ref]$value) -or $value -lt 1) {
                throw "Timing.$($pair.N) must be a positive whole number of minutes (got '$($pair.V)')."
            }
        }
        if ($null -ne $estimated -and $null -ne $maximum -and [int]$estimated -gt [int]$maximum) {
            throw "Timing.EstimatedMinutes ($estimated) cannot exceed Timing.MaximumMinutes ($maximum)."
        }
    }

    foreach ($section in @('Install', 'Uninstall')) {
        $block = $Profile[$section]
        if (-not ($block -is [System.Collections.IDictionary])) { continue }
        if ($block.Contains('Mode') -and $null -ne $block['Mode'] -and [string]$block['Mode'] -notin @('Generated', 'Extend', 'Custom')) {
            throw "$section.Mode must be Generated, Extend or Custom (got '$($block['Mode'])')."
        }
    }

    $detection = $Profile['Detection']
    if ($detection -is [System.Collections.IDictionary] -and $detection.Contains('Mode') -and $null -ne $detection['Mode']) {
        if ([string]$detection['Mode'] -notin @('Inherit', 'Custom')) {
            throw "Detection.Mode must be Inherit or Custom (got '$($detection['Mode'])')."
        }
        if ([string]$detection['Mode'] -eq 'Custom' -and -not $detection['Rule']) {
            throw 'Detection.Mode Custom requires a Detection.Rule.'
        }
    }

    $requirements = $Profile['Requirements']
    if ($requirements -is [System.Collections.IDictionary] -and $requirements['Operations']) {
        foreach ($operation in @($requirements['Operations'])) {
            if ($null -eq $operation) { continue }
            $op = [string](Get-WorkbenchMember -InputObject $operation -Path 'Op').Value
            $ruleId = [string](Get-WorkbenchMember -InputObject $operation -Path 'RuleId').Value
            if ($op -notin @('Add', 'Replace', 'Remove')) { throw "Requirement operation Op must be Add, Replace or Remove (got '$op')." }
            if ([string]::IsNullOrWhiteSpace($ruleId)) { throw 'Every requirement operation needs a RuleId.' }
            if ($op -ne 'Remove' -and -not (Get-WorkbenchMember -InputObject $operation -Path 'Rule').Value) {
                throw "Requirement operation '$op' for rule '$ruleId' needs a Rule."
            }
        }
    }

    foreach ($entry in @($Profile['SourceFiles'])) {
        if ($null -eq $entry) { continue }
        $destination = [string](Get-WorkbenchMember -InputObject $entry -Path 'Destination').Value
        if ([string]::IsNullOrWhiteSpace($destination)) { throw 'Every SourceFiles entry needs a Destination.' }
        if ([System.IO.Path]::IsPathRooted($destination)) { throw "SourceFiles destination '$destination' must be relative to the content root." }
    }
}

function Copy-Profile {
    <#
    .SYNOPSIS
        Save as: copies a profile and its managed assets under a new id.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [Parameter(Mandatory)][string]$NewName,
        [string]$DataRoot,
        [switch]$SetActive
    )

    if ([string]::IsNullOrWhiteSpace($NewName)) { throw 'Save as requires a profile name.' }
    $source = Get-Profile -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot
    $data = ConvertTo-WorkbenchHashtable -InputObject $source
    [void]$data.Remove('IsDefault')
    $newId = [guid]::NewGuid().ToString('N')
    $data['ProfileId'] = $newId
    $data['Name'] = $NewName
    $data['Revision'] = 0
    $data['CopiedFrom'] = $ProfileId

    $target = Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $newId -DataRoot $DataRoot
    if ($ProfileId -ne 'default') {
        $sourceAssets = Join-Path (Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot -NoCreate) 'assets'
        if (Test-Path -LiteralPath $sourceAssets) {
            Copy-Item -LiteralPath $sourceAssets -Destination (Join-Path $target 'assets') -Recurse -Force -ErrorAction Stop
        }
    }
    return (Save-Profile -Profile ([pscustomobject]$data) -DataRoot $DataRoot -SetActive:$SetActive)
}

function Remove-ProfileField {
    <#
    .SYNOPSIS
        Clears one profile field to inherit, or sets it to an explicit none.

    .DESCRIPTION
        Inherit removes the key so the next precedence level supplies the
        value. Explicit stores $null, which means "none" and stops
        inheritance. The two are deliberately different states.
    #>
    param(
        [Parameter(Mandatory)]$Profile,
        [Parameter(Mandatory)][string]$Field,
        [ValidateSet('Inherit', 'Explicit')][string]$Mode = 'Inherit'
    )

    $data = ConvertTo-WorkbenchHashtable -InputObject $Profile
    if ($Mode -eq 'Inherit') { [void](Remove-WorkbenchMember -InputObject $data -Path $Field) }
    else { [void](Set-WorkbenchMember -InputObject $data -Path $Field -Value $null) }
    return [pscustomobject]$data
}

# ---------------------------------------------------------------------------
# Profile assets
# ---------------------------------------------------------------------------

function Add-ProfileAsset {
    <#
    .SYNOPSIS
        Copies a file into managed profile storage keyed by content hash.

    .DESCRIPTION
        Linked assets are recorded by provenance only and are re-read and
        snapshotted at build time; a missing linked file fails the build.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [Parameter(Mandatory)][string]$Path,
        [string]$DataRoot,
        [switch]$Linked
    )

    if ($ProfileId -eq 'default') { throw 'Assets cannot be added to the packager default profile.' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Asset source not found: $Path" }
    if (Test-WorkbenchReparsePoint -Path $Path) { throw "Asset source '$Path' is a reparse point; copy the real file instead." }

    $full = [System.IO.Path]::GetFullPath($Path)
    $sha = Get-WorkbenchFileSha256 -Path $full
    $assetId = $sha.Substring(0, 16)
    $name = Split-Path -Path $full -Leaf
    $size = (Get-Item -LiteralPath $full).Length

    $storedPath = $null
    if (-not $Linked) {
        $folder = New-WorkbenchFolder -Path (Join-Path (Join-Path (Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot) 'assets') $assetId)
        $storedPath = Join-Path $folder $name
        Copy-Item -LiteralPath $full -Destination $storedPath -Force -ErrorAction Stop
    }

    return [pscustomobject]@{
        AssetId    = $assetId
        FileName   = $name
        Sha256     = $sha
        Size       = $size
        Provenance = $full
        Linked     = [bool]$Linked
        Path       = $(if ($Linked) { $full } else { $storedPath })
        AddedAt    = (Get-Date -Format 'o')
    }
}

function Remove-ProfileAsset {
    <#
    .SYNOPSIS
        Deletes one managed asset folder from a profile.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [Parameter(Mandatory)][string]$AssetId,
        [string]$DataRoot
    )

    if ($AssetId -notmatch '^[0-9a-fA-F]{16}$') { throw "Invalid AssetId '$AssetId'." }
    $folder = Join-Path (Join-Path (Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot -NoCreate) 'assets') $AssetId
    if (-not (Test-Path -LiteralPath $folder)) { return $false }
    Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
    return $true
}

function Resolve-ProfileAssetPath {
    <#
    .SYNOPSIS
        Resolves an asset reference to a readable file for this build.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [Parameter(Mandatory)]$Asset,
        [string]$DataRoot
    )

    if ($Asset -is [string]) { $Asset = @{ Asset = $Asset } }
    $assetId = [string](Get-WorkbenchMember -InputObject $Asset -Path 'Asset').Value
    if ([string]::IsNullOrWhiteSpace($assetId)) { $assetId = [string](Get-WorkbenchMember -InputObject $Asset -Path 'AssetId').Value }
    $linked = [bool](Get-WorkbenchMember -InputObject $Asset -Path 'Linked').Value
    $provenance = [string](Get-WorkbenchMember -InputObject $Asset -Path 'Provenance').Value

    if ($linked) {
        if ([string]::IsNullOrWhiteSpace($provenance) -or -not (Test-Path -LiteralPath $provenance -PathType Leaf)) {
            throw "Linked source file '$provenance' is missing; the build cannot snapshot it."
        }
        return $provenance
    }
    if ([string]::IsNullOrWhiteSpace($assetId)) { throw 'Asset reference carries neither an asset id nor a linked path.' }
    $folder = Join-Path (Join-Path (Get-WorkbenchProfileFolder -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot -NoCreate) 'assets') $assetId
    $file = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($file.Count -eq 0) { throw "Managed asset '$assetId' is missing from profile '$ProfileId'." }
    return $file[0].FullName
}

# ---------------------------------------------------------------------------
# Effective settings
# ---------------------------------------------------------------------------

function Get-WorkbenchFieldMap {
    # Field -> where each precedence level stores it. NeedsStaging marks the
    # fields only a completed Stage can resolve.
    return @(
        [pscustomobject]@{ Field = 'DisplayName';      Global = $null;                  Packager = 'AppName';              Profile = 'Application.DisplayName'; NeedsStaging = $true }
        [pscustomobject]@{ Field = 'Publisher';        Global = $null;                  Packager = 'Publisher';            Profile = 'Application.Publisher';   NeedsStaging = $true }
        [pscustomobject]@{ Field = 'Description';      Global = 'Description';          Packager = 'Description';          Profile = 'Application.Description'; NeedsStaging = $false }
        [pscustomobject]@{ Field = 'TitleMode';        Global = 'TitleMode';            Packager = $null;                  Profile = 'Application.TitleMode';   NeedsStaging = $false }
        [pscustomobject]@{ Field = 'Icon';             Global = $null;                  Packager = 'Icon';                 Profile = 'Application.Icon';        NeedsStaging = $true }
        [pscustomobject]@{ Field = 'PinnedVersion';    Global = $null;                  Packager = 'SoftwareVersion';      Profile = 'PinnedVersion';           NeedsStaging = $true }
        [pscustomobject]@{ Field = 'InstallMode';      Global = $null;                  Packager = 'InstallMode';          Profile = 'InstallMode';             NeedsStaging = $false }
        [pscustomobject]@{ Field = 'InstallCommand';   Global = $null;                  Packager = 'InstallCommandLine';   Profile = 'Install.Command';         NeedsStaging = $true }
        [pscustomobject]@{ Field = 'UninstallCommand'; Global = $null;                  Packager = 'UninstallCommandLine'; Profile = 'Uninstall.Command';       NeedsStaging = $true }
        [pscustomobject]@{ Field = 'InstallScriptMode';Global = $null;                  Packager = $null;                  Profile = 'Install.Mode';            NeedsStaging = $false }
        [pscustomobject]@{ Field = 'UninstallScriptMode'; Global = $null;               Packager = $null;                  Profile = 'Uninstall.Mode';          NeedsStaging = $false }
        [pscustomobject]@{ Field = 'EstimatedMinutes'; Global = 'EstimatedRuntimeMins'; Packager = $null;                  Profile = 'Timing.EstimatedMinutes'; NeedsStaging = $false }
        [pscustomobject]@{ Field = 'MaximumMinutes';   Global = 'MaximumRuntimeMins';   Packager = $null;                  Profile = 'Timing.MaximumMinutes';   NeedsStaging = $false }
        [pscustomobject]@{ Field = 'Context';          Global = $null;                  Packager = 'InstallationBehaviorType'; Profile = 'Execution.Context';   NeedsStaging = $false }
        [pscustomobject]@{ Field = 'LogonRequirement'; Global = $null;                  Packager = 'LogonRequirementType'; Profile = 'Execution.LogonRequirement'; NeedsStaging = $false }
        [pscustomobject]@{ Field = 'UserInteraction';  Global = $null;                  Packager = 'RequireUserInteraction'; Profile = 'Execution.UserInteraction'; NeedsStaging = $false }
        [pscustomobject]@{ Field = 'ScriptHost';       Global = $null;                  Packager = $null;                  Profile = 'Execution.ScriptHost';    NeedsStaging = $false }
        [pscustomobject]@{ Field = 'Detection';        Global = $null;                  Packager = 'Detection';            Profile = 'Detection.Rule';          NeedsStaging = $true }
    )
}

function Resolve-EffectiveSettings {
    <#
    .SYNOPSIS
        Merges every precedence level into one field-by-field result.

    .DESCRIPTION
        Precedence, highest last: Global, Packager, Profile, Variant, Target,
        Run. Each field reports its value and the level that supplied it.
        A field nothing supplies reports NeedsStaging when only a completed
        Stage can resolve it, otherwise Global with a null value.

        A profile field holding an explicit $null means "none" and still
        wins; an absent key inherits.
    #>
    param(
        [AllowNull()]$GlobalDefaults,
        [AllowNull()]$BaseManifest,
        [AllowNull()]$Profile,
        [string]$Variant,
        [AllowNull()]$TargetOverrides,
        [AllowNull()]$RunOverrides
    )

    $variantOverride = $null
    if (-not [string]::IsNullOrWhiteSpace($Variant) -and $Profile) {
        $member = Get-WorkbenchMember -InputObject $Profile -Path ('Variants.Overrides.' + $Variant)
        if ($member.Found) { $variantOverride = $member.Value }
    }

    $result = [ordered]@{}
    foreach ($entry in (Get-WorkbenchFieldMap)) {
        $value = $null
        $source = $null

        if ($GlobalDefaults -and $entry.Global) {
            $member = Get-WorkbenchMember -InputObject $GlobalDefaults -Path $entry.Global
            if ($member.Found -and $null -ne $member.Value) { $value = $member.Value; $source = 'Global' }
        }
        if ($BaseManifest -and $entry.Packager) {
            $member = Get-WorkbenchMember -InputObject $BaseManifest -Path $entry.Packager
            if ($member.Found -and $null -ne $member.Value -and '' -ne [string]$member.Value) { $value = $member.Value; $source = 'Packager' }
        }
        if ($Profile -and $entry.Profile) {
            $member = Get-WorkbenchMember -InputObject $Profile -Path $entry.Profile
            if ($member.Found) { $value = $member.Value; $source = 'Profile' }
        }
        if ($variantOverride) {
            $member = Get-WorkbenchMember -InputObject $variantOverride -Path $entry.Field
            if ($member.Found) { $value = $member.Value; $source = 'Variant' }
        }
        if ($TargetOverrides) {
            $member = Get-WorkbenchMember -InputObject $TargetOverrides -Path $entry.Field
            if ($member.Found) { $value = $member.Value; $source = 'Target' }
        }
        if ($RunOverrides) {
            # Found alone, like the variant and target levels: an explicit
            # null run override means "none" and must outrank the profile.
            $member = Get-WorkbenchMember -InputObject $RunOverrides -Path $entry.Field
            if ($member.Found) { $value = $member.Value; $source = 'Run' }
        }

        if (-not $source) {
            $source = $(if ($entry.NeedsStaging) { 'NeedsStaging' } else { 'Global' })
        }
        $result[$entry.Field] = [pscustomobject]@{ Field = $entry.Field; Value = $value; Source = $source }
    }
    return $result
}

# ---------------------------------------------------------------------------
# Legacy preference migration
# ---------------------------------------------------------------------------

function Invoke-LegacyPreferenceMigration {
    <#
    .SYNOPSIS
        Migrates legacy per-app preference maps into workbench profiles.

    .DESCRIPTION
        Reads DeploymentConditions.Apps, CommandOverrides.Apps, and the
        InstallMode/TitleMode fields inside them, keyed by packager base name,
        and writes one profile named Migrated per application, which becomes
        active. The legacy keys stay in place so the environment-variable
        bridge keeps working until an active profile supersedes them.

        The legacy key is a packager base name, so the application id it maps
        to depends on where that script was discovered: a catalog script
        becomes catalog:<name> and a user-authored one custom:<name>.

        Provenance is the stored MigratedFrom marker, never the profile's
        display name, so renaming the profile does not create a second one.
        The migrated profile is activated only while the application still
        sits on the virtual default profile; an operator's own active choice
        is left alone.

        Idempotent: an application whose migrated profile already records the
        same legacy content is skipped.
    #>
    param(
        [Parameter(Mandatory)]$Preferences,
        [string]$DataRoot,
        [string]$PackagersRoot,
        [string]$CustomScriptRoot
    )

    $migrated = New-Object System.Collections.ArrayList
    $skipped = New-Object System.Collections.Generic.List[string]

    $originOf = @{}
    foreach ($application in (Get-WorkbenchApplications -PackagersRoot $PackagersRoot -CustomScriptRoot $CustomScriptRoot -DataRoot $DataRoot)) {
        if ([string]::IsNullOrWhiteSpace([string]$application.ScriptName)) { continue }
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$application.ScriptName)
        # Catalog wins a name collision: the legacy panel only ever listed
        # packagers discovered under the catalog root.
        if ($originOf.ContainsKey($base) -and $originOf[$base] -like 'catalog:*') { continue }
        $originOf[$base] = [string]$application.ApplicationId
    }

    $conditionApps = (Get-WorkbenchMember -InputObject $Preferences -Path 'DeploymentConditions.Apps').Value
    $commandApps = (Get-WorkbenchMember -InputObject $Preferences -Path 'CommandOverrides.Apps').Value

    $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($source in @($conditionApps, $commandApps)) {
        if (-not $source) { continue }
        foreach ($property in $source.PSObject.Properties) { [void]$keys.Add($property.Name) }
    }

    foreach ($key in @($keys | Sort-Object)) {
        $condition = $null
        $command = $null
        if ($conditionApps -and $conditionApps.PSObject.Properties[$key]) { $condition = $conditionApps.PSObject.Properties[$key].Value }
        if ($commandApps -and $commandApps.PSObject.Properties[$key]) { $command = $commandApps.PSObject.Properties[$key].Value }

        $legacy = @{
            Conditions = ConvertTo-WorkbenchHashtable -InputObject $condition
            Commands   = ConvertTo-WorkbenchHashtable -InputObject $command
        }
        $legacyDigest = Get-WorkbenchTextSha256 -Text (ConvertTo-WorkbenchJson -InputObject $legacy)

        $applicationId = $(if ($originOf.ContainsKey($key)) { $originOf[$key] } else { New-ApplicationId -Kind Catalog -Name $key })

        $existing = $null
        foreach ($candidate in (Get-Profiles -ApplicationId $applicationId -DataRoot $DataRoot)) {
            if ($candidate.IsDefault) { continue }
            $stored = Get-Profile -ApplicationId $applicationId -ProfileId $candidate.ProfileId -DataRoot $DataRoot
            if ([string]$stored.MigratedFrom -eq $key) { $existing = $stored; break }
        }
        if ($existing -and [string]$existing.MigratedDigest -eq $legacyDigest) { $skipped.Add($applicationId); continue }

        $profile = if ($existing) {
            ConvertTo-WorkbenchHashtable -InputObject $existing
        }
        else {
            New-WorkbenchProfileObject -ApplicationId $applicationId -Name 'Migrated'
        }
        [void]$profile.Remove('IsDefault')
        $profile['MigratedFrom'] = $key
        $profile['MigratedDigest'] = $legacyDigest

        if ($condition) {
            $rules = @()
            $architecture = [string](Get-WorkbenchMember -InputObject $condition -Path 'Architecture').Value
            if ($architecture -in @('x64', 'ARM64')) {
                $rules += @{ Op = 'Add'; RuleId = 'cpu-arch'; Rule = @{ ConditionId = 'cpu-arch'; Value = $architecture }; AppliesTo = @('*') }
            }
            $languages = @(@((Get-WorkbenchMember -InputObject $condition -Path 'Languages').Value) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
            if ($languages.Count -gt 0) {
                $rules += @{ Op = 'Add'; RuleId = 'os-language'; Rule = @{ ConditionId = 'os-language'; Cultures = $languages }; AppliesTo = @('*') }
            }
            switch ([string](Get-WorkbenchMember -InputObject $condition -Path 'Network').Value) {
                'VpnOnly' { $rules += @{ Op = 'Add'; RuleId = 'vpn-connected'; Rule = @{ ConditionId = 'vpn-connected'; Value = $true }; AppliesTo = @('*') } }
                'OnSiteOnly' { $rules += @{ Op = 'Add'; RuleId = 'vpn-connected'; Rule = @{ ConditionId = 'vpn-connected'; Value = $false }; AppliesTo = @('*') } }
            }
            if ($rules.Count -gt 0) { $profile['Requirements'] = @{ Operations = $rules } }

            $split = [string](Get-WorkbenchMember -InputObject $condition -Path 'Split').Value
            if ($split -in @('Architecture', 'Language', 'Network')) {
                $variants = @{ SchemaVersion = 1; Split = $split }
                if ($split -eq 'Language' -and $languages.Count -gt 0) { $variants['Languages'] = $languages }
                $profile['Variants'] = @{ Split = $variants; Overrides = @{} }
            }

            $installMode = [string](Get-WorkbenchMember -InputObject $condition -Path 'InstallMode').Value
            if ($installMode -in @('CurrentUser', 'AllUsers')) { $profile['InstallMode'] = $installMode }

            $titleMode = [string](Get-WorkbenchMember -InputObject $condition -Path 'TitleMode').Value
            if ($titleMode -in @('IncludeVersion', 'NoVersion')) {
                if (-not ($profile['Application'] -is [System.Collections.IDictionary])) { $profile['Application'] = @{} }
                $profile['Application']['TitleMode'] = $titleMode
            }
        }

        if ($command) {
            $install = ([string](Get-WorkbenchMember -InputObject $command -Path 'Install').Value).Trim()
            $uninstall = ([string](Get-WorkbenchMember -InputObject $command -Path 'Uninstall').Value).Trim()
            if ($install) {
                if (-not ($profile['Install'] -is [System.Collections.IDictionary])) { $profile['Install'] = @{} }
                $profile['Install']['Command'] = $install
            }
            if ($uninstall) {
                if (-not ($profile['Uninstall'] -is [System.Collections.IDictionary])) { $profile['Uninstall'] = @{} }
                $profile['Uninstall']['Command'] = $uninstall
            }
        }

        $active = [string](Get-ApplicationDefinition -ApplicationId $applicationId -DataRoot $DataRoot).ActiveProfileId
        $activate = ([string]::IsNullOrWhiteSpace($active) -or $active -eq 'default' -or $active -eq [string]$profile['ProfileId'])
        $saved = Save-Profile -Profile ([pscustomobject]$profile) -DataRoot $DataRoot -SetActive:$activate
        [void]$migrated.Add($saved)
    }

    return [pscustomobject]@{
        Migrated      = @($migrated)
        MigratedCount = @($migrated).Count
        Skipped       = @($skipped)
        SkippedCount  = @($skipped).Count
    }
}

# ---------------------------------------------------------------------------
# Run snapshots
# ---------------------------------------------------------------------------

function New-BuildId {
    return ('{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)))
}

function New-RunSnapshot {
    <#
    .SYNOPSIS
        Writes the immutable per-run snapshot the packager child consumes.

    .DESCRIPTION
        The snapshot freezes the profile, the resolved asset paths, the run
        overrides, the signing policy and the target for one build. Run
        overrides live only in the snapshot and the build record; they are
        never written back into the profile.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$ProfileId = 'default',
        [ValidateSet('ContentOnly', 'MECM', 'MECMAndIntune', 'IntuneOnly')][string]$Target = 'MECM',
        [AllowNull()]$RunOverrides,
        [AllowNull()]$SigningPolicy,
        [string]$PackagerScriptPath,
        [string]$DownloadRoot,
        [string]$DataRoot,
        [string]$BuildId
    )

    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    if ([string]::IsNullOrWhiteSpace($ProfileId)) { $ProfileId = 'default' }
    if ([string]::IsNullOrWhiteSpace($BuildId)) { $BuildId = New-BuildId }

    $profile = Get-Profile -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot
    $overrides = ConvertTo-WorkbenchHashtable -InputObject $RunOverrides
    if ($null -eq $overrides) { $overrides = @{} }
    Assert-WorkbenchRunOverrides -RunOverrides $overrides
    [void](Assert-WorkbenchSigningCertificate -Policy $SigningPolicy)

    $assets = @{}
    if ($ProfileId -ne 'default') {
        foreach ($reference in (Get-WorkbenchProfileAssetReferences -Profile $profile)) {
            $assets[$reference] = Resolve-ProfileAssetPath -ApplicationId $ApplicationId -ProfileId $ProfileId -Asset @{ Asset = $reference } -DataRoot $DataRoot
        }
        foreach ($entry in @($profile.SourceFiles)) {
            if ($null -eq $entry) { continue }
            $linked = [bool](Get-WorkbenchMember -InputObject $entry -Path 'Linked').Value
            if (-not $linked) { continue }
            $provenance = [string](Get-WorkbenchMember -InputObject $entry -Path 'Provenance').Value
            $assets['linked:' + $provenance] = Resolve-ProfileAssetPath -ApplicationId $ApplicationId -ProfileId $ProfileId -Asset $entry -DataRoot $DataRoot
        }
    }

    $snapshot = @{
        SchemaVersion      = $script:WorkbenchProfileSchemaVersion
        BuildId            = $BuildId
        CreatedAt          = (Get-Date -Format 'o')
        ApplicationId      = $ApplicationId
        ProfileId          = $ProfileId
        ProfileRevision    = [int]$profile.Revision
        Profile            = (ConvertTo-WorkbenchHashtable -InputObject $profile)
        AssetPaths         = $assets
        RunOverrides       = $overrides
        SigningPolicy      = (ConvertTo-WorkbenchHashtable -InputObject $SigningPolicy)
        Target             = $Target
        DataRoot           = $DataRoot
        DownloadRoot       = $DownloadRoot
        PackagerScriptPath = $PackagerScriptPath
    }

    $path = Join-Path (New-WorkbenchFolder -Path (Join-Path $DataRoot 'runs')) ("$BuildId.json")
    [void](Write-WorkbenchJsonFile -Path $path -InputObject $snapshot)
    $snapshot['Path'] = $path
    return [pscustomobject]$snapshot
}

function Assert-WorkbenchSigningCertificate {
    <#
    .SYNOPSIS
        Resolves the signing certificate before a build writes anything.

    .DESCRIPTION
        The in-hook category signing is the second line of defence; by the
        time it runs the wrappers, payload and icon are already on disk. Any
        Sign or Require switch therefore resolves the certificate here, while
        the run snapshot is being created and before the packager child is
        launched, so an unusable certificate fails the run with
        SigningCertificateUnavailable and no partial content.
    #>
    param([AllowNull()]$Policy)

    if (-not (Test-WorkbenchSigningRequested -Policy $Policy)) { return $null }
    $resolve = Get-Command -Name Resolve-SigningCertificate -ErrorAction SilentlyContinue
    if (-not $resolve) {
        throw 'Script signing is requested but AppPackagerSigning is not loaded; no build may ship unsigned under an enabled signing policy.'
    }
    return (& $resolve -Policy $Policy)
}

function Assert-WorkbenchRunOverrides {
    param([AllowNull()]$RunOverrides)
    if (-not $RunOverrides) { return }
    $estimated = (Get-WorkbenchMember -InputObject $RunOverrides -Path 'EstimatedMinutes').Value
    $maximum = (Get-WorkbenchMember -InputObject $RunOverrides -Path 'MaximumMinutes').Value
    foreach ($pair in @(@{ N = 'EstimatedMinutes'; V = $estimated }, @{ N = 'MaximumMinutes'; V = $maximum })) {
        if ($null -eq $pair.V) { continue }
        $value = 0
        if (-not [int]::TryParse([string]$pair.V, [ref]$value) -or $value -lt 1) {
            throw "Run override $($pair.N) must be a positive whole number of minutes (got '$($pair.V)')."
        }
    }
    if ($null -ne $estimated -and $null -ne $maximum -and [int]$estimated -gt [int]$maximum) {
        throw "Run override EstimatedMinutes ($estimated) cannot exceed MaximumMinutes ($maximum)."
    }
}

function Get-WorkbenchProfileAssetReferences {
    param([Parameter(Mandatory)]$Profile)
    $references = New-Object System.Collections.Generic.List[string]
    foreach ($path in @('Application.Icon.Asset', 'Install.Script', 'Install.Before', 'Install.After',
            'Uninstall.Script', 'Uninstall.Before', 'Uninstall.After')) {
        $member = Get-WorkbenchMember -InputObject $Profile -Path $path
        if ($member.Found -and -not [string]::IsNullOrWhiteSpace([string]$member.Value)) { $references.Add([string]$member.Value) }
    }
    foreach ($entry in @($Profile.SourceFiles)) {
        if ($null -eq $entry) { continue }
        if ([bool](Get-WorkbenchMember -InputObject $entry -Path 'Linked').Value) { continue }
        $assetId = [string](Get-WorkbenchMember -InputObject $entry -Path 'Asset').Value
        if (-not [string]::IsNullOrWhiteSpace($assetId)) { $references.Add($assetId) }
    }
    return @($references | Select-Object -Unique)
}

function Get-RunSnapshot {
    <#
    .SYNOPSIS
        Reads the run snapshot named by APP_PACKAGER_RUN_SNAPSHOT, or $null.
    #>
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { $Path = [string]$env:APP_PACKAGER_RUN_SNAPSHOT }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Run snapshot not found: $Path"
    }
    $snapshot = Read-WorkbenchJsonFile -Path $Path
    if (-not $snapshot) { throw "Run snapshot is empty: $Path" }
    if ([int]$snapshot.SchemaVersion -gt $script:WorkbenchProfileSchemaVersion) {
        throw "Run snapshot schema $($snapshot.SchemaVersion) is newer than this build supports ($script:WorkbenchProfileSchemaVersion): $Path"
    }
    $data = ConvertTo-WorkbenchHashtable -InputObject $snapshot
    $data['Path'] = $Path
    return [pscustomobject]$data
}

# ---------------------------------------------------------------------------
# Token binding
# ---------------------------------------------------------------------------

function Resolve-WorkbenchTokens {
    <#
    .SYNOPSIS
        Binds the four supported build tokens in a command or script body.

    .DESCRIPTION
        Supported: {{InstallerFile}}, {{Version}}, {{ProductCode}},
        {{ContentRoot}}. ContentRoot binds to the caller's own root
        expression, because a BAT command line and a PowerShell script name
        the deployed content root differently. Any token left after binding
        throws: a literal brace pair reaching a deployed command is a defect,
        never a default.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory)]$ManifestData,
        [ValidateSet('Command', 'Script')][string]$Context = 'Command'
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $values = @{
        InstallerFile = [string](Get-WorkbenchMember -InputObject $ManifestData -Path 'InstallerFile').Value
        Version       = [string](Get-WorkbenchMember -InputObject $ManifestData -Path 'SoftwareVersion').Value
        ProductCode   = [string](Get-WorkbenchMember -InputObject $ManifestData -Path 'ProductCode').Value
        ContentRoot   = $(if ($Context -eq 'Script') { '$PSScriptRoot' } else { '%~dp0' })
    }

    $result = [regex]::Replace($Text, '\{\{\s*([A-Za-z]+)\s*\}\}', {
            param($match)
            $name = $match.Groups[1].Value
            if ($script:WorkbenchTokenNames -notcontains $name) {
                throw "Unknown build token '{{$name}}'; supported tokens are $($script:WorkbenchTokenNames -join ', ')."
            }
            $value = $values[$name]
            if ([string]::IsNullOrWhiteSpace($value)) {
                throw "Build token '{{$name}}' has no value at build time; the stage did not resolve it."
            }
            return $value
        })

    if ($result -match '\{\{') {
        throw "Unresolved build token in '$Text'."
    }
    return $result
}

# ---------------------------------------------------------------------------
# Stage finalization
# ---------------------------------------------------------------------------

function Set-WorkbenchSchema4Defaults {
    param([Parameter(Mandatory)][hashtable]$ManifestData)

    if (-not $ManifestData.ContainsKey('BuildId')) { $ManifestData['BuildId'] = '' }
    if (-not $ManifestData.ContainsKey('ApplicationId')) { $ManifestData['ApplicationId'] = '' }
    if (-not $ManifestData.ContainsKey('ProfileId')) { $ManifestData['ProfileId'] = 'default' }
    if (-not $ManifestData.ContainsKey('ProfileRevision')) { $ManifestData['ProfileRevision'] = 0 }
    if (-not $ManifestData.ContainsKey('Timing')) { $ManifestData['Timing'] = @{ EstimatedMinutes = $null; MaximumMinutes = $null } }
    if (-not $ManifestData.ContainsKey('Execution')) {
        $ManifestData['Execution'] = @{ Context = $null; LogonRequirement = $null; UserInteraction = $null; ScriptHost = $null }
    }
    if (-not $ManifestData.ContainsKey('DetectionSource')) { $ManifestData['DetectionSource'] = 'Default' }
    if (-not $ManifestData.ContainsKey('CustomAssets')) { $ManifestData['CustomAssets'] = @() }
    if (-not $ManifestData.ContainsKey('RunOverrides')) { $ManifestData['RunOverrides'] = @{} }
    if ([string]::IsNullOrWhiteSpace([string]$ManifestData['InstallCommandLine'])) { $ManifestData['InstallCommandLine'] = 'install.bat' }
    if ([string]::IsNullOrWhiteSpace([string]$ManifestData['UninstallCommandLine'])) { $ManifestData['UninstallCommandLine'] = 'uninstall.bat' }
    return $ManifestData
}

function Get-WorkbenchSigningPolicy {
    param([AllowNull()]$Snapshot)

    if (Get-Command -Name Get-SigningPolicy -ErrorAction SilentlyContinue) {
        if ($Snapshot -and $Snapshot.SigningPolicy) { return (Get-SigningPolicy -Policy $Snapshot.SigningPolicy) }
        return (Get-SigningPolicy)
    }
    if ($Snapshot -and $Snapshot.SigningPolicy) { return $Snapshot.SigningPolicy }
    $envJson = [string]$env:APP_PACKAGER_SIGNING
    if (-not [string]::IsNullOrWhiteSpace($envJson)) {
        try { return ($envJson | ConvertFrom-Json) }
        catch { throw "APP_PACKAGER_SIGNING is not valid JSON: $($_.Exception.Message)" }
    }
    return $null
}

function Test-WorkbenchSigningRequested {
    param([AllowNull()]$Policy)
    if (-not $Policy) { return $false }
    foreach ($name in @('SignDetection', 'SignRequirements', 'SignDeployment', 'RequireDetection', 'RequireRequirements', 'RequireDeployment')) {
        $member = Get-WorkbenchMember -InputObject $Policy -Path $name
        if ($member.Found -and [bool]$member.Value) { return $true }
    }
    return $false
}

function Invoke-WorkbenchSigning {
    <#
    .SYNOPSIS
        Routes the finalization signing step to the signing module.

    .DESCRIPTION
        The signing functions are resolved by name through Get-Command so the
        module stays an optional dependency in the unit-test host. A policy
        with any switch on and no signing module present is a configuration
        error, not a silent unsigned build.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [AllowNull()]$Policy
    )

    $requested = Test-WorkbenchSigningRequested -Policy $Policy
    $categorySigning = Get-Command -Name Invoke-CategorySigning -ErrorAction SilentlyContinue
    if (-not $categorySigning) {
        if ($requested) {
            throw 'Script signing is requested but AppPackagerSigning is not loaded; no build may ship unsigned under an enabled signing policy.'
        }
        return @{
            PolicyDigest = ''
            Detection    = @{ Status = 'NotRequested'; File = $null; Reason = 'signing module not loaded' }
            Requirements = @{ Status = 'NotRequested'; Items = @() }
            Deployment   = @{ Status = 'NotRequested'; Files = @(); LaunchersBypassFree = $true }
        }
    }

    # The signing module regenerates scripts\detect*.ps1 and
    # scripts\requirements\*.ps1 from the current manifest. Clearing the
    # folder first stops a previous build's detector, which the current
    # profile no longer selects, from being signed and shipped.
    $scriptsFolder = Join-Path $StageRoot 'scripts'
    if (Test-Path -LiteralPath $scriptsFolder) {
        Remove-Item -LiteralPath $scriptsFolder -Recurse -Force -ErrorAction Stop
    }

    $signing = ConvertTo-WorkbenchHashtable -InputObject (& $categorySigning -StageRoot $StageRoot -ManifestData $ManifestData -Policy $Policy)

    # A failed deployment category is fatal whenever deployment signing is on:
    # the launchers were already generated without an execution-policy
    # argument, so an unsigned payload cannot run on an AllSigned endpoint.
    if ([bool](Get-WorkbenchMember -InputObject $Policy -Path 'SignDeployment').Value) {
        $deploymentStatus = [string](Get-WorkbenchMember -InputObject $signing -Path 'Deployment.Status').Value
        if ($deploymentStatus -eq 'Failed') {
            $reason = [string](Get-WorkbenchMember -InputObject $signing -Path 'Deployment.Reason').Value
            throw "Deployment script signing failed and the launchers carry no execution-policy relaxation: $reason"
        }
    }

    $launcherCheck = Get-Command -Name Test-DeploymentLauncherChain -ErrorAction SilentlyContinue
    if ($launcherCheck) {
        $chain = & $launcherCheck -StageRoot $StageRoot -Manifest $ManifestData -Policy $Policy
        if ($chain) {
            $findings = (Get-WorkbenchMember -InputObject $chain -Path 'Findings').Value
            # ConvertTo-WorkbenchHashtable already returns the array; wrapping
            # it in @() again would nest it one level deeper.
            $signing['LauncherFindings'] = ConvertTo-WorkbenchHashtable -InputObject $findings
            $bypassFree = Get-WorkbenchMember -InputObject $chain -Path 'BypassFree'
            if ($bypassFree.Found) { $signing['LaunchersBypassFree'] = [bool]$bypassFree.Value }
        }
    }
    return $signing
}

function New-ExtendOrchestratorContent {
    <#
    .SYNOPSIS
        Builds the Extend-mode install.ps1 that runs each step as its own
        powershell.exe child.

    .DESCRIPTION
        Separate child processes keep an exit in a generated wrapper from
        skipping the remaining steps. Exit codes propagate unchanged and a
        3010 anywhere in the chain becomes the orchestrator's exit code when
        every step otherwise succeeded. A step is never launched with an
        execution-policy argument when the deployment scripts are signed.
    #>
    param(
        [string]$Before,
        [Parameter(Mandatory)][string]$Generated,
        [string]$After,
        [bool]$AfterRunsOnFailure = $false,
        [bool]$SignedDeployment = $false,
        [string]$ScriptHost = 'x64'
    )

    $hostExpression = if ($ScriptHost -eq 'x86') {
        "Join-Path ([Environment]::GetFolderPath('Windows')) 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'"
    }
    else {
        "'powershell.exe'"
    }
    $policyArgument = if ($SignedDeployment) { '' } else { ", '-ExecutionPolicy', 'Bypass'" }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('$ErrorActionPreference = ' + "'Stop'")
    $lines.Add('$hostExe = ' + $hostExpression)
    $lines.Add('$rebootPending = $false')
    $lines.Add('function Invoke-Step {')
    $lines.Add('    param([string]$Name)')
    $lines.Add('    $script = Join-Path $PSScriptRoot $Name')
    $lines.Add('    if (-not (Test-Path -LiteralPath $script)) { throw "Step script not found: $script" }')
    $lines.Add("    `$stepArgs = @('-NoProfile', '-NonInteractive'$policyArgument, '-File', `$script)")
    $lines.Add('    $proc = Start-Process -FilePath $hostExe -ArgumentList $stepArgs -Wait -PassThru -NoNewWindow')
    $lines.Add('    return [int]$proc.ExitCode')
    $lines.Add('}')
    $lines.Add('')

    $emitStep = {
        param([string]$name, [bool]$isFinalGate)
        $lines.Add("`$code = Invoke-Step -Name '$name'")
        $lines.Add('if ($code -eq 3010) { $rebootPending = $true; $code = 0 }')
        if ($isFinalGate) { return }
        $lines.Add('if ($code -ne 0) { exit $code }')
    }

    if (-not [string]::IsNullOrWhiteSpace($Before)) { & $emitStep $Before $false }

    $lines.Add("`$code = Invoke-Step -Name '$Generated'")
    $lines.Add('if ($code -eq 3010) { $rebootPending = $true; $code = 0 }')
    if (-not [string]::IsNullOrWhiteSpace($After)) {
        if ($AfterRunsOnFailure) {
            $lines.Add('$generatedCode = $code')
            $lines.Add("`$afterCode = Invoke-Step -Name '$After'")
            $lines.Add('if ($afterCode -eq 3010) { $rebootPending = $true; $afterCode = 0 }')
            $lines.Add('if ($generatedCode -ne 0) { exit $generatedCode }')
            $lines.Add('if ($afterCode -ne 0) { exit $afterCode }')
        }
        else {
            $lines.Add('if ($code -ne 0) { exit $code }')
            $lines.Add("`$code = Invoke-Step -Name '$After'")
            $lines.Add('if ($code -eq 3010) { $rebootPending = $true; $code = 0 }')
            $lines.Add('if ($code -ne 0) { exit $code }')
        }
    }
    else {
        $lines.Add('if ($code -ne 0) { exit $code }')
    }
    $lines.Add('if ($rebootPending) { exit 3010 }')
    $lines.Add('exit 0')

    return (($lines -join "`r`n") + "`r`n")
}

function Copy-WorkbenchStageFile {
    <#
    .SYNOPSIS
        Copies one profile file into the stage root with the containment,
        collision and reparse-point checks the build contract requires.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$Destination,
        [string]$ExpectedSha256,
        [switch]$AllowOverwrite
    )

    if ([System.IO.Path]::IsPathRooted($Destination)) {
        throw "Source file destination '$Destination' must be relative to the content root."
    }
    $rootFull = [System.IO.Path]::GetFullPath($StageRoot).TrimEnd('\')
    $targetFull = [System.IO.Path]::GetFullPath((Join-Path $rootFull $Destination))
    if (-not $targetFull.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Source file destination '$Destination' escapes the content root."
    }
    if (-not $AllowOverwrite -and (Test-Path -LiteralPath $targetFull)) {
        # A byte-identical file at this destination is this profile's own copy
        # from an earlier stage of the same version, which the build-record
        # prune misses when no record was written. Anything else is a real
        # collision with generated or third-party content.
        $replaceable = $false
        if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
            $replaceable = ((Get-WorkbenchFileSha256 -Path $targetFull) -eq $ExpectedSha256)
        }
        if (-not $replaceable) {
            throw "Source file destination '$Destination' collides with a file already in the stage root."
        }
    }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "Source file '$SourcePath' is missing; the build cannot snapshot it."
    }
    if (Test-WorkbenchReparsePoint -Path $SourcePath) {
        throw "Source file '$SourcePath' is a reparse point; copy the real file into the profile instead."
    }
    Test-WorkbenchPathChainForReparse -Root $rootFull -Path $targetFull

    $folder = Split-Path -Path $targetFull -Parent
    [void](New-WorkbenchFolder -Path $folder)
    Test-WorkbenchPathChainForReparse -Root $rootFull -Path $folder
    Copy-Item -LiteralPath $SourcePath -Destination $targetFull -Force -ErrorAction Stop
    return $targetFull
}

function Set-WorkbenchRequirementOperations {
    <#
    .SYNOPSIS
        Applies Add/Replace/Remove requirement operations keyed by RuleId.

    .DESCRIPTION
        AppliesTo '*' reaches the base manifest requirements and every
        deployment type entry. A named variant list reaches only the matching
        NameSuffix entries, so an app-wide rule never gates a deliberate
        unconditional fallback by accident.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [Parameter(Mandatory)]$Operations
    )

    $ruleIdOf = {
        param($rule)
        $id = [string](Get-WorkbenchMember -InputObject $rule -Path 'RuleId').Value
        if ([string]::IsNullOrWhiteSpace($id)) { $id = [string](Get-WorkbenchMember -InputObject $rule -Path 'ConditionId').Value }
        return $id
    }

    $apply = {
        param([object[]]$rules, $operation)
        $op = [string](Get-WorkbenchMember -InputObject $operation -Path 'Op').Value
        $ruleId = [string](Get-WorkbenchMember -InputObject $operation -Path 'RuleId').Value
        $rule = ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $operation -Path 'Rule').Value
        if ($rule -is [System.Collections.IDictionary]) { $rule['RuleId'] = $ruleId }
        $kept = @(@($rules) | Where-Object { $null -ne $_ -and (& $ruleIdOf $_) -ne $ruleId })
        switch ($op) {
            'Remove' { return , $kept }
            'Replace' { return , ($kept + , $rule) }
            default {
                if (@($rules).Count -ne $kept.Count) { return , ($kept + , $rule) }
                return , (@(@($rules) | Where-Object { $null -ne $_ }) + , $rule)
            }
        }
    }

    $entries = @()
    if ($ManifestData['DeploymentTypes']) { $entries = @($ManifestData['DeploymentTypes']) }

    foreach ($operation in @($Operations)) {
        if ($null -eq $operation) { continue }
        $appliesTo = @((Get-WorkbenchMember -InputObject $operation -Path 'AppliesTo').Value)
        if ($appliesTo.Count -eq 0) { $appliesTo = @('*') }
        $all = ($appliesTo -contains '*')

        if ($all -or $entries.Count -eq 0) {
            $ManifestData['Requirements'] = & $apply @($ManifestData['Requirements']) $operation
        }
        for ($i = 0; $i -lt $entries.Count; $i++) {
            $entry = $entries[$i]
            $suffix = [string](Get-WorkbenchMember -InputObject $entry -Path 'NameSuffix').Value
            if (-not $all -and ($appliesTo -notcontains $suffix)) { continue }
            $entryData = ConvertTo-WorkbenchHashtable -InputObject $entry
            $entryData['Requirements'] = & $apply @($entryData['Requirements']) $operation
            $entries[$i] = $entryData
        }
    }
    if ($entries.Count -gt 0) { $ManifestData['DeploymentTypes'] = @($entries) }
    return $ManifestData
}

function Resolve-WorkbenchSetupFile {
    <#
    .SYNOPSIS
        Resolves the real Intune setup entry file for a staged build.

    .DESCRIPTION
        The install command's own first file reference wins, so a PSADT tree
        or an edited wrapper reports its real entry instead of the generic
        install.bat assumption. A multi-deployment-type manifest resolves
        against its first entry's content subfolder.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][hashtable]$ManifestData
    )

    $subRoot = $StageRoot
    $command = [string]$ManifestData['InstallCommandLine']
    if ($ManifestData['DeploymentTypes']) {
        $first = @($ManifestData['DeploymentTypes'])[0]
        if ($first) {
            $subpath = [string](Get-WorkbenchMember -InputObject $first -Path 'ContentSubpath').Value
            if (-not [string]::IsNullOrWhiteSpace($subpath)) { $subRoot = Join-Path $StageRoot $subpath }
            $entryCommand = [string](Get-WorkbenchMember -InputObject $first -Path 'InstallCommandLine').Value
            if (-not [string]::IsNullOrWhiteSpace($entryCommand)) { $command = $entryCommand }
        }
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches([string]$command, '"([^"]+)"|(\S+)')) {
        $token = $match.Groups[1].Value
        if ([string]::IsNullOrEmpty($token)) { $token = $match.Groups[2].Value }
        $token = $token.Trim()
        if ([string]::IsNullOrWhiteSpace($token) -or $token.StartsWith('-') -or $token.StartsWith('/')) { continue }
        $token = $token -replace '^\.\\', '' -replace '^%~dp0', ''
        if ($token -match '\.(bat|cmd|exe|ps1|msi)$') { $candidates.Add($token) }
    }
    foreach ($fallback in @('Deploy-Application.exe', 'Deploy-Application.ps1', 'install.bat')) { $candidates.Add($fallback) }

    foreach ($candidate in $candidates) {
        $full = Join-Path $subRoot $candidate
        if (Test-Path -LiteralPath $full -PathType Leaf) {
            $rootFull = [System.IO.Path]::GetFullPath($StageRoot).TrimEnd('\')
            return ([System.IO.Path]::GetFullPath($full).Substring($rootFull.Length + 1))
        }
    }
    return 'install.bat'
}

function Remove-WorkbenchStaleStageArtifacts {
    <#
    .SYNOPSIS
        Deletes the profile-owned files a previous build left in this stage
        root.

    .DESCRIPTION
        A stage folder is reused for the same vendor version, so a source
        file or hook script removed from the profile would otherwise stay on
        disk, be hashed into FileHashes and ship. Only the categories the
        profile owns are pruned; the packager's own generated wrappers and
        payload are never touched.

    .OUTPUTS
        [string[]] the relative paths removed.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$BuildId,
        [string]$DataRoot
    )

    $prunable = @('SourceFile', 'InstallBefore', 'InstallAfter', 'UninstallBefore', 'UninstallAfter')
    $rootFull = [System.IO.Path]::GetFullPath($StageRoot).TrimEnd('\')
    $removed = New-Object System.Collections.ArrayList
    $folders = New-Object System.Collections.ArrayList

    $records = @(Get-BuildRecords -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot |
        Where-Object { [string]$_.BuildId -ne [string]$BuildId })
    foreach ($record in $records) {
        $recordRoot = [string]$record.StageRoot
        if ([string]::IsNullOrWhiteSpace($recordRoot)) { continue }
        if (([System.IO.Path]::GetFullPath($recordRoot).TrimEnd('\')) -ine $rootFull) { continue }
        foreach ($asset in @($record.CustomAssets)) {
            if ($null -eq $asset) { continue }
            $category = [string](Get-WorkbenchMember -InputObject $asset -Path 'Category').Value
            if ($prunable -notcontains $category) { continue }
            $relative = [string](Get-WorkbenchMember -InputObject $asset -Path 'RelativePath').Value
            if ([string]::IsNullOrWhiteSpace($relative)) { continue }
            $full = [System.IO.Path]::GetFullPath((Join-Path $rootFull $relative))
            if (-not $full.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if (Test-Path -LiteralPath $full -PathType Leaf) {
                Remove-Item -LiteralPath $full -Force -ErrorAction Stop
                [void]$removed.Add($relative)
            }
            $parent = Split-Path -Path $full -Parent
            if ($parent -and ($parent.TrimEnd('\').Length -gt $rootFull.Length)) { [void]$folders.Add($parent) }
        }
    }

    foreach ($folder in @($folders | Select-Object -Unique | Sort-Object -Property Length -Descending)) {
        if (-not (Test-Path -LiteralPath $folder)) { continue }
        if (@(Get-ChildItem -LiteralPath $folder -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item -LiteralPath $folder -Force -ErrorAction SilentlyContinue
        }
    }
    return @($removed)
}

function Invoke-StageFinalization {
    <#
    .SYNOPSIS
        Applies the run snapshot's profile to a staged build before hashing.

    .DESCRIPTION
        Called by Write-StageManifest after the install mode and icon land
        and before the file hashes are computed. Without a run snapshot it
        stamps the schema-4 fields with their defaults and leaves the rest of
        the manifest untouched, so an unprofiled build stays byte-identical
        to the legacy path.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)][hashtable]$ManifestData,
        [AllowNull()][string]$PackagerScriptPath,
        [string]$SnapshotPath
    )

    [void](Set-WorkbenchSchema4Defaults -ManifestData $ManifestData)

    $snapshot = Get-RunSnapshot -Path $SnapshotPath
    $policy = Get-WorkbenchSigningPolicy -Snapshot $snapshot

    if (-not $snapshot) {
        $ManifestData['SetupFile'] = Resolve-WorkbenchSetupFile -StageRoot $StageRoot -ManifestData $ManifestData
        $ManifestData['ScriptSigning'] = Invoke-WorkbenchSigning -StageRoot $StageRoot -ManifestData $ManifestData -Policy $policy
        return $ManifestData
    }

    $profile = $snapshot.Profile
    $applicationId = [string]$snapshot.ApplicationId
    $profileId = [string]$snapshot.ProfileId
    $dataRoot = [string]$snapshot.DataRoot
    $runOverrides = $snapshot.RunOverrides
    $customAssets = New-Object System.Collections.ArrayList

    $ManifestData['BuildId'] = [string]$snapshot.BuildId
    $ManifestData['ApplicationId'] = $applicationId
    $ManifestData['ProfileId'] = $profileId
    $ManifestData['ProfileRevision'] = [int]$snapshot.ProfileRevision
    $ManifestData['RunOverrides'] = (ConvertTo-WorkbenchHashtable -InputObject $runOverrides)

    # Prune first: this stage folder is reused for the same vendor version, so
    # every file the previous build's profile owned goes before the new
    # content lands. What the current profile still carries is rewritten
    # below, which also keeps the collision check honest about generated files.
    foreach ($section in @('Install', 'Uninstall')) {
        $block = (Get-WorkbenchMember -InputObject $profile -Path $section).Value
        if ([string](Get-WorkbenchMember -InputObject $block -Path 'Mode').Value -eq 'Extend') { continue }
        $stale = Join-Path $StageRoot ($section.ToLowerInvariant() + '-generated.ps1')
        if (Test-Path -LiteralPath $stale -PathType Leaf) { Remove-Item -LiteralPath $stale -Force -ErrorAction Stop }
    }
    [void](Remove-WorkbenchStaleStageArtifacts -StageRoot $StageRoot -ApplicationId $applicationId -ProfileId $profileId `
            -BuildId ([string]$snapshot.BuildId) -DataRoot $dataRoot)

    $resolveAsset = {
        param([string]$assetId)
        $paths = $snapshot.AssetPaths
        if ($paths) {
            $member = Get-WorkbenchMember -InputObject $paths -Path $assetId
            if ($member.Found -and $member.Value) { return [string]$member.Value }
        }
        return (Resolve-ProfileAssetPath -ApplicationId $applicationId -ProfileId $profileId -Asset @{ Asset = $assetId } -DataRoot $dataRoot)
    }

    $settings = Resolve-EffectiveSettings -BaseManifest $ManifestData -Profile $profile -RunOverrides $runOverrides

    # 2a. Install and uninstall scripts, hooks and commands.
    $signedDeployment = [bool](Get-WorkbenchMember -InputObject $policy -Path 'SignDeployment').Value
    $scriptHost = [string]$settings['ScriptHost'].Value
    if ([string]::IsNullOrWhiteSpace($scriptHost)) { $scriptHost = 'x64' }

    foreach ($section in @('Install', 'Uninstall')) {
        $block = (Get-WorkbenchMember -InputObject $profile -Path $section).Value
        if (-not $block) { continue }
        $mode = [string](Get-WorkbenchMember -InputObject $block -Path 'Mode').Value
        $target = ($section.ToLowerInvariant() + '.ps1')

        if ($mode -eq 'Custom') {
            $assetId = [string](Get-WorkbenchMember -InputObject $block -Path 'Script').Value
            if ([string]::IsNullOrWhiteSpace($assetId)) { throw "$section.Mode Custom requires a script asset." }
            $source = & $resolveAsset $assetId
            $text = Resolve-WorkbenchTokens -Text ([System.IO.File]::ReadAllText($source)) -ManifestData $ManifestData -Context Script
            [System.IO.File]::WriteAllText((Join-Path $StageRoot $target), $text, (New-Object System.Text.UTF8Encoding($false)))
            [void]$customAssets.Add(@{ Category = $section; RelativePath = $target; Asset = $assetId })
        }
        elseif ($mode -eq 'Extend') {
            $generatedPath = Join-Path $StageRoot $target
            if (-not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
                throw "$section.Mode Extend needs the packager's generated $target, which the stage did not produce."
            }
            $generatedName = ($section.ToLowerInvariant() + '-generated.ps1')
            Move-Item -LiteralPath $generatedPath -Destination (Join-Path $StageRoot $generatedName) -Force -ErrorAction Stop

            $stepNames = @{}
            foreach ($hook in @('Before', 'After')) {
                $assetId = [string](Get-WorkbenchMember -InputObject $block -Path $hook).Value
                if ([string]::IsNullOrWhiteSpace($assetId)) { continue }
                $source = & $resolveAsset $assetId
                $name = ('{0}-{1}.ps1' -f $section.ToLowerInvariant(), $hook.ToLowerInvariant())
                $text = Resolve-WorkbenchTokens -Text ([System.IO.File]::ReadAllText($source)) -ManifestData $ManifestData -Context Script
                [System.IO.File]::WriteAllText((Join-Path $StageRoot $name), $text, (New-Object System.Text.UTF8Encoding($false)))
                $stepNames[$hook] = $name
                [void]$customAssets.Add(@{ Category = ($section + $hook); RelativePath = $name; Asset = $assetId })
            }

            $orchestrator = New-ExtendOrchestratorContent -Before ([string]$stepNames['Before']) -Generated $generatedName `
                -After ([string]$stepNames['After']) `
                -AfterRunsOnFailure ([bool](Get-WorkbenchMember -InputObject $block -Path 'AfterRunsOnFailure').Value) `
                -SignedDeployment $signedDeployment -ScriptHost $scriptHost
            [System.IO.File]::WriteAllText((Join-Path $StageRoot $target), $orchestrator, [System.Text.Encoding]::ASCII)
            [void]$customAssets.Add(@{ Category = ($section + 'Orchestrator'); RelativePath = $target; Asset = $null })
        }

        foreach ($field in @('ReturnCodes', 'RebootPolicy', 'WorkingDirectory')) {
            $member = Get-WorkbenchMember -InputObject $block -Path $field
            if ($member.Found) { $ManifestData[($section + $field)] = $member.Value }
        }
    }

    foreach ($pair in @(@{ Field = 'InstallCommand'; Key = 'InstallCommandLine' }, @{ Field = 'UninstallCommand'; Key = 'UninstallCommandLine' })) {
        $resolved = $settings[$pair.Field]
        if ($resolved.Source -in @('Profile', 'Variant', 'Target', 'Run') -and -not [string]::IsNullOrWhiteSpace([string]$resolved.Value)) {
            $ManifestData[$pair.Key] = Resolve-WorkbenchTokens -Text ([string]$resolved.Value) -ManifestData $ManifestData -Context Command
        }
    }

    # 2b. Detection.
    $detectionMode = [string](Get-WorkbenchMember -InputObject $profile -Path 'Detection.Mode').Value
    if ($detectionMode -eq 'Custom') {
        $rule = ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $profile -Path 'Detection.Rule').Value
        if (-not $rule) { throw 'Detection.Mode Custom requires a Detection.Rule.' }
        $ManifestData['Detection'] = $rule
        $ManifestData['DetectionSource'] = 'Custom'
    }
    $conversion = Get-WorkbenchMember -InputObject $profile -Path 'Detection.IntuneScriptConversion'
    if ($conversion.Found -and $null -ne $conversion.Value) {
        if (-not ($ManifestData['Detection'] -is [System.Collections.IDictionary])) {
            $ManifestData['Detection'] = ConvertTo-WorkbenchHashtable -InputObject $ManifestData['Detection']
        }
        if ($ManifestData['Detection'] -is [System.Collections.IDictionary]) {
            $ManifestData['Detection']['IntuneScriptConversion'] = [bool]$conversion.Value
        }
    }

    # 2c. Requirements.
    $operations = (Get-WorkbenchMember -InputObject $profile -Path 'Requirements.Operations').Value
    if ($operations -and @($operations).Count -gt 0) {
        [void](Set-WorkbenchRequirementOperations -ManifestData $ManifestData -Operations $operations)
    }

    # 2d. Variant overrides that map onto deployment type entries.
    $variantOverrides = (Get-WorkbenchMember -InputObject $profile -Path 'Variants.Overrides').Value
    if ($variantOverrides -and $ManifestData['DeploymentTypes']) {
        $entries = @($ManifestData['DeploymentTypes'])
        for ($i = 0; $i -lt $entries.Count; $i++) {
            $suffix = [string](Get-WorkbenchMember -InputObject $entries[$i] -Path 'NameSuffix').Value
            $override = (Get-WorkbenchMember -InputObject $variantOverrides -Path $suffix).Value
            if (-not $override) { continue }
            $entryData = ConvertTo-WorkbenchHashtable -InputObject $entries[$i]
            foreach ($map in @(@{ From = 'InstallCommand'; To = 'InstallCommandLine' }, @{ From = 'UninstallCommand'; To = 'UninstallCommandLine' },
                    @{ From = 'Detection'; To = 'Detection' }, @{ From = 'Context'; To = 'InstallationBehaviorType' },
                    @{ From = 'LogonRequirement'; To = 'LogonRequirementType' })) {
                $member = Get-WorkbenchMember -InputObject $override -Path $map.From
                if (-not $member.Found) { continue }
                $value = $member.Value
                if ($map.To -like '*CommandLine' -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $value = Resolve-WorkbenchTokens -Text ([string]$value) -ManifestData $ManifestData -Context Command
                }
                $entryData[$map.To] = $value
            }
            $entries[$i] = $entryData
        }
        $ManifestData['DeploymentTypes'] = @($entries)
    }

    # 2e. Timing and execution.
    $timing = @{ EstimatedMinutes = $null; MaximumMinutes = $null }
    foreach ($field in @('EstimatedMinutes', 'MaximumMinutes')) {
        $resolved = $settings[$field]
        if ($resolved.Source -in @('Profile', 'Variant', 'Target', 'Run', 'Global') -and $null -ne $resolved.Value) {
            $timing[$field] = [int]$resolved.Value
        }
    }
    if ($null -ne $timing['EstimatedMinutes'] -and $null -ne $timing['MaximumMinutes'] -and $timing['EstimatedMinutes'] -gt $timing['MaximumMinutes']) {
        throw "Estimated runtime ($($timing['EstimatedMinutes'])) cannot exceed maximum runtime ($($timing['MaximumMinutes']))."
    }
    $timing['IntuneApplied'] = $false
    $ManifestData['Timing'] = $timing

    $execution = @{}
    foreach ($field in @('Context', 'LogonRequirement', 'UserInteraction', 'ScriptHost')) {
        $execution[$field] = $settings[$field].Value
    }
    $ManifestData['Execution'] = $execution
    if ($null -ne $execution['Context']) { $ManifestData['InstallationBehaviorType'] = $execution['Context'] }
    if ($null -ne $execution['LogonRequirement']) { $ManifestData['LogonRequirementType'] = $execution['LogonRequirement'] }
    if ($null -ne $execution['UserInteraction']) { $ManifestData['RequireUserInteraction'] = [bool]$execution['UserInteraction'] }

    # 2f. Icon asset override; it replaces whatever Add-StageIcon produced.
    $iconAsset = [string](Get-WorkbenchMember -InputObject $profile -Path 'Application.Icon.Asset').Value
    $iconMember = Get-WorkbenchMember -InputObject $profile -Path 'Application.Icon'
    if (-not [string]::IsNullOrWhiteSpace($iconAsset)) {
        $source = & $resolveAsset $iconAsset
        $extension = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
        if ($extension -notin @('.ico', '.png', '.jpg', '.jpeg')) { throw "Custom icon '$source' is not an .ico, .png or .jpg file." }
        foreach ($stale in @(Get-ChildItem -LiteralPath $StageRoot -Filter 'app-icon.*' -File -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $stale.FullName -Force -ErrorAction SilentlyContinue
        }
        $destination = Join-Path $StageRoot ('app-icon' + $extension)
        Copy-Item -LiteralPath $source -Destination $destination -Force -ErrorAction Stop
        $ManifestData['Icon'] = Split-Path -Leaf $destination
        [void]$customAssets.Add(@{ Category = 'Icon'; RelativePath = $ManifestData['Icon']; Asset = $iconAsset })
    }
    elseif ($iconMember.Found -and $null -eq $iconMember.Value) {
        foreach ($stale in @(Get-ChildItem -LiteralPath $StageRoot -Filter 'app-icon.*' -File -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $stale.FullName -Force -ErrorAction SilentlyContinue
        }
        if ($ManifestData.ContainsKey('Icon')) { [void]$ManifestData.Remove('Icon') }
    }

    # 2g. Title mode from the profile, applied the same way the legacy bridge does.
    $titleMode = [string]$settings['TitleMode'].Value
    if ($titleMode -in @('IncludeVersion', 'NoVersion') -and (Get-Command -Name Get-PackagedApplicationName -ErrorAction SilentlyContinue)) {
        $ManifestData['AppName'] = Get-PackagedApplicationName -AppName ([string]$ManifestData['AppName']) -Version ([string]$ManifestData['SoftwareVersion']) -Mode $titleMode
    }

    # 2h. Source files, after every generated file exists so a collision is real.
    foreach ($entry in @($profile.SourceFiles)) {
        if ($null -eq $entry) { continue }
        $destination = [string](Get-WorkbenchMember -InputObject $entry -Path 'Destination').Value
        $source = Resolve-ProfileAssetPath -ApplicationId $applicationId -ProfileId $profileId -Asset $entry -DataRoot $dataRoot
        $sourceHash = Get-WorkbenchFileSha256 -Path $source
        $written = Copy-WorkbenchStageFile -StageRoot $StageRoot -SourcePath $source -Destination $destination -ExpectedSha256 $sourceHash
        [void]$customAssets.Add(@{
            Category     = 'SourceFile'
            RelativePath = $destination
            Asset        = [string](Get-WorkbenchMember -InputObject $entry -Path 'Asset').Value
            Sha256       = (Get-WorkbenchFileSha256 -Path $written)
            Size         = (Get-Item -LiteralPath $written).Length
            Linked       = [bool](Get-WorkbenchMember -InputObject $entry -Path 'Linked').Value
        })
    }

    # 3. Resolved entry commands and the real setup file.
    $ManifestData['CustomAssets'] = @($customAssets)
    $ManifestData['SetupFile'] = Resolve-WorkbenchSetupFile -StageRoot $StageRoot -ManifestData $ManifestData

    # 4. Signing, last, so nothing edits a signed file afterwards.
    $ManifestData['ScriptSigning'] = Invoke-WorkbenchSigning -StageRoot $StageRoot -ManifestData $ManifestData -Policy $policy

    Write-WorkbenchLog ("Workbench profile applied    : {0} / {1} rev {2}" -f $applicationId, $profileId, $ManifestData['ProfileRevision'])
    return $ManifestData
}

# ---------------------------------------------------------------------------
# Build records
# ---------------------------------------------------------------------------

function Get-WorkbenchBuildFolder {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$BuildId,
        [string]$DataRoot,
        [switch]$NoCreate
    )
    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    $folder = Join-Path (Join-Path (Join-Path $DataRoot 'builds') (ConvertTo-ApplicationKey -ApplicationId $ApplicationId)) $ProfileId
    if ($BuildId) { $folder = Join-Path $folder $BuildId }
    if (-not $NoCreate) { [void](New-WorkbenchFolder -Path $folder) }
    return $folder
}

function Write-BuildRecord {
    <#
    .SYNOPSIS
        Seals one build: writes build.json and copies the signed detection
        and requirement scripts out of the stage root.
    #>
    param(
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter(Mandatory)]$Manifest,
        [string]$DataRoot
    )

    $applicationId = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'ApplicationId').Value
    $profileId = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'ProfileId').Value
    $buildId = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'BuildId').Value
    if ([string]::IsNullOrWhiteSpace($applicationId) -or [string]::IsNullOrWhiteSpace($buildId)) {
        # A default-profile build carries no workbench identity and needs no record.
        return $null
    }
    if ([string]::IsNullOrWhiteSpace($profileId)) { $profileId = 'default' }

    $folder = Get-WorkbenchBuildFolder -ApplicationId $applicationId -ProfileId $profileId -BuildId $buildId -DataRoot $DataRoot
    $scriptsSource = Join-Path $StageRoot 'scripts'
    if (Test-Path -LiteralPath $scriptsSource) {
        Copy-Item -LiteralPath $scriptsSource -Destination (Join-Path $folder 'scripts') -Recurse -Force -ErrorAction Stop
    }

    $record = @{
        SchemaVersion   = $script:WorkbenchManifestSchemaVersion
        BuildId         = $buildId
        ApplicationId   = $applicationId
        ProfileId       = $profileId
        ProfileRevision = [int](Get-WorkbenchMember -InputObject $Manifest -Path 'ProfileRevision').Value
        SoftwareVersion = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'SoftwareVersion').Value
        AppName         = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'AppName').Value
        PlanDigest      = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'PlanDigest').Value
        DetectionSource = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'DetectionSource').Value
        SetupFile       = [string](Get-WorkbenchMember -InputObject $Manifest -Path 'SetupFile').Value
        RunOverrides    = (ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $Manifest -Path 'RunOverrides').Value)
        ScriptSigning   = (ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $Manifest -Path 'ScriptSigning').Value)
        Timing          = (ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $Manifest -Path 'Timing').Value)
        CustomAssets    = (ConvertTo-WorkbenchHashtable -InputObject (Get-WorkbenchMember -InputObject $Manifest -Path 'CustomAssets').Value)
        StageRoot       = $StageRoot
        SealedAt        = (Get-Date -Format 'o')
    }
    $policyDigest = Get-WorkbenchMember -InputObject $record['ScriptSigning'] -Path 'PolicyDigest'
    $record['PolicyDigest'] = $(if ($policyDigest.Found) { [string]$policyDigest.Value } else { '' })

    [void](Write-WorkbenchJsonFile -Path (Join-Path $folder 'build.json') -InputObject $record)
    $record['Path'] = (Join-Path $folder 'build.json')
    return [pscustomobject]$record
}

function Get-BuildRecords {
    <#
    .SYNOPSIS
        Lists sealed build records for one application, newest first.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$ProfileId,
        [string]$DataRoot
    )

    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    $root = Join-Path (Join-Path $DataRoot 'builds') (ConvertTo-ApplicationKey -ApplicationId $ApplicationId)
    if (-not (Test-Path -LiteralPath $root)) { return @() }

    $records = New-Object System.Collections.ArrayList
    $profileFolders = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)
    if ($ProfileId) { $profileFolders = @($profileFolders | Where-Object { $_.Name -eq $ProfileId }) }
    foreach ($profileFolder in $profileFolders) {
        foreach ($buildFolder in @(Get-ChildItem -LiteralPath $profileFolder.FullName -Directory -ErrorAction SilentlyContinue)) {
            $record = Read-WorkbenchJsonFile -Path (Join-Path $buildFolder.FullName 'build.json')
            if (-not $record) { continue }
            $data = ConvertTo-WorkbenchHashtable -InputObject $record
            $data['Path'] = (Join-Path $buildFolder.FullName 'build.json')
            [void]$records.Add([pscustomobject]$data)
        }
    }
    return @($records | Sort-Object -Property BuildId -Descending)
}

function Get-LatestBuildRecord {
    <#
    .SYNOPSIS
        Returns the newest sealed build record, used for One Click freshness.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$ProfileId,
        [string]$DataRoot
    )
    return (@(Get-BuildRecords -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot) | Select-Object -First 1)
}

function Resolve-StageManifestForBuild {
    <#
    .SYNOPSIS
        Finds the stage manifest belonging to an exact BuildId.

    .DESCRIPTION
        Replaces newest-manifest selection: Package consumes the build the
        operator chose, and a manifest whose BuildId does not match is
        refused as stale rather than packaged.
    #>
    param(
        [Parameter(Mandatory)][string]$BuildId,
        [Parameter(Mandatory)][string]$SearchRoot
    )

    if (-not (Test-Path -LiteralPath $SearchRoot)) {
        throw "Stage search root not found: $SearchRoot"
    }
    $manifests = @(Get-ChildItem -LiteralPath $SearchRoot -Filter 'stage-manifest.json' -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending)
    foreach ($manifest in $manifests) {
        $data = Read-WorkbenchJsonFile -Path $manifest.FullName
        if (-not $data) { continue }
        if ([string]$data.BuildId -eq $BuildId) {
            return [pscustomobject]@{ Path = $manifest.FullName; StageRoot = $manifest.DirectoryName; Manifest = $data }
        }
    }
    throw "stale build: no staged manifest under '$SearchRoot' carries BuildId '$BuildId'."
}

# ---------------------------------------------------------------------------
# BYO applications
# ---------------------------------------------------------------------------

function Save-ByoApplication {
    <#
    .SYNOPSIS
        Persists a dropped installer as a durable BYO application.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [string]$DisplayName,
        [string]$Publisher,
        [string]$SoftwareVersion,
        [AllowNull()]$Analysis,
        [string]$DataRoot,
        [string]$ApplicationId
    )

    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) { throw "Installer not found: $InstallerPath" }
    if ([string]::IsNullOrWhiteSpace($ApplicationId)) { $ApplicationId = New-ApplicationId -Kind Byo }

    if (-not $Analysis -and (Get-Command -Name Get-InstallerAnalysis -ErrorAction SilentlyContinue)) {
        $Analysis = Get-InstallerAnalysis -Path $InstallerPath
    }
    if (-not $DisplayName -and $Analysis) { $DisplayName = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'AppName').Value }
    if (-not $Publisher -and $Analysis) { $Publisher = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'Publisher').Value }
    if (-not $SoftwareVersion -and $Analysis) { $SoftwareVersion = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'SoftwareVersion').Value }
    if ([string]::IsNullOrWhiteSpace($DisplayName)) { throw 'A BYO application requires a display name.' }

    $definition = @{
        SchemaVersion   = $script:WorkbenchProfileSchemaVersion
        ApplicationId   = $ApplicationId
        Origin          = 'byo'
        DisplayName     = $DisplayName
        Publisher       = $Publisher
        Description     = $null
        ProviderId      = $null
        UpdatePolicy    = 'Manual'
        ActiveProfileId = 'default'
        Sources         = @()
    }
    [void](Save-ApplicationDefinition -Definition $definition -DataRoot $DataRoot)
    return (Update-ByoSource -ApplicationId $ApplicationId -InstallerPath $InstallerPath -SoftwareVersion $SoftwareVersion -Analysis $Analysis -DataRoot $DataRoot)
}

function Update-ByoSource {
    <#
    .SYNOPSIS
        Adds a new source revision to a BYO application and compares identity.

    .DESCRIPTION
        The installer is copied into managed storage under its own revision
        number, re-analyzed when installer analysis is available, and its
        identity and architecture compared with the previous revision.
        Overrides that the change invalidates are flagged rather than
        silently carried forward.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$InstallerPath,
        [string]$SoftwareVersion,
        [AllowNull()]$Analysis,
        [string]$DataRoot
    )

    if ($ApplicationId -notlike 'byo:*') { throw "Source revisions apply to BYO applications only (got '$ApplicationId')." }
    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) { throw "Installer not found: $InstallerPath" }

    $definitionData = ConvertTo-WorkbenchHashtable -InputObject (Get-ApplicationDefinition -ApplicationId $ApplicationId -DataRoot $DataRoot)
    $sources = @(@($definitionData['Sources']) | Where-Object { $null -ne $_ })
    $previous = $(if ($sources.Count -gt 0) { $sources[-1] } else { $null })
    $revision = $sources.Count + 1

    if (-not $Analysis -and (Get-Command -Name Get-InstallerAnalysis -ErrorAction SilentlyContinue)) {
        $Analysis = Get-InstallerAnalysis -Path $InstallerPath
    }
    if ([string]::IsNullOrWhiteSpace($SoftwareVersion) -and $Analysis) {
        $SoftwareVersion = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'SoftwareVersion').Value
    }

    $folder = New-WorkbenchFolder -Path (Join-Path (Join-Path (Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot) 'sources') ([string]$revision))
    $fileName = Split-Path -Path $InstallerPath -Leaf
    $stored = Join-Path $folder $fileName
    Copy-Item -LiteralPath $InstallerPath -Destination $stored -Force -ErrorAction Stop

    $source = @{
        Revision        = $revision
        FileName        = $fileName
        Path            = $stored
        Sha256          = (Get-WorkbenchFileSha256 -Path $stored)
        Size            = (Get-Item -LiteralPath $stored).Length
        SoftwareVersion = $SoftwareVersion
        InstallerType   = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'InstallerType').Value
        ProductCode     = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'ProductCode').Value
        Architecture    = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'Architecture').Value
        AppName         = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'AppName').Value
        Publisher       = [string](Get-WorkbenchMember -InputObject $Analysis -Path 'Publisher').Value
        AddedAt         = (Get-Date -Format 'o')
    }

    $flags = New-Object System.Collections.ArrayList
    if ($previous) {
        foreach ($field in @('InstallerType', 'Architecture', 'ProductCode', 'AppName', 'Publisher')) {
            $old = [string](Get-WorkbenchMember -InputObject $previous -Path $field).Value
            $new = [string]$source[$field]
            if ($old -ne $new) {
                [void]$flags.Add([pscustomobject]@{
                    Field    = $field
                    Previous = $old
                    Current  = $new
                    Severity = $(if ($field -in @('InstallerType', 'Architecture')) { 'Blocking' } else { 'Review' })
                })
            }
        }
    }
    $source['IncompatibleOverrides'] = @($flags | Where-Object { $_.Severity -eq 'Blocking' } | ForEach-Object { $_.Field })

    $definitionData['Sources'] = @($sources + , $source)
    if ($SoftwareVersion) { $definitionData['SoftwareVersion'] = $SoftwareVersion }
    [void](Save-ApplicationDefinition -Definition $definitionData -DataRoot $DataRoot)

    return [pscustomobject]@{
        ApplicationId         = $ApplicationId
        Revision              = $revision
        Source                = [pscustomobject]$source
        IdentityChanges       = @($flags)
        IncompatibleOverrides = @($source['IncompatibleOverrides'])
    }
}

# ---------------------------------------------------------------------------
# Drafts
# ---------------------------------------------------------------------------

function Get-WorkbenchDraftPath {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$DataRoot
    )
    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    if ($ProfileId -notmatch '^[A-Za-z0-9._-]+$') { throw "Invalid ProfileId '$ProfileId'." }
    $folder = Join-Path (Join-Path $DataRoot 'drafts') (ConvertTo-ApplicationKey -ApplicationId $ApplicationId)
    return (Join-Path $folder "$ProfileId.draft.json")
}

function Save-Draft {
    <#
    .SYNOPSIS
        Stores an unsaved editor state so it survives a crash.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [Parameter(Mandatory)]$Draft,
        [string]$DataRoot
    )
    $data = ConvertTo-WorkbenchHashtable -InputObject $Draft
    $data['ApplicationId'] = $ApplicationId
    $data['ProfileId'] = $ProfileId
    $data['SavedAt'] = (Get-Date -Format 'o')
    $path = Get-WorkbenchDraftPath -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot
    [void](Write-WorkbenchJsonFile -Path $path -InputObject $data)
    $data['Path'] = $path
    return [pscustomobject]$data
}

function Get-Draft {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$DataRoot
    )
    $path = Get-WorkbenchDraftPath -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot
    $stored = Read-WorkbenchJsonFile -Path $path
    if (-not $stored) { return $null }
    $data = ConvertTo-WorkbenchHashtable -InputObject $stored
    $data['Path'] = $path
    return [pscustomobject]$data
}

function Remove-Draft {
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$ProfileId,
        [string]$DataRoot
    )
    $path = Get-WorkbenchDraftPath -ApplicationId $ApplicationId -ProfileId $ProfileId -DataRoot $DataRoot
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
    return $true
}

# ---------------------------------------------------------------------------
# Portable bundles
# ---------------------------------------------------------------------------

function Add-WorkbenchZipAssembly {
    if (-not ([System.Management.Automation.PSTypeName]'System.IO.Compression.ZipFile').Type) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    }
}

function Export-WorkbenchBundle {
    <#
    .SYNOPSIS
        Exports an application definition, its profiles and assets as a zip.

    .DESCRIPTION
        The bundle carries a schema version and a SHA-256 for every entry.
        Tenant credentials and machine-local connection settings are never
        part of a bundle; BYO installer binaries are included only with
        -IncludeSources and the manifest states when they are absent.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$Path,
        [string]$DataRoot,
        [switch]$IncludeSources
    )

    Add-WorkbenchZipAssembly
    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    $applicationFolder = Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot -NoCreate
    if (-not (Test-Path -LiteralPath $applicationFolder)) { throw "Application '$ApplicationId' has nothing stored to export." }

    $staging = New-WorkbenchFolder -Path (Join-Path ([System.IO.Path]::GetTempPath()) ('apwb-' + [guid]::NewGuid().ToString('N')))
    try {
        $payload = New-WorkbenchFolder -Path (Join-Path $staging 'application')
        Copy-Item -Path (Join-Path $applicationFolder '*') -Destination $payload -Recurse -Force -ErrorAction SilentlyContinue
        $sourcesFolder = Join-Path $payload 'sources'
        $sourcesIncluded = $true
        if (-not $IncludeSources -and (Test-Path -LiteralPath $sourcesFolder)) {
            Remove-Item -LiteralPath $sourcesFolder -Recurse -Force -ErrorAction Stop
            $sourcesIncluded = $false
        }

        $entries = @()
        foreach ($file in @(Get-ChildItem -LiteralPath $payload -File -Recurse -ErrorAction SilentlyContinue)) {
            $relative = $file.FullName.Substring($payload.Length).TrimStart('\')
            $entries += @{ RelativePath = $relative; Sha256 = (Get-WorkbenchFileSha256 -Path $file.FullName); Size = $file.Length }
        }

        $manifest = @{
            SchemaVersion   = $script:WorkbenchBundleSchemaVersion
            ApplicationId   = $ApplicationId
            ExportedAt      = (Get-Date -Format 'o')
            SourcesIncluded = $sourcesIncluded
            SourceStatus    = $(if ($sourcesIncluded) { 'Complete' } else { 'IncompleteSource' })
            Entries         = @($entries)
        }
        [void](Write-WorkbenchJsonFile -Path (Join-Path $staging 'bundle.json') -InputObject $manifest)

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force -ErrorAction Stop }
        [void](New-WorkbenchFolder -Path (Split-Path -Path $Path -Parent))
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $Path)
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
    return [pscustomobject]@{ Path = $Path; ApplicationId = $ApplicationId; SourcesIncluded = [bool]$IncludeSources }
}

function Import-WorkbenchBundle {
    <#
    .SYNOPSIS
        Imports a bundle after verifying its schema version and hashes.

    .DESCRIPTION
        Imported scripts are content until a build explicitly uses them; this
        function never runs anything it extracts.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$DataRoot,
        [string]$ApplicationId,
        [switch]$Force
    )

    Add-WorkbenchZipAssembly
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Bundle not found: $Path" }
    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }

    $staging = New-WorkbenchFolder -Path (Join-Path ([System.IO.Path]::GetTempPath()) ('apwb-' + [guid]::NewGuid().ToString('N')))
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $staging)
        $manifest = Read-WorkbenchJsonFile -Path (Join-Path $staging 'bundle.json')
        if (-not $manifest) { throw "Bundle '$Path' has no bundle.json." }
        if ([int]$manifest.SchemaVersion -gt $script:WorkbenchBundleSchemaVersion) {
            throw "Bundle schema $($manifest.SchemaVersion) is newer than this build supports ($script:WorkbenchBundleSchemaVersion)."
        }
        $payload = Join-Path $staging 'application'
        foreach ($entry in @($manifest.Entries)) {
            if ($null -eq $entry) { continue }
            $file = Join-Path $payload ([string]$entry.RelativePath)
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Bundle entry '$($entry.RelativePath)' is missing from the archive." }
            $actual = Get-WorkbenchFileSha256 -Path $file
            if ($actual -ne [string]$entry.Sha256) { throw "Bundle entry '$($entry.RelativePath)' failed its hash check." }
        }

        if ([string]::IsNullOrWhiteSpace($ApplicationId)) { $ApplicationId = [string]$manifest.ApplicationId }
        $target = Get-WorkbenchApplicationFolder -ApplicationId $ApplicationId -DataRoot $DataRoot -NoCreate
        if ((Test-Path -LiteralPath $target) -and -not $Force) {
            throw "Application '$ApplicationId' already exists in this data root; re-run with -Force to replace it."
        }
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop }
        [void](New-WorkbenchFolder -Path $target)
        Copy-Item -Path (Join-Path $payload '*') -Destination $target -Recurse -Force -ErrorAction SilentlyContinue

        $definitionPath = Join-Path $target 'application.json'
        if (Test-Path -LiteralPath $definitionPath) {
            $definition = ConvertTo-WorkbenchHashtable -InputObject (Read-WorkbenchJsonFile -Path $definitionPath)
            $definition['ApplicationId'] = $ApplicationId
            [void](Write-WorkbenchJsonFile -Path $definitionPath -InputObject $definition)
        }

        return [pscustomobject]@{
            ApplicationId   = $ApplicationId
            EntryCount      = @($manifest.Entries).Count
            SourcesIncluded = [bool]$manifest.SourcesIncluded
            SourceStatus    = [string]$manifest.SourceStatus
        }
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Publications
# ---------------------------------------------------------------------------

function Write-PublicationRecord {
    <#
    .SYNOPSIS
        Appends one publication result to the monthly JSON lines log.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$BuildId,
        [Parameter(Mandatory)][ValidateSet('ContentOnly', 'MECM', 'Intune')][string]$Target,
        [string]$ProfileId = 'default',
        [string]$RemoteAppId,
        [string]$SiteOrTenant,
        [ValidateSet('Succeeded', 'Failed', 'Skipped')][string]$Result = 'Succeeded',
        [string]$Message,
        [string]$DataRoot
    )

    if ([string]::IsNullOrWhiteSpace($DataRoot)) { $DataRoot = Get-WorkbenchDataRoot }
    $folder = New-WorkbenchFolder -Path (Join-Path (Join-Path $DataRoot 'publications') (ConvertTo-ApplicationKey -ApplicationId $ApplicationId))
    $path = Join-Path $folder ((Get-Date -Format 'yyyyMM') + '.jsonl')

    $record = [ordered]@{
        Timestamp     = (Get-Date -Format 'o')
        ApplicationId = $ApplicationId
        ProfileId     = $ProfileId
        BuildId       = $BuildId
        Target        = $Target
        RemoteAppId   = $RemoteAppId
        SiteOrTenant  = $SiteOrTenant
        Result        = $Result
        Message       = $Message
    }
    $line = ($record | ConvertTo-Json -Depth 4 -Compress)
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::AppendAllText($path, $line + "`r`n", $encoding)
    return [pscustomobject]@{ Path = $path; Record = [pscustomobject]$record }
}

Export-ModuleMember -Function @(
    'Get-WorkbenchDataRoot'
    'Get-SharedDownloadCacheRoot'
    'New-ApplicationId'
    'ConvertTo-ApplicationKey'
    'ConvertFrom-ApplicationKey'
    'Get-WorkbenchApplications'
    'Get-ApplicationDefinition'
    'Save-ApplicationDefinition'
    'New-WorkbenchProfileObject'
    'Get-Profiles'
    'Get-Profile'
    'Save-Profile'
    'Copy-Profile'
    'Remove-ProfileField'
    'Add-ProfileAsset'
    'Remove-ProfileAsset'
    'Resolve-ProfileAssetPath'
    'Resolve-EffectiveSettings'
    'Invoke-LegacyPreferenceMigration'
    'New-BuildId'
    'New-RunSnapshot'
    'Get-RunSnapshot'
    'Resolve-WorkbenchTokens'
    'New-ExtendOrchestratorContent'
    'Invoke-StageFinalization'
    'Write-BuildRecord'
    'Get-BuildRecords'
    'Get-LatestBuildRecord'
    'Resolve-StageManifestForBuild'
    'Save-ByoApplication'
    'Update-ByoSource'
    'Save-Draft'
    'Get-Draft'
    'Remove-Draft'
    'Export-WorkbenchBundle'
    'Import-WorkbenchBundle'
    'Write-PublicationRecord'
)
